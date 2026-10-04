(function() {
    'use strict';
    const config = require('aw-config');
    const testing = config.testing;
    const maxDaysBack = config.maxDaysBack != null ? config.maxDaysBack : 30;
    const maxCalls = config.maxCalls != null ? config.maxCalls : 200;
    // Allowed invitation responses / show-as statuses for a meeting to be recorded.
    const meetingResponses = config.meetingResponses != null ? config.meetingResponses : ['Accept', 'Organizer'];
    const meetingShowAs = config.meetingShowAs != null ? config.meetingShowAs : ['Busy'];
    const baseURL = testing ? 'http://localhost:5666' : 'http://localhost:5600';
    const reTime = /^(?:(?<h>[0-9]+)h\s*)?(?:(?<m>[0-9]+)m\s*)?(?:(?<s>[0-9]+)s\s*)?$/;
    const reName = /^(?<s>[A-Z]+\s*)(?<f>[A-Za-z]+\s*)$/;
    // First-party Teams app ids, used to click the app-bar buttons of the new
    // Teams client (teams.cloud.microsoft). These ids are the same for every tenant.
    const APP_ID_CALLS = "20c3440d-c67e-4420-9f80-0e50c39693df";
    const APP_ID_CALENDAR = "ef56c0de-36fc-4ef8-b417-3d82ba9d073c";
    function navigateToApp(appId) {
        // The new Teams client is a React SPA with no usable hash routes, so we
        // navigate by clicking the app-bar buttons rather than setting location.href.
        const button = document.querySelector(`button[data-tid='${appId}']`);
        if (button) {
            button.click();
        } else {
            console.warn(`Could not find app-bar button for app ${appId}`);
        }
    }
    function getText(item, selector) {
        var spanNode = item.querySelector(selector);
        return spanNode ? spanNode.innerText : null;
    };
    // GM_xmlhttpRequest
    function Request(url, opt) {
        Object.assign(opt, {
            url: `${baseURL}/${url}`,
            timeout: 2000,
            responseType: 'json'
        })
        return new Promise((resolve, reject) => {
            opt.onabort = opt.onerror = opt.ontimeout = reject
            opt.onload = resolve
            GM_xmlhttpRequest(opt)
        })
    }
    function createBucket(bucketName, hostname) {
        const data = {client: 'aw-teams-history-plugin', type: 'app.comms.activity', hostname: hostname || ''};
        return Request(`api/0/buckets/${bucketName}`, {
            method: "POST",
            headers: {'Content-Type': 'application/json;charset=UTF-8'},
            data: JSON.stringify(data)
        }).then(function(response) { console.log(response); });
    }
    function postEvents(bucketName, events) {
        return Request(`api/0/buckets/${bucketName}/events`, {
            method: "POST",
            headers: {'Content-Type': 'application/json;charset=UTF-8'},
            data: JSON.stringify(events)
        }).then(function(response) { console.log(response); });
    }
    function getHostname() {
        // Watchers attach their events to a device by using the machine hostname in
        // both the bucket id (aw-watcher-teams_<hostname>) and the bucket metadata.
        // A userscript can't read the OS hostname, so ask the AW server for it.
        return Request('api/0/info', { method: "GET" })
            .then(r => (r && r.response && r.response.hostname) || '')
            .catch(() => '');
    }
    function getExistingEvents(bucketName) {
        return Request(`api/0/buckets/${bucketName}/events?limit=-1`, { method: "GET" })
            .then(r => (r && r.response) || [])
            .catch(() => []);
    }
    function eventKey(event) {
        // Normalize timestamp/duration so keys match regardless of the serialization
        // the server returns ("+00:00" vs "Z", 186.0 vs 186).
        const t = new Date(event.timestamp).getTime();
        const d = Math.round(event.duration);
        const title = event.data && event.data.title;
        return `${t}|${d}|${title}`;
    }
    function durationToSeconds(durationText) {
        const m = durationText.match(reTime);
        if (!m) {
            // Unexpected duration format (e.g. tokens out of order); treat as unknown
            // rather than throwing, which would hang the whole run.
            console.warn(`Could not parse call duration '${durationText}'`);
            return 0;
        }
        const d = m.groups;
        const duration = (d.h == null ? 0 : parseInt(d.h) * 60 * 60) +
                         (d.m == null ? 0 : parseInt(d.m) * 60) +
                         (d.s == null ? 0 : parseInt(d.s));
        return duration;
    };
    function normalizeName(callerName) {
        const m = callerName.match(reName);
        if (m && m.groups) {
            return (m.groups.f + ' ' + m.groups.s.charAt(0) + m.groups.s.substr(1).toLowerCase()).trim();
        }
        return callerName;
    };
    function describeCall(callerName, callType) {
        // 1:1 calls use a bare "Incoming"/"Outgoing" type; group calls carry the
        // initiator too, e.g. "Incoming from Eric MANUGUERRA", with callerName
        // holding the participant list ("Eric, Philippe COLLIN"). Key off the
        // leading direction word so both shapes produce a clean title.
        callType = (callType || '').trim();
        const direction = /^incoming\b/i.test(callType) ? 'Incoming'
                        : /^outgoing\b/i.test(callType) ? 'Outgoing'
                        : null;
        const isGroup = direction && callType.toLowerCase() !== direction.toLowerCase();
        if (isGroup) return `${direction} group call with ${callerName}`;
        if (direction === 'Outgoing') return `Outgoing call to ${callerName}`;
        if (direction === 'Incoming') return `Incoming call from ${callerName}`;
        return `${callType} call with ${callerName}`;
    };
    function textToCall(textParams) {
        const {callType, callLength, callDate, exactDate} = textParams;
        // The list view only shows a date (or weekday) for calls older than today,
        // so prefer an exact timestamp pulled from the row's accessibility label.
        const when = exactDate || parseDate(callDate);
        if (!when || isNaN(when.getTime())) {
            console.warn(`Skipping call with unparseable date '${callDate}'`);
            return null;
        }
        const displayName = normalizeName(textParams.displayName);
        return {
            timestamp: when.toISOString(),
            duration: durationToSeconds(callLength),
            data: {
                caller: displayName,
                title: describeCall(displayName, callType)
            }
        }
    };
    let exactTimeMisses = 0;
    function extractExactTimestamp(row) {
        // The list view's text only exposes a weekday ("Wednesday") for recent calls,
        // never a time of day, so it can't give us an exact timestamp. The new Teams
        // client virtualizes the list and keys each row by the call's start time in
        // epoch milliseconds, so read that off the row's React fiber instead.
        let node = row, fiberKey;
        while (node && !(fiberKey = Object.keys(node).find(k => k.startsWith('__reactFiber$')))) {
            node = node.parentElement;
        }
        if (node) {
            let fiber = node[fiberKey], depth = 0;
            while (fiber && depth < 50) {
                const props = fiber.memoizedProps;
                if (props && props.virtualRow && props.virtualRow.key != null) {
                    const ms = Number(props.virtualRow.key);
                    // Guard against a non-timestamp key (e.g. a row index): require a
                    // 13-digit millisecond epoch.
                    if (Number.isFinite(ms) && ms > 1e12) {
                        const d = new Date(ms);
                        if (!isNaN(d.getTime())) return d;
                    }
                    break;
                }
                fiber = fiber.return;
                depth++;
            }
        }
        if (exactTimeMisses < 3) {
            exactTimeMisses++;
            console.warn("No exact call time found on row's React fiber; falling back to date text");
        }
        return null;
    };
    function extractCall(item) {
        const displayName = getText(item, "[data-style-id='call-history-display-name']");
        const callType = getText(item, "[data-style-id='call-history-description']");
        // Missed calls expose no duration (the secondary column is just a date), so
        // they'd land as zero-duration events. This bucket measures call time, not a
        // call log, so skip them.
        if (/missed/i.test(callType || '')) return null;
        // The secondary column concatenates the time/date and the duration,
        // e.g. "6:22 PM 21m 2s". Split off the trailing duration.
        const secondary = (getText(item, "[data-style-id='call-history-secondary-data']") || "").replace(/\s+/g, ' ').trim();
        const durMatch = secondary.match(/(\d+\s*[hms](?:\s*\d+\s*[hms])*)\s*$/);
        const callLength = durMatch ? durMatch[1].trim() : "";
        // No duration shown (declined / no-answer / never-connected) means no call
        // time to record, so skip it like a missed call.
        if (!callLength) return null;
        const callDate = durMatch ? secondary.slice(0, durMatch.index).trim() : secondary;
        const exactDate = extractExactTimestamp(item);
        return textToCall({displayName, callType, callLength, callDate, exactDate});
    };
    function detectCalls(seen) {
        seen = seen || new Set();
        console.log("Navigating to calls screen...");
        navigateToApp(APP_ID_CALLS);
        const cutoff = Date.now() - maxDaysBack * 86400000;
        let attempts = 0;
        function waitForRows(onReady, resolve) {
            const items = document.querySelectorAll("div[data-tid='call-history-row']");
            if (items.length >= 1) {
                onReady();
            } else if (attempts++ < 50) {
                setTimeout(() => waitForRows(onReady, resolve), 200);
            } else {
                console.warn("No call history rows found after waiting; returning no calls");
                resolve([]);
            }
        }
        return new Promise((resolve) => {
            waitForRows(async () => {
                console.log("Retrieving calls...");
                // The list is virtualized, so scroll down in passes to load older
                // calls. A per-call key avoids reprocessing recycled DOM nodes.
                const processed = new Set();
                const calls = [];
                let scrolls = 0;
                const maxScrolls = 100;
                while (true) {
                    const rows = document.querySelectorAll("div[data-tid='call-history-row']");
                    let newThisPass = 0, stop = false;
                    for (const item of rows) {
                        const call = extractCall(item);
                        if (!call) continue;
                        const key = eventKey(call);
                        if (processed.has(key)) continue;
                        processed.add(key);
                        newThisPass++;
                        if (new Date(call.timestamp).getTime() < cutoff) {
                            console.log(`Reached a call older than ${maxDaysBack} days; stopping`);
                            stop = true; break;
                        }
                        if (seen.has(key)) {
                            console.log("Reached a call already stored in ActivityWatch; stopping");
                            stop = true; break;
                        }
                        calls.push(call);
                        if (calls.length >= maxCalls) {
                            console.log(`Reached maxCalls (${maxCalls}); stopping`);
                            stop = true; break;
                        }
                    }
                    if (stop) break;
                    if (newThisPass === 0) {
                        console.log("No more calls loaded; reached end of history");
                        break;
                    }
                    if (scrolls++ >= maxScrolls) {
                        console.warn(`Reached max scroll passes (${maxScrolls}); stopping`);
                        break;
                    }
                    rows[rows.length - 1].scrollIntoView({ block: 'end' });
                    await new Promise(r => setTimeout(r, 400));
                }
                console.log(`Found ${calls.length} calls to send to ActivityWatch:`, calls);
                console.log("Completed call detection");
                resolve(calls);
            }, resolve);
        });
    }
    const weekdayNames = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
    function parseDate(dateText) {
        dateText = (dateText || "").trim();
        if (!dateText) return null;
        const now = new Date();
        const year = now.getFullYear();
        // New Teams call history shows only a time for today's calls, e.g. "6:22 PM".
        if (/^\d{1,2}:\d{2}(\s*[AP]M)?$/i.test(dateText)) {
            return new Date(`${now.toDateString()} ${dateText}`);
        }
        // Calls from the previous day may be prefixed with "Yesterday".
        const yst = dateText.match(/^yesterday\b[,\s]*(.*)$/i);
        if (yst) {
            const y = new Date(now.getTime() - 86400000);
            return new Date(`${y.toDateString()} ${yst[1]}`.trim());
        }
        // Calls within the last week are labelled by weekday name only, with no time
        // (e.g. "Tuesday"). Today's calls show a time instead, so a matching weekday
        // always refers to the most recent *past* occurrence (never today). The time
        // of day isn't exposed in the list, so these land at local midnight.
        const wd = weekdayNames.indexOf(dateText.toLowerCase());
        if (wd >= 0) {
            const d = new Date(now);
            let diff = (now.getDay() - wd + 7) % 7;
            if (diff === 0) diff = 7;
            d.setDate(now.getDate() - diff);
            d.setHours(0, 0, 0, 0);
            return d;
        }
        // Older calls show an absolute date with no time, e.g. "9/18/2026".
        const date = new Date(dateText);
        if (isNaN(date.getTime())) {
            console.warn(`Could not parse call date '${dateText}'`);
            return null;
        }
        // If the text had no explicit year, snap it to the closest actual date.
        if (!/\d{4}/.test(dateText)) {
            date.setFullYear(year);
            if (now - date < -180*24*3600*1000) {
                date.setFullYear(year-1);
            }
        }
        return date
    };
    function meetingFromOwaEvent(e, nowMs, cutoffMs) {
        // Maps one Outlook (OWA) calendar event from the offline cache to an
        // ActivityWatch event, or null if it shouldn't be recorded. We only keep
        // real, timed meetings the user is actually tied to and that have already
        // happened within the lookback window.
        if (!e || e.IsAllDayEvent || !e.IsMeeting) return null;
        if (e.IsCancelled || e.IsSeriesCancelled) return null;
        // Only count meetings the user actually committed to (meetingResponses, e.g.
        // Accept/Organizer) and that occupy them (meetingShowAs, e.g. Busy) — drops
        // Tentative / NoResponse / Decline responses and Free / OOF show-as statuses.
        if (!meetingResponses.includes(e.ResponseType)) return null;
        if (!meetingShowAs.includes(e.FreeBusyType)) return null;
        const start = Date.parse(e.Start), end = Date.parse(e.End);
        if (!(start > 0) || !(end > start)) return null;
        if (end > nowMs) return null;      // not finished yet — you haven't had it
        if (start < cutoffMs) return null; // older than maxDaysBack
        const data = { title: e.Subject || '(no subject)' };
        const cid = e.SkypeTeamsProperties && e.SkypeTeamsProperties.cid;
        if (cid) data.url = `https://teams.microsoft.com/l/meetup-join/${encodeURIComponent(cid)}/0`;
        return { timestamp: new Date(start).toISOString(), duration: (end - start) / 1000, data };
    }
    function openOwaEventsDb() {
        // The embedded Outlook calendar caches events in an indexedDB named
        // "owa-offline-data-<...>" (store "events") on the outlook.office.com origin.
        return new Promise((resolve, reject) => {
            if (!indexedDB.databases) { reject(new Error("indexedDB.databases() unavailable")); return; }
            indexedDB.databases().then((dbs) => {
                const name = (dbs || []).map(d => d.name).find(n => /owa-offline-data/.test(n || ''));
                if (!name) { reject(new Error("OWA offline-data db not found")); return; }
                const req = indexedDB.open(name);
                req.onsuccess = () => resolve(req.result);
                req.onerror = () => reject(req.error);
            }, reject);
        });
    }
    function collectOwaMeetings() {
        const nowMs = Date.now();
        const cutoffMs = nowMs - maxDaysBack * 86400000;
        return openOwaEventsDb().then((db) => new Promise((resolve, reject) => {
            const req = db.transaction("events").objectStore("events").getAll();
            req.onsuccess = () => {
                const rows = req.result || [];
                const meetings = [];
                for (const e of rows) {
                    const m = meetingFromOwaEvent(e, nowMs, cutoffMs);
                    if (m) meetings.push(m);
                }
                console.log(`Mapped ${meetings.length} meetings from ${rows.length} OWA calendar events`);
                resolve(meetings);
            };
            req.onerror = () => reject(req.error);
        }));
    }
    async function owaCalendarToActivityWatch() {
        // Runs inside the Outlook iframe. Collects calendar meetings and posts the
        // new ones straight to ActivityWatch (same bucket as the calls). Returns the
        // number of new meetings sent. Note: meetings are not trimmed against calls
        // for overlaps here (the calls live in a different frame); ActivityWatch shows
        // both, which is acceptable for scheduled vs ad-hoc comms.
        const hostname = await getHostname();
        const bucketName = hostname ? `aw-watcher-teams_${hostname}` : "aw-watcher-teams";
        await createBucket(bucketName, hostname);
        const existing = await getExistingEvents(bucketName);
        const seen = new Set(existing.map(eventKey));
        const meetings = await collectOwaMeetings();
        const fresh = meetings.filter(e => !seen.has(eventKey(e)));
        console.log(`${fresh.length}/${meetings.length} meetings are new`);
        if (fresh.length === 0) return 0;
        await postEvents(bucketName, fresh);
        return fresh.length;
    }
    function findOwaIframe() {
        return [...document.querySelectorAll('iframe')].find(f => /outlook\.office\.com/.test(f.src || ''));
    }
    function detectMeetings() {
        // Top-frame side: the calendar is a cross-origin Outlook iframe we can't read
        // directly, so we open the Calendar app (to load the iframe) and message the
        // plugin instance running inside it, which collects and posts meetings itself.
        console.log("Navigating to Calendar to trigger meeting collection...");
        navigateToApp(APP_ID_CALENDAR);
        return new Promise((resolve) => {
            let settled = false;
            const deadline = Date.now() + 30000;
            const finish = (fn) => { if (settled) return; settled = true; window.removeEventListener('message', onMessage); fn && fn(); resolve(); };
            const onMessage = (ev) => {
                const msg = ev.data;
                if (!msg || msg.source !== 'aw-teams-plugin-owa') return;
                if (!/outlook\.office\.com/.test(ev.origin)) return;
                if (msg.status === 'done') {
                    finish(() => console.log(`Calendar collection reported ${msg.count} new meeting(s)`));
                } else {
                    // Transient failure (e.g. the OWA cache DB wasn't ready yet). Don't
                    // give up — let the poke loop retry until the deadline.
                    console.warn("Calendar collection reported an error, will retry:", msg.message);
                }
            };
            window.addEventListener('message', onMessage);
            // Posting before the iframe's listener exists is simply lost, so retry the
            // trigger until it acks or we time out.
            (function poke() {
                if (settled) return;
                if (Date.now() > deadline) {
                    finish(() => console.warn("Calendar iframe did not respond within 30s; skipping meetings"));
                    return;
                }
                const iframe = findOwaIframe();
                if (iframe && iframe.contentWindow) {
                    iframe.contentWindow.postMessage({ source: 'aw-teams-plugin', cmd: 'collect' }, 'https://outlook.office.com');
                }
                setTimeout(poke, 700);
            })();
        });
    }
    function removeOverlaps(events) {
        // removes overlaps between events, letting calls take priority over meetings (as their times are more accurate)
        let lastEvent = null;
        let lastStart = null;
        let lastEnd = null;
        let lastType = null;
        console.log("Removing overlaps");
        for (var i = 0; i < events.length; i++) {
            let thisEvent = events[i];
            let thisStart = new Date(thisEvent.timestamp).getTime();
            let thisEnd = thisStart + thisEvent.duration * 1000;
            let thisType = thisEvent.data ? (thisEvent.data.caller ? 'call' : 'meeting') : null
            if (lastEnd !== null && lastEvent !== null) {
                if (thisStart < lastEnd) {
                    // we always prioritize calls over meetings, but otherwise the future over the past
                    console.log("Start", thisEvent.timestamp, "is before end", new Date(lastEnd).toISOString());
                    console.log("Comparing", lastEvent, "and", thisEvent);
                    if (thisType === 'meeting' && lastType === 'call') {
                        thisStart = lastEnd;
                        let newThisEvent = {...thisEvent};
                        newThisEvent.timestamp = new Date(thisStart).toISOString();
                        newThisEvent.duration = (thisEnd - thisStart) / 1000;
                        console.log("Last boundary wins; this event should start later at", newThisEvent);
                        events[i] = thisEvent = newThisEvent;
                    } else if (thisType === 'call' && thisEvent.duration < 30) {
                        console.log("This call is too short to care, no change");
                        continue;
                    } else if (thisType === 'meeting' && lastType === 'meeting') {
                        console.log("Assuming ActivityWatch will handle overlapping scheduled meetings");
                    } else if ((thisEnd - lastStart) / 1000 / lastEvent.duration < 0.25) {
                        // this event doesn't go more than 25% of the way into the last meeting; probably joined that meeting late
                        let newLastEvent = {...lastEvent};
                        lastStart = new Date(thisEnd).getTime();
                        newLastEvent.timestamp = new Date(lastStart).toISOString();
                        newLastEvent.duration = (lastEnd - lastStart) / 1000;
                        console.log("This boundary wins, last event should start later at", newLastEvent);
                        // in theory this could trigger a cascade, resort, and reevaluate. But not doing that as it's complex
                        events[i-1] = lastEvent = newLastEvent;
                    } else {
                        let newLastEvent = {...lastEvent};
                        lastEnd = new Date(thisStart).getTime();
                        newLastEvent.duration = (lastEnd - lastStart) / 1000;
                        console.log("This boundary wins, last event should shorten to", newLastEvent);
                        events[i-1] = lastEvent = newLastEvent;
                    }
                }
            }
            [ lastEvent, lastStart, lastEnd, lastType ] = [ thisEvent, thisStart, thisEnd, thisType ]
        }
    }
    async function teamsToActivityWatch() {
        console.log("Collecting information from Teams");

        // Attach the bucket to this device so it shows up alongside the other
        // watchers (aw-watcher-teams_<hostname>) instead of an orphaned bucket.
        const hostname = await getHostname();
        const bucketName = hostname ? `aw-watcher-teams_${hostname}` : "aw-watcher-teams";
        console.log(`Using bucket ${bucketName} (hostname '${hostname}')`);
        await createBucket(bucketName, hostname);

        // Load what's already stored first: it lets call detection stop early once
        // it reaches calls we've already captured, and dedupes what remains.
        const existing = await getExistingEvents(bucketName);
        const seen = new Set(existing.map(eventKey));

        let calls = await detectCalls(seen);
        calls.sort((a, b) => {
            const startDiff = Date.parse(a.timestamp) - Date.parse(b.timestamp);
            if (startDiff !== 0) return startDiff;
            return a.duration - b.duration;
        });
        removeOverlaps(calls);

        // Safety net: early-stop usually handles this, but drop anything already stored.
        const before = calls.length;
        calls = calls.filter(e => !seen.has(eventKey(e)));
        console.log(`Filtered ${before - calls.length} duplicate calls; ${calls.length} new`);
        if (calls.length > 0) {
            console.log(`Sending ${calls.length} call(s) to ${bucketName}`);
            await postEvents(bucketName, calls);
        } else {
            console.log("No new calls to send");
        }

        // Meetings live in the embedded Outlook calendar (a cross-origin iframe), so
        // the plugin instance running inside that iframe collects and posts them. We
        // just trigger it and wait for it to finish.
        try {
            await detectMeetings();
        } catch (e) {
            console.warn("Calendar collection failed:", e);
        }
    }
    function whenTeamsLoads() {
        console.log("Waiting for Teams to load...")
        function waitForTeamsLoad(resolve, reject) {
            var items = document.querySelectorAll("div[data-tid='app-layout-area--main']");
            if (items.length > 0) {
                console.log("Teams has loaded")
                resolve();
            } else {
                console.log("Waiting for Teams to load...")
                setTimeout(() => {
                    waitForTeamsLoad(resolve, reject);
                }, 200);
            }
        }
        return new Promise((resolve, reject) => {
            waitForTeamsLoad(resolve, reject);
        });
    }
    function registerContextMenu() {
        whenTeamsLoads().then(() => {
            console.log("Registering Context Menu")
            GM_registerMenuCommand("Run ActivityWatch Teams History Plugin", function() {
                teamsToActivityWatch().then(() => {
                    console.log("Completed teams watcher update");
                })
            }, "w");
        })
    }
    function registerOwaResponder() {
        // Runs inside the Outlook calendar iframe. Waits for the Teams top-frame to
        // ask us to collect, then scrapes the OWA calendar cache into ActivityWatch
        // and acks back so the top-frame knows we're done.
        console.log("ActivityWatch Teams plugin: Outlook calendar context ready; awaiting collect trigger");
        let running = false;
        window.addEventListener('message', function(ev) {
            const msg = ev.data;
            if (!msg || msg.source !== 'aw-teams-plugin' || msg.cmd !== 'collect') return;
            if (!/^https:\/\/teams\.(microsoft\.com|cloud\.microsoft)$/.test(ev.origin)) {
                console.warn("Ignoring collect trigger from untrusted origin", ev.origin);
                return;
            }
            const reply = (status, extra) => {
                try { ev.source && ev.source.postMessage(Object.assign({ source: 'aw-teams-plugin-owa', status }, extra || {}), ev.origin); } catch (e) {}
            };
            if (running) { console.log("Calendar collection already running"); return; }
            running = true;
            owaCalendarToActivityWatch()
                .then((count) => { console.log(`Calendar collection done (${count} new)`); reply('done', { count }); })
                .catch((err) => { console.warn("Calendar collection failed", err); reply('error', { message: String(err) }); })
                .finally(() => { running = false; });
        });
    }
    const isOutlookContext = window.location.hostname === 'outlook.office.com';
    if (isOutlookContext) {
        registerOwaResponder();
    } else {
        window.addEventListener('load', function() {
            if (window.location.href.includes('activity-watch-plugin')) {
                whenTeamsLoads().then(() => {
                    teamsToActivityWatch().then(() => {
                        console.log("Completed teams watcher update");
                        console.log("Navigating to ActivityWatch interface");
                        window.location.href = `${baseURL}/#/timeline`;
                    });
                });
            } else {
                registerContextMenu();
            }
        });
    }
})();