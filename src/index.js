(function() {
    'use strict';
    const testing = require('aw-config').testing;
    const baseURL = testing ? 'http://localhost:5666' : 'http://localhost:5600';
    const reTime = /^(?:(?<h>[0-9]+)h\s*)?(?:(?<m>[0-9]+)m\s*)?(?:(?<s>[0-9]+)s\s*)?$/;
    const reName = /^(?<s>[A-Z]+\s*)(?<f>[A-Za-z]+\s*)$/;
    const reDate = /^(?<title>.*), (?<start>.* [AP]M) to (?<end>.* [AP]M), (?:(?<location>location: .*), )?(?<info>.*)$/
    function getText(item, selector) {
        var spanNode = item.querySelector(selector);
        return spanNode ? spanNode.innerText : null;
    };
    function getTitle(item, selector) {
        var spanNode = item.querySelector(selector);
        return spanNode ? spanNode.getAttribute('title') : null;
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
    function createBucket(bucketName) {
        const data = {client: 'aw-teams-history-plugin', type: 'app.comms.activity', hostname: ''};
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
    function durationToSeconds(durationText) {
        const d = durationText.match(reTime).groups;
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
        if (callType == 'Outgoing') return `${callType} call to ${callerName}`;
        if (callType == 'Incoming') return `${callType} call from ${callerName}`;
        if (callType == 'Missed call') return `${callType} from ${callerName}`;
        else return `${callType} call with ${callerName}`;
    };
    function textToCall(textParams) {
        const {callType, callLength, callDate} = textParams;
        const displayName = normalizeName(textParams.displayName);
        return {
            timestamp: parseDate(callDate).toISOString(),
            duration: durationToSeconds(callLength),
            data: {
                caller: displayName,
                title: describeCall(displayName, callType)
            }
        }
    };
    function detectCalls() {
        window.location.href = "https://teams.microsoft.com/_#/calls/all-calls";
        console.log("Navigating to calls screen...")
        function gatherCalls(resolve, reject) {
            console.log("Looking for call items...");
            var items = document.querySelectorAll("all-call-list div.td-call-list-container div.item-row");
            var calls = [];
            if (items.length >= 1) {
                console.log("Retrieving calls...");
                for (var item of items) {
                    var displayName = getText(item, "div.display-name > span");
                    var callType = getText(item, "div.call-type > span");
                    var callLength = getText(item, "div.length > span");
                    var callDate = getTitle(item, "div.date > span");
                    calls.push(textToCall({displayName, callType, callLength, callDate}));
                }
                console.log("Found calls to send to ActivityWatch:", calls);
                console.log("Completed call detection");
                resolve(calls);
            } else {
                setTimeout(() => {
                    gatherCalls(resolve, reject);
                }, 200);
            }
        }
        return new Promise((resolve, reject) => {
            gatherCalls(resolve, reject);
        });
    }
    function parseDate(dateText) {
        // parses a date with no year and gets the closest actual date
        const now = new Date();
        const year = now.getFullYear();
        const date = new Date(dateText);
        date.setFullYear(year);
        if (now - date < -180*24*3600*1000) {
            date.setFullYear(year-1);
        }
        return date
    };
    function parseMeetingDescription(description) {
        const d = description.match(reDate);
        if (d && d.groups) {
            const now = new Date();
            const year = now.getFullYear();
            const start = parseDate(d.groups.start), end = parseDate(d.groups.end);
            const duration = end - start;
            if (end >= new Date()) {
                // ignore dates past the present, you probably haven't had those meetings yet
                return null;
            }
            return {
                timestamp: start.toISOString(),
                duration: duration/1000,
                data: {
                    title: d.groups.title,
                    info: d.groups.info,
                }
            }
        } else {
            console.warn(`Could not parse event ${description}`);
            return null;
        }
    };
    function gatherEventIds() {
        return new Promise((resolve, reject) => {
            console.log("Identifying calendar indexedDb name");
            const calendarDbName = JSON.parse(localStorage.getItem('ts.indexDbs')).filter(({name}) => (name.startsWith("skypexspaces-calendar-")))[0].name;
            if (!calendarDbName) {
                console.warn("Could not find calendar db; will not resolve calendar IDs");
                resolve({});
                return;
            }
            console.log(`Found calendar db ${calendarDbName}; opening...`);
            const calendarDb = window.indexedDB.open(calendarDbName);
            let calendarEventsMap = {};
            calendarDb.onsuccess = function () {
                console.log("Querying for events...");
                const calendarObjectStore = calendarDb.result.transaction("CalendarEvents").objectStore("CalendarEvents");
                let weeksHandled = [], weeksToHandle = [];
                const handleCalendarWeek = function(calendarEventsQuery, calendarWeek) {
                    return function () {
                        console.log(`Processing events for IDs in week ${calendarWeek}...`);
                        const eventsCache = calendarEventsQuery.result ? calendarEventsQuery.result.data.calendarEventsCacheV2 : null;
                        if (eventsCache === null) {
                            console.warn("Could not query calendar events");
                            resolve({});
                            return;
                        }
                        const groupList = [...eventsCache.allDayEvents, ...eventsCache.inDayEvents, ...eventsCache.recurrenceEvents];
                        for (var g = 0; g < groupList.length; g++) {
                            const eventsGroup = groupList[g];
                            for (var e = 0; e < eventsGroup.events.length; e++) {
                                const cacheEvent = eventsGroup.events[e];
                                const objectId = cacheEvent.objectId;
                                const eventData = cacheEvent.skypeTeamsDataObj;
                                if (eventData && eventData.cid && objectId) {
                                    console.log(`Found objectId for ${cacheEvent.subject} at ${cacheEvent.startTime}`);
                                    calendarEventsMap[objectId] = eventData.cid;
                                } else {
                                    console.log(`No objectId for ${cacheEvent.subject} at ${cacheEvent.startTime}`);
                                }
                            }
                        }
                        console.log(`Mapped ${calendarEventsMap.length} event ids for week ${calendarWeek}...`);
                        weeksHandled.push(calendarWeek);
                        if (weeksHandled.length === weeksToHandle.length) {
                            resolve(calendarEventsMap);
                        }
                    }
                }
                const calendarKeysQuery = calendarObjectStore.getAllKeys();
                calendarKeysQuery.onsuccess = function() {
                    console.log(`Found the following keys on calendar db: ${calendarKeysQuery.result}`);
                    for (let key of calendarKeysQuery.result) {
                        if (!key.startsWith('W')) continue;
                        weeksToHandle.push(key);
                    }
                    for (let week of weeksToHandle) {
                        const calendarEventsQuery = calendarObjectStore.get(week);
                        calendarEventsQuery.onsuccess = handleCalendarWeek(calendarEventsQuery, week);
                    }
                }
            }
        });
    }
    function gatherMeetings(promise) {
        function doGather(resolve, reject) {
            console.log("Screen scraping events...");
            var eventCards = document.querySelectorAll("div[aria-label='Calendar grid view'] div[aria-label][class*='components-calendar-event-card']");
            let events = [];
            if (eventCards.length >= 1) {
                console.log("Processing events...");
                for (var i = 0; i < eventCards.length; i++) {
                    const eventCard = eventCards[i];
                    const eventId = eventCard.getAttribute('data-tid');
                    const description = eventCard.getAttribute('aria-label');
                    let event = parseMeetingDescription(description);
                    if (event !== null) {
                        event.objectId = eventId;
                        events.push(event);
                    }
                }
                resolve(events);
            } else {
                console.log("Let's try again...");
                setTimeout(( )=> { console.log("Timeout reached"); doGather(resolve, reject); }, 200);
            }
        };
        return new Promise((resolve, reject) => { doGather(resolve, reject); });
    }
    function detectMeetings() {
        console.log("Navigating to Calendar page...");
        window.location.href = "https://teams.microsoft.com/_#/calendarv2";
        return new Promise((resolve, reject) => {
            Promise.all([gatherEventIds(), gatherMeetings()]).then(([eventIdMap, calendarEvents]) => {
                let activityWatchEvents = [];
                for (var i = 0; i < calendarEvents.length; i++) {
                    let event = calendarEvents[i];
                    if (event.objectId) {
                        const cid = eventIdMap[event.objectId];
                        if (cid) {
                            event.data.url = `https://teams.microsoft.com/_#/conversations/${cid}?ctx=chat`
                        }
                        delete event.objectId;
                    }
                    activityWatchEvents.push(event);
                }
                console.log("Found Teams events to send to ActivityWatch:", activityWatchEvents);
                console.log("Completed processing events");
                resolve(activityWatchEvents);
            });
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
    function teamsToActivityWatch() {
        console.log("Collecting information from Teams")
        let calls = [], meetings = [];
        return detectCalls().then(detectedCalls => {
            calls.push(...detectedCalls);
            return detectMeetings();
        }).then(detectedMeetings => {
            meetings.push(...detectedMeetings);
            let teamsEvents = [...calls, ...meetings];
            teamsEvents.sort((a, b) => {
                const startDiff = Date.parse(a.timestamp) - Date.parse(b.timestamp);
                if (startDiff !== 0) return startDiff;
                return a.duration - b.duration;
            });
            console.log("Combined events", teamsEvents);
            removeOverlaps(teamsEvents);
            console.log(`Sending combined data to teams (${teamsEvents.length} events)`);
            createBucket("aw-watcher-teams");
            return postEvents("aw-watcher-teams", teamsEvents);
        })
    }
    function whenTeamsLoads() {
        console.log("Waiting for Teams to load...")
        function waitForTeamsLoad(resolve, reject) {
            var items = document.querySelectorAll("div.teams-title");
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
    window.addEventListener('load', function() {
        if (window.location.href.contains('activity-watch-plugin')) {
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
})();