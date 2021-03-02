(function() {
    'use strict';
    const testing = true;
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
            timestamp: new Date(callDate).toISOString(),
            duration: durationToSeconds(callLength),
            data: {
                caller: displayName,
                title: describeCall(displayName, callType)
            }
        }
    };
    function detectCalls() {
        window.location.href = "https://teams.microsoft.com/_#/calls/all-calls";
        var items = document.querySelectorAll("all-call-list div.td-call-list-container div.item-row");
        var calls = [];
        if (items.length >= 1) {
            for (var item of items) {
                var displayName = getText(item, "div.display-name > span");
                var callType = getText(item, "div.call-type > span");
                var callLength = getText(item, "div.length > span");
                var callDate = getTitle(item, "div.date > span");
                calls.push(textToCall({displayName, callType, callLength, callDate}));
            }
            console.log("Found", calls);
            createBucket("aw-watcher-teams");
            postEvents("aw-watcher-teams", calls);
            detectMeetings();
        } else {
            setTimeout(detectCalls, 200);
        }
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
    function detectMeetings() {
        window.location.href = "https://teams.microsoft.com/_#/calendarv2";
        var eventCards = document.querySelectorAll("div[aria-label='Calendar grid view'] div[aria-label][class*='components-calendar-event-card']");
        var events = [];
        if (eventCards.length >= 1) {
            for (var i = 0; i < eventCards.length; i++) {
                var eventCard = eventCards[i];
                var description = eventCard.getAttribute('aria-label');
                const event = parseMeetingDescription(description);
                if (event != null) {
                    events.push(event)
                }
            }
            console.log("Found", events);
            createBucket("aw-watcher-teams");
            postEvents("aw-watcher-teams", events);
        } else {
            setTimeout(detectMeetings, 200);
        }
    };
    detectCalls();
})();