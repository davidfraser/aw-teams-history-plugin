// ==UserScript==
// @name         Microsoft Teams Call History Export
// @namespace    http://tampermonkey.net/
// @version      0.1
// @description  export call history from Microsoft Teams
// @author       You
// @match        https://teams.microsoft.com/*
// @grant        none
// ==/UserScript==

(function() {
    'use strict';
    const testing = true;
    const baseURL = testing ? 'http://localhost:5666' : 'http://localhost:5600';
    const reTime = /^(?:(?<h>[0-9]+)h\s*)?(?:(?<m>[0-9]+)m\s*)?(?:(?<s>[0-9]+)s\s*)?$/;
    const reName = /^(?<s>[A-Z]+\s*)(?<f>[A-Za-z]+\s*)$/;
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
        } else {
            setTimeout(detectCalls, 200);
        }
    }
    window.addEventListener('load', function() {
        detectCalls();
    });
})();