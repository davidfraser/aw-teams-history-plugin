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
    function getText(item, selector) {
        var spanNode = item.querySelector(selector);
        return spanNode ? spanNode.innerText : null;
    }
    function getTitle(item, selector) {
        var spanNode = item.querySelector(selector);
        return spanNode ? spanNode.getAttribute('title') : null;
    }
    function detectCalls() {
        var items = document.querySelectorAll("all-call-list div.td-call-list-container div.item-row");
        if (items.length >= 1) {
            for (var item of items) {
                var displayName = getText(item, "div.display-name > span");
                var callType = getText(item, "div.call-type > span");
                var callLength = getText(item, "div.length > span");
                var callDate = getTitle(item, "div.date > span");
                console.log("Found", callType, "call to/from", displayName, "at", callDate, "for", callLength);
            }
        } else {
            setTimeout(detectCalls, 200);
        }
    }
    window.addEventListener('load', function() {
        detectCalls();
    });
})();