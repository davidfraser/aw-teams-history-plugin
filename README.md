# aw-teams-history-plugin
Userscript to retrieve Microsoft Teams history information and feed it to ActivityWatch buckets

[ActivityWatch](https://activitywatch.net) is a bundle of software that tracks computer activity.
It supports watchers that record information about what you do and what happens on your computer.

This is a manually triggered watcher that can extract recent historical information from Microsoft Teams
and report it to the ActivityWatch server.
Note that this is different to most ActivityWatch watchers that observe changes continuously
and report them to ActivityWatch in real time.
The reason for this is that keeping the Microsoft Teams website open can cause Teams to see the user as active. 

It does this as a [Userscript](https://en.wikipedia.org/wiki/Userscript) that runs in the browser
and screen-scrapes information from the Microsoft Teams web interface.
It's been tested with [Tampermonkey](https://www.tampermonkey.net/) as a userscript manager,
but could be adapted to contexts fairly easily.

## Information Captured

Events are created in a bucket called `aw-watcher-teams_<hostname>` for:

* **Calls** logged in your Microsoft Teams Call History, including 1:1 and ad-hoc
  group calls. Missed calls and calls with no duration are skipped (no time spent).
* **Meetings** from your calendar. In the new Teams client the calendar is an
  embedded Outlook web app, so these are read from Outlook's local event cache
  (see [Configuration](#configuration) for which meetings are kept).

Calls are deduplicated against what's already stored, and overlapping calls are
trimmed (shorter / later-joined calls yield to the more accurate ones). Meetings
are **not** trimmed against calls, so a scheduled meeting that overlaps an ad-hoc
call will show both in ActivityWatch.

### Overlapping / conflicting meetings

The plugin does not pick a "winner" between two meetings scheduled at the same
time — every meeting that passes the [filters](#configuration) is recorded. To
resolve a conflict, mark the meeting you did **not** attend as *Tentative* (either
by responding Tentative or setting its show-as to Tentative): with the default
configuration that meeting is excluded, leaving only the one you actually attended.

## Build Instructions

This plugin connects to the ActivityWatch test interface by default.
To connect to the standard interface, set `testing: false` in the config file
(`src/aw-prod-config.js` is used by `npm run build:production` / `npm run clip`).

Run `npm run build` to just package the source into a usable (but compressed) Userscript.

Run `npm run clip` to build but also copy the source files onto the clipboard.
You can then paste the contents into the Tampermonkey dashboard editor.

## Configuration

Options live in `src/aw-dev-config.js` (test server) and `src/aw-prod-config.js`
(standard server):

| Option | Default | Meaning |
| --- | --- | --- |
| `maxDaysBack` | `30` | How far back to scrape call history, in days. |
| `maxCalls` | `200` | Safety cap on the number of calls processed per run. |
| `meetingResponses` | `['Accept', 'Organizer']` | A meeting is kept only if your invitation response is one of these. Possible values: `Accept`, `Organizer`, `Tentative`, `NoResponseReceived`, `Decline`, `Unknown`. |
| `meetingShowAs` | `['Busy']` | A meeting is kept only if its show-as (free/busy) is one of these. Possible values: `Busy`, `Tentative`, `Free`, `OOF` (out-of-office), `WorkingElsewhere`. |

A meeting is recorded only if it is a real, timed, non-cancelled meeting that has
already ended within `maxDaysBack`, **and** its response is in `meetingResponses`
**and** its show-as is in `meetingShowAs`. The defaults keep only meetings you
accepted (or organized) and that marked you Busy.

## Runtime Instructions

* Ensure that your browser has the userscript loaded and enabled.
  The new Teams client embeds the calendar as an Outlook iframe, so the userscript
  also matches `https://outlook.office.com/hosted/*` and must be allowed to run in
  frames (Tampermonkey does this by default).
* Ensure that you have the ActivityWatch server running locally. 
* Log in to the [Microsoft Teams web interface](https://teams.microsoft.com/go#)
* Click on the Tampermonkey extension icon, and under *aw-teams-history-plugin*
  select *Run ActivityWatch Teams History Plugin* (or use the `W` keyboard shortcut).
  
  ![Extension Screenshot](screenshots/executing-userscript.png)
* Navigate to your local [ActivityWatch web interface](http://localhost:5600/#/timeline)
  (or [test interface](http://localhost:5666/#/timeline)) and see your teams activity there.
  
  ![Events Screenshot](screenshots/timeline-events.png)
  
### Bookmark for auto-navigation

If the Userscript is enabled, then navigating to https://teams.microsoft.com/_#/calls/all-calls?activity-watch-plugin=1
will automatically run the plugin and then redirect to the ActivityWatch web interface when done

It is recommended to add this URL as a bookmark to your browser for convenience.
