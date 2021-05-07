# aw-teams-history-plugin
Userscript to retrieve Microsoft Teams history information and feed it to ActivityWatch buckets

[ActivityWatch](https://activitywatch.net) is a bundle of software that tracks computer activity.
It supports watchers that record information about what you do and what happens on your computer.

This is a manually triggered watcher that can extract recent historical information from Microsoft Teams
and report it to the ActivityWatch server.

It does this as a [Userscript](https://en.wikipedia.org/wiki/Userscript) that runs in the browser
and screen-scrapes information from the Microsoft Teams web interface.
It's been tested with [Tampermonkey](https://www.tampermonkey.net/) as a userscript manager,
but could be adapted to contexts fairly easily.

## Information Captured

Events are created in a bucket called `aw-watcher-teams` for:

* Calls logged in your Microsoft Teams Call History. (Note that this will not include group calls)
* Events visible in your Microsoft Teams Calendar.

There is a heuristic algorithm for dealing with overlapping events and calls.

## Build Instructions

This plugin connects to the ActivityWatch test interface by default.
To connect to the standard interface, set `testing = false` near the top of `index.js`.

Run `npm run build` to just package the source into a usable (but compressed) Userscript.

Run `npm run clip` to build but also copy the source files onto the clipboard.
You can then paste the contents into the Tampermonkey dashboard editor.

## Runtime Instructions

* Ensure that your browser has the userscript loaded and enabled
* Ensure
* Log in to the [Microsoft Teams web interface](https://teams.microsoft.com/go#)
* Click on the Tampermonkey extension icon, and under *aw-teams-history-plugin*
  select *Run ActivityWatch Teams History Plugin* (or use the `W` keyboard shortcut)
* Navigate to your local [ActivityWatch web interface](http://localhost:5600/#/timeline)
  (or [test interface](http://localhost:5666/#/timeline)) and see your teams activity there.