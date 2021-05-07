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
* Ensure that you have the ActivityWatch server running locally. 
* Log in to the [Microsoft Teams web interface](https://teams.microsoft.com/go#)
* Click on the Tampermonkey extension icon, and under *aw-teams-history-plugin*
  select *Run ActivityWatch Teams History Plugin* (or use the `W` keyboard shortcut)
* Navigate to your local [ActivityWatch web interface](http://localhost:5600/#/timeline)
  (or [test interface](http://localhost:5666/#/timeline)) and see your teams activity there.
  
### Bookmark for auto-navigation

If the Userscript is enabled, then navigating to https://teams.microsoft.com/_#/calls/all-calls?activity-watch-plugin=1
will automatically run the plugin and then redirect to the ActivityWatch web interface when done

It is recommended to add this URL as a bookmark to your browser for convenience.
