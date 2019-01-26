# O2G Calendar Sync
An Outlook add-in to sync Google calendars with Outlook calendars.

The first release of the O2G Calendar Sync add-in is available for use. It comes with the ability to sync multiple calendars either by syncing the full calendar or a date range. You can also schedule calendar syncs as you see fit. **Any calendar you add during the inital setup will be scheduled to automatically update.** As changes are made to Outlook or Google the calendar will automatically sync the established calendar pairs. If you would like to download it just click the Release tab at the top of the page and you can either download by Source or a Setup file.

## Current Version
> Version 1.1.0.3
> Download from [Releases](https://github.com/gatchmart/O2G-Calendar-Sync/releases)  
> You can download either the sources for a binary for installation in Windows. 
> ### Updates & Fixes
> * Added an exception handler that pushes any exceptions to the BugTracker database so I can fix and track them.
> * Fixed an issue when trying to update Google events that were automatically generated.
> * Fixed an issue that would allow the initial setup dialog to be displayed multiple times after setup.
> * Changed the way calendar items are compared to one another to increase accuracy.

## Previous Versions
> Version 1.1
> Download from [Releases](https://github.com/gatchmart/O2G-Calendar-Sync/releases)  
> You can download either the sources for a binary for installation in Windows.   
> ### Updates & Fixes
> * Updated the Scheduler to increase the amount of time between Google requests if there are not any differences between the calendars.
> * Created the EventHasher that hashs the CalendarEvents to be used as a comparison when checking if calendars contain events.
> * Modified the Archiver to create an event backup at the lunch of Outlook.
> * Bug fixes
---
> Version 1.0  
> Download from [Releases](https://github.com/gatchmart/O2G-Calendar-Sync/releases)  
> You can download either the sources for a binary for installation in Windows.   
> ### Features:
> * Many-to-Many Synchronization Options.
> * Automatic syncing provided by a Scheduler.
> * Custom option for the Scheduler.
> * Manual syncing available.

## Installation
> Installation is simple just extract the O2G Calendar Sync.1.1.zip and run the setup executable.  
> Upon initial use the Initial setup will show up. This will allow you to connect to Google, select the calendars you want to sync and perform the initial sync.

## Issues

> If you have any issues running the plugin please create an [Issue](https://github.com/gatchmart/O2G-Calendar-Sync/issues) for me to troubleshoot and fix. Make sure you provide details of how to recreate the issue. If you have the call stack paste a copy of it in the issue as well.  
> You can find the call stack in the log files located in the %USER%/AppData/Roaming/OutlookGoogleSync/Logs directory.
---
Icons made by [Smashicons](https://www.flaticon.com/authors/smashicons), [Amit Jakhu](https://www.flaticon.com/authors/amit-jakhu), and [Google](https://www.flaticon.com/authors/google) from [www.flaticon.com](https://www.flaticon.com/) is licensed by [CC 3.0 BY](http://creativecommons.org/licenses/by/3.0/).
