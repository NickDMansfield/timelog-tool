# timelog-tool

This is a simple tool to allow easier logging of time and hopefully more accurate detailed timelogs.  
It presently requires:
A. The specific XLSX timelog file (or one with identical cell structure/formatting). This is provided in both a 5-day and 15-day format.
B. A settings file (provided, but you need to fill in the data.
C. A Harvest account, and the appropriate keys and and API token.


Got it? Great! To use this, you install it via your usual node package solution
> npm install timelog-bot

To run it. make sure you configure the settings file and your remote permissions first.  When that is done, you run it like so

> timelogbot -J 'settings.json' -S 'test22.xlsx'

Note that this was primarily tested in a Windows environment using a Git Bash terminal.
