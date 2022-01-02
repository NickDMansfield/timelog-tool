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

#Settings docs
    
      "accountName": "Company A",
        - This is purely descriptive
        
      "accountId": "444222,",
        - The Harvest account Id. This can be grabbed from the network tab
        
      "authorization": "Bearer foo",
        - The Harvest API token. You need to generate one and authenticate it for your app.  Harvest has docs for this
        
      "projectId": 1111111,
        - The Harvest Project ID. This can be grabbed from the network tab
        
      "taskId": 22222222,
        - The Harvest Task ID. This can be grabbed from the network tab
        
      "userId": 3333333,
        - Your Harvest UserID. This can be grabbed from the network tab
        
      "url": "https://api.harvestapp.com/api/v2/time_entries",
      
      "type": "harvest",
        - This will be used in the future for variant types
        
      "excludeSheets": [],
        - Ignore. Unimplemented
        
      "excludedTicketNames": [],
        - Excludes all rows where the task title matches a case-insensitive member of this array.
        - Note that you can end the title with a * to make it a wildcard filter.  
          - Right now Wildcard filters are broken, so including a * just cuts off everything after. Just use it like "TES*" or "TES" for now.
          
      "lumpSettings": {
        - These settings determine how work items are grouped. As long as either hours or title is populated, it will group the day's logs into a single entry
        
        "hours": 4,
          - For users who want to round their total daily hours.  The amount you enter here is the rounding point. It will always group and upload the logs in multiples of this size. So if it is 4, and your daily total is >= six hours, it will round it to eight hours.  Anything less and it will go to four hours.  If you would be to set it to one hour, then it would always round up or down to the nearest hour (5.75 -> 6, 8.25 -> 8 etc)
          
        "title": "General Development"
          - The task title given for the grouped entry activity. This is not required for rounded grouping, but if you do not wish to round/lump your time, then a title should be provided so the script knows to group.
