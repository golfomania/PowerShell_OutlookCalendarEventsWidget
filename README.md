# PowerShell_OutlookCalendarEventsWidget
List the next events in my Outlook calendar of today as "quick and dirty" Widget on my side screen 

## Status
Working, now also including recurring events like weeklys and the timespans to start and end in minutes

## Example
![Alt text](image.png)

## What i'm not happy with 
Due to the fact that its a console application on every refresh the whole list gets rewritten. 

But having this running on a side display (whats i do) means, that the refreh can be catched by the eye and this can be distracting.


Would be cool to update the script, that its only the backend updating a database like SQLite and then have a frontend showing the data with obly component refresh.