#############################################
# List my next calender events from lokal Outlook (no MS Graph API needed)
# Martin Löffler 
# 02.11.2023
# WORKING (but without regular meetings)
#############################################

Function Get-OutlookCalendar  
 {  
  # use Outlook interop to get the default calendar folder
  # Search your PC in file explorer for "Microsoft.Office.Interop.Outlook.dll" to find one of this DLLs and paste the path here
  Add-Type -Path "C:\Program Files\Microsoft Office\root\Office16\ADDINS\Microsoft Power Query for Excel Integrated\bin\Microsoft.Office.Interop.Outlook.dll" | out-null
  $olFolders = “Microsoft.Office.Interop.Outlook.OlDefaultFolders” -as [type]  
  $outlook = new-object -comobject outlook.application  
  $namespace = $outlook.GetNameSpace(“MAPI”)  
  $folder = $namespace.getDefaultFolder($olFolders::olFolderCalendar) 
  
  # get current time and tomorrow to filter the calendar items 
  $now = Get-Date
  $tomorrow = (Get-Date).AddDays(1).Date

  # get all calendar items for today which are not ended ye, sort them by start time and select the properties we want to display
  $folder.items | Where-Object { ($_.End -gt $now) -and ($_.Start -lt $tomorrow) } | Sort-Object -Property Start | Select-Object -Property Subject, Start, Duration, Location | Format-Table -AutoSize 

 } #end function Get-OutlookCalendar  
 
 # call function
 Clear-Host
 Get-OutlookCalendar

 # refresh every 10 minutes
 while ($true) {
  Clear-Host
  Get-OutlookCalendar
  Start-Sleep -Seconds 600
}
