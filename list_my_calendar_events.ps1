#############################################
# List my next calender events from lokal Outlook (no MS Graph API needed)
# Martin Löffler 
# 11.11.2023
# WORKING
#############################################

Function Get-OutlookCalendar  
 {  
  # use Outlook interop to get the default calendar folder
  # Search your PC in file explorer for "Microsoft.Office.Interop.Outlook.dll" to find one of this DLLs on your machine and paste the path here
  Add-Type -Path "C:\Program Files\Microsoft Office\root\Office16\ADDINS\Microsoft Power Query for Excel Integrated\bin\Microsoft.Office.Interop.Outlook.dll" | out-null
  $olFolders = “Microsoft.Office.Interop.Outlook.OlDefaultFolders” -as [type]  
  $outlook = new-object -comobject outlook.application  
  $namespace = $outlook.GetNameSpace(“MAPI”)  
  $folder = $namespace.getDefaultFolder($olFolders::olFolderCalendar) 
  
  # get current time and tomorrow to filter the calendar items 
  $now = Get-Date
  $tomorrow = (Get-Date).AddDays(1).Date

  # get all calendar items for today which are not ended yet, sort them by start time and select the properties to display
  $elements = $folder.items               
  $elements.Sort("[Start]")               
  $elements.IncludeRecurrences = $true    
 
  $filter = "[End] >= `"$($now.toString('g'))`" and [End] <= `"$($tomorrow.toString('g'))`" "
  $events = $elements.restrict($filter)
  # too long subjects will be cutted to 40 characters + "..."
  $subjectLength = 40
  $events | Select-Object -Property @{Name='Subject'; Expression={if ($_.Subject.Length -gt $subjectLength) {$_.Subject.Substring(0, $subjectLength) + "..."} else {$_.Subject}}}, @{Name='Start'; Expression={$_.Start.ToString("HH:mm")}}, 
  Duration,
  @{Name='EndIn'; Expression={if ($_.Start -lt (Get-Date))  {((New-TimeSpan -Start (Get-Date) -End $_.End).TotalMinutes) -as [int]} else {""}}}, 
  @{Name='StartIn'; Expression={if ($_.Start -gt (Get-Date)) {((New-TimeSpan -Start (Get-Date) -End $_.Start).TotalMinutes) -as [int]} else {""}}}, 
   Location | Format-Table -AutoSize
 } #end function Get-OutlookCalendar  
 
 # run function and refresh every 10 minutes
 while ($true) {
  Clear-Host
  Get-OutlookCalendar
  Start-Sleep -Seconds 600
}

