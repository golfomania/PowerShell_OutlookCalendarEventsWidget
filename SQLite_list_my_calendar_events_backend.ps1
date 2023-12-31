#############################################
# project:  List my next calender events from lokal Outlook (no MS Graph API needed)
# part:     backend (saving the events to a SQLite DB)
# Martin Löffler 
# 23.12.2023
# WORKS
#############################################

#############################################
# preperation of SQLite DB	
# download SQLite from https://www.sqlite.org/download.html
# 
# create a new DB named calendar.db with the command: .open calendar.db
# create a new table named table_meetings with the following columns:
# id (INTEGER PRIMARY KEY AUTOINCREMENT), subject (TEXT), start (TEXT), duration (TEXT), , startin (TEXT), endin (TEXT), location (TEXT) with the command: 
# CREATE TABLE meetings_table (id INTEGER PRIMARY KEY, name TEXT, start TEXT, duration INTEGER, endin INTERGER, startin INTEGER, location TEXT);
#
# get path to database with the command: .databases  e.g. C:\sqlite3\calendar.db
# check if the table is created with the command: .tables
# check the schama of the table with the command: .schema
# exit the SQLite CLI with the command: .exit
# 
# read: SELECT * FROM meetings_table;
# write: INSERT INTO meetings_table (name, start, duration, endin, startin, location) VALUES ('name1', 'start1', '60', '10', '1', 'location1');
# delete: DELETE FROM meetings_table WHERE id = 1;
#############################################

Function Get-OutlookCalendar  
 {  
  #############################################
  # use Outlook interop to get the default calendar folder
  #############################################
  # Search your PC in file explorer for "Microsoft.Office.Interop.Outlook.dll" to find one of this DLLs on your machine and paste the path to this DLL here
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

  
  #############################################
  # SQLite DB
  # get all events from the DB
  #############################################
  # Load the System.Data.SQLite assembly replace with path to your DLL
  # source https://system.data.sqlite.org/
  [Reflection.Assembly]::LoadFile("C:\sqlite3\sqlite-netFx46-static-binary-bundle-x64-2015-1.0.118.0\System.Data.SQLite.dll")

  # Create a connection to the database
  $connection = New-Object System.Data.SQLite.SQLiteConnection
  # replace with your path to the DB
  $connection.ConnectionString = "Data Source='C:\sqlite3\calendar.db';Version=3;"
  $connection.Open()

  if ($connection.State -eq 'Open') {
    Write-Host "Connection to DB is open"
  }

  #############################################
  # get all events from the DB
  #############################################
  $command = $connection.CreateCommand()
  $command.CommandText = "SELECT * FROM meetings_table;"
  $eventFromDatabaseSQLiteDataReader = $command.ExecuteReader()

  # internal List of DB events & reset the list for every run
  $eventsFromDatabase = New-Object System.Collections.ArrayList
 
  while ($eventFromDatabaseSQLiteDataReader.Read())
  {
    # write to console for debugging
    Write-Host $eventFromDatabaseSQLiteDataReader.GetValue(0) $eventFromDatabaseSQLiteDataReader.GetValue(1) $eventFromDatabaseSQLiteDataReader.GetValue(2) $eventFromDatabaseSQLiteDataReader.GetValue(3) $eventFromDatabaseSQLiteDataReader.GetValue(4) $eventFromDatabaseSQLiteDataReader.GetValue(5) $eventFromDatabaseSQLiteDataReader.GetValue(6)
    
    # create an object/hashtable for every event in the DB
    $newEvent = @{
      id = $eventFromDatabaseSQLiteDataReader.GetValue(0); 
      name = $eventFromDatabaseSQLiteDataReader.GetValue(1);
      start = $eventFromDatabaseSQLiteDataReader.GetValue(2);
      duration = $eventFromDatabaseSQLiteDataReader.GetValue(3);
      endin = $eventFromDatabaseSQLiteDataReader.GetValue(4);
      startin = $eventFromDatabaseSQLiteDataReader.GetValue(5);
      location = $eventFromDatabaseSQLiteDataReader.GetValue(6);
    }
    
    # add the event object to the internal list of DB events
    $eventsFromDatabase.Add($newEvent)
  }
  $eventFromDatabaseSQLiteDataReader.Close()

  foreach ($eventFromDatabase in $eventsFromDatabase)
  {
    # write to console for debugging
    Write-Host $eventFromDatabase.id $eventFromDatabase.name $eventFromDatabase.start $eventFromDatabase.duration $eventFromDatabase.endin $eventFromDatabase.startin $eventFromDatabase.location
  }

  #############################################
  # updated existing events in the DB
  # delete events which are no longer in Outlook but in the DB
  #############################################
  
  # Iterate through these events of the Database and check them against the Outlook events
  Write-Host "Start update & delete loop"
  foreach ($eventFromDatabase in $eventsFromDatabase)
  {
    # write to console for debugging
    Write-Host "Current events from DB:"
    Write-Host $eventFromDatabase.id $eventFromDatabase.name $eventFromDatabase.start $eventFromDatabase.duration $eventFromDatabase.endin $eventFromDatabase.startin $eventFromDatabase.location

    # check if there is an event in $events where the name/subject and the Starttime is the same as in the DB
    $eventExistsInOutlook = $events | Where-Object {$_.Subject -eq $eventFromDatabase.name -and $_.Start.ToString("HH:mm") -eq $eventFromDatabase.start}

    # if event exist in DB and Outlook, update the entry in the DB
    if ($eventExistsInOutlook) {
      # write to console for debugging
      Write-Host "Database event exists in Outlook, DB entry will be updated"
      # Create a command to insert data into a table
      $command = $connection.CreateCommand()
      $command.CommandText = "UPDATE meetings_table SET name = '$($eventExistsInOutlook.Subject)', start = '$($eventExistsInOutlook.Start.ToString("HH:mm"))', duration = '$($eventExistsInOutlook.Duration)', endin = $(if ($eventExistsInOutlook.Start -lt (Get-Date))  {((New-TimeSpan -Start (Get-Date) -End $eventExistsInOutlook.End).TotalMinutes) -as [int]} else {'NULL'}), startin = $(if ($eventExistsInOutlook.Start -gt (Get-Date)) {((New-TimeSpan -Start (Get-Date) -End $eventExistsInOutlook.Start).TotalMinutes) -as [int]} else {'NULL'}), location = '$($eventExistsInOutlook.Location)' WHERE name = '$($eventFromDatabase.name)' AND start = '$($eventFromDatabase.start)';"
      $command.ExecuteNonQuery()  
    } 

    # if event does exist in DB but not in DB, delete the entry from the DB
    else {
      # write to console for debugging
      Write-Host "Database event does no longer exist in Outlook, DB entry will be deleted"
      # Create a command to delete data from a table
      $command = $connection.CreateCommand()
      $command.CommandText = "DELETE FROM meetings_table WHERE id = $($eventFromDatabase.id);"
      $command.ExecuteNonQuery()
    }
  } # end of updating and deleting loop 
  Write-Host "End update & delete loop"

  #############################################
  # add new Outlook events to the DB
  #############################################
  # iterate through all events from Outlook
  Write-Host "Start add new loop"
  foreach ($outlookEvent in $events) {
    # check if there is an event in the DB where the name/subject and the Starttime is the same as in Outlook
    # check $outlookEvent.Subject and $outlookEvent.Start.ToString("HH:mm") against the $eventsFromDatabase
    $eventExistsInDB = $eventsFromDatabase | Where-Object {$_.name -eq $outlookEvent.Subject -and $_.start -eq $outlookEvent.Start.ToString("HH:mm")}

    # if event does not exist in DB, add new entry to the DB
    if (!$eventExistsInDB) {
      # write to console for debugging
      Write-Host "Outlook event does not exist in DB, entry will be created"
      # Create a command to insert data into a table
      $command = $connection.CreateCommand()
      $command.CommandText = "INSERT INTO meetings_table (name, start, duration, endin, startin, location) VALUES ('$($outlookEvent.Subject)', '$($outlookEvent.Start.ToString("HH:mm"))', '$($outlookEvent.Duration)', $(if ($outlookEvent.Start -lt (Get-Date))  {((New-TimeSpan -Start (Get-Date) -End $outlookEvent.End).TotalMinutes) -as [int]} else {'NULL'}), $(if ($outlookEvent.Start -gt (Get-Date)) {((New-TimeSpan -Start (Get-Date) -End $outlookEvent.Start).TotalMinutes) -as [int]} else {'NULL'}), '$($outlookEvent.Location)')"
      $command.ExecuteNonQuery()  
    } 
  } # end of adding loop
  Write-Host "End add new loop"


  # Close the DB connection
  $connection.Close()
 } # end of Get-OutlookCalendar function 
 

#############################################
# run function and refresh every minute
#############################################
while ($true) {
 Clear-Host
 Get-OutlookCalendar
 Start-Sleep -Seconds 60
}

