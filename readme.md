# Change log
-
## To-Do
Convert all functions that operates on hash table to format:
Name                           Value
----                           -----
PC-012AW                       {@{Device name=PC-012AW; Average boot time=2502; Number of events=3379}}
PC-013OX                       {@{Device name=PC-013OX; Average boot time=6406; Number of events=1292}}
PC-017CO                       {@{Device name=PC-017CO; Average boot time=1643; Number of events=7172}}
PC-017TS                       {@{Device name=PC-017TS; Average boot time=9370; Number of events=895}}
PC-01DTK                       {@{Device name=PC-01DTK; Average boot time=493; Number of events=2187}}

!! Test performance on big amount of data !!

CSV file structure:
"Device name","Average boot time","Number of events"
"PC-012AW","2502","3379"
"PC-013OX","6406","1292"
"PC-017CO","1643","7172"
"PC-017TS","9370","895"
"PC-01DTK","493","2187"

How to acces particlar data?
Retrieve Number of events from device PC-012AW:
$hash.'PC-012AW'.'Number of events'

how to add data to overall hash?
$hash.Add("My Device", @{"Device name" = "My Device"; "Average boot time" = 123; "Number of events" = 456})

how to add details for particular device?
$hash.'My Device'.Add("Number of system boots", 12345) 

#### Get-EngineList 
Connets to Nexthink Portal and retrieves list of all engines, next select only connected ones.

#### Invoke-Nxql
Sends an NXQL query to the Web API of Nexthink Engine as HTTP GET using HTTPS.

#### Get-NxqlExport
Returns formatted output from multiple Nexthink Engines, based on provided Engine list (Hashtable). Output of function can be saved to file or variable. Uses Invoke-Nxql function to retrieve data from each engine one by one.

#### Invoke-HashTableSort
Sorts hash table that contains array in value and returns new one.

#### Invoke-Popup
Display Windows 10 pop-up message based on the title and description provided
Powershell icon will be visible in the pop-up

#### Invoke-ExcelFileUpdate
Function copies report template to destination location, updates all table queries and pivot tables which exists in a given file, breaks all external connections

#### Export-ExcelToCsv
Exports selected Excel sheet to csv format. As working directory C:\temp\ is used
Function generated random string name for temp file, at the end file is removed

#### Get-AverageAppStartupTime
Calculates Average Startup Time form Nexthink RA "Get Startup impact" from given column name

#### Get-AverageGpoTime
Calculates GPO procesing time for each gpo category based on Nexthink RA

#### Remove-duplicates
Group entries by given column, sort by given column and select first result from sorting

#### Export-HashTableToCsv
Converts hashtable to csv format, works with hash where value is 1 dimmension array,
result can be saved to file or variable

#### New-RandomHashTable
Generates Hash table with 1 dimmension array in value

#### New-DataSummary
Creates summary like Excel pivot table, based on table provided and selected columns