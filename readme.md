# Change log

06.12.2022 - Get-NxqlExport - Handling queries, which returns blank output from some engines, loading bar added
30.11.2022 - Remove-Duplicates performance improvement


## To-Do
Review performance on HashTable structure on big data 
Potential use case - fast vlookup for example Device properties
$CSV | Group-Object -AsHashTable -Property $ColumnID

# Function Description

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