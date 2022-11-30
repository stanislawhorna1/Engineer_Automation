Add-Type -AssemblyName System.Web
Add-Type -AssemblyName System.Windows.Forms

function Get-EngineList {
	<#
.SYNOPSIS
Returns list of Engines connected to Nexthink Portal

.DESCRIPTION
Connets to Nexthink Portal and retrieves list of all engines,
next select only connected ones.

.PARAMETER portal
The Nexthink Portal DNS Name to retrieve connected engines

.PARAMETER credentials
Nexthink account authorised to extract list of engines

.EXAMPLE
Get-EngineList -portal "test.eu.nexthink.cloud" -credentials <Account_UserName>

.INPUTS
String

.OUTPUTS
Hastable

.NOTES
    Author:  Stanislaw Horna
#>
	param (
		[Parameter(Mandatory = $true)]
		[string] $portal,
		[Parameter(Mandatory = $true)]
		[System.Management.Automation.PSCredential]$credentials
	)
	$web = [Net.WebClient]::new()
	$web.Credentials = $credentials
	$pair = [string]::Join(":", $web.Credentials.UserName, $web.Credentials.Password)
	$base64 = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($pair))
	$web.Headers.Add('Authorization', "Basic $base64")
	$baseUrl = "https://$portal/api/configuration/v1/engines"
	$result = $web.downloadString($baseUrl)
	$engineList = $result | ConvertFrom-Json
	# Listing Connected Engines only
	$engineList | Where-Object { $_.status -eq "CONNECTED" }
};
Function Invoke-Nxql {
	<#
	.SYNOPSIS
	Sends an NXQL query to a Nexthink engine.

	.DESCRIPTION
	 Sends an NXQL query to the Web API of Nexthink Engine as HTTP GET using HTTPS.
	 
	.PARAMETER ServerName
	 Nexthink Engine name or IP address.

	.PARAMETER PortNumber
	Port number of the Web API (default 1671).

	.PARAMETER UserName
	User name of the Finder account under which the query is executed.

	.PARAMETER UserPassword
	User password of the Finder account under which the query is executed.

	.PARAMETER NxqlQuery
	NXQL query.

	.PARAMETER FirstParamter
	Value of %1 in the NXQL query.

	.PARAMETER SecondParamter
	Value of %2 in the NXQL query.

	.PARAMETER OuputFormat
	NXQL query output format i.e. csv, xml, html, json (default csv).

	.PARAMETER Platforms
	Platforms on which the query applies i.e. windows, mac_os, mobile (default windows).
	
	.EXAMPLE
	Invoke-Nxql -ServerName 176.31.63.200 -UserName "admin" -UserPassword "admin" 
	-Platforms=windows,mac_os -NxqlQuery "(select (name) (from device))"
	#>
	Param(
		[Parameter(Mandatory = $true)]
		[string]$ServerName,
		[Parameter(Mandatory = $true)]
		[System.Management.Automation.PSCredential]$credentials,
		[Parameter(Mandatory = $true)]
		[string]$Query,
		[Parameter(Mandatory = $false)]
		[int]$PortNumber = 1671,
		[Parameter(Mandatory = $false)]
		[string]$OuputFormat = "csv",
		[Parameter(Mandatory = $false)]
		[string[]]$Platforms = "windows",
		[Parameter(Mandatory = $false)]
		[string]$FirstParameter,
		[Parameter(Mandatory = $false)]
		[string]$SecondParameter
	)
	$PlaformsString = ""
	Foreach ($platform in $Platforms) {
		$PlaformsString += "&platform={0}" -f $platform
	}
	$EncodedNxqlQuery = [System.Web.HttpUtility]::UrlEncode($Query)
	$Url = "https://{0}:{1}/2/query?query={2}&format={3}{4}" -f $ServerName, $PortNumber, $EncodedNxqlQuery, $OuputFormat, $PlaformsString
	if ($FirstParameter) { 
		$EncodedFirstParameter = [System.Web.HttpUtility]::UrlEncode($FirstParameter)
		$Url = "{0}&p1={1}" -f $Url, $EncodedFirstParameter
	}
	if ($SecondParameter) { 
		$EncodedSecondParameter = [System.Web.HttpUtility]::UrlEncode($SecondParameter)
		$Url = "{0}&p2={1}" -f $Url, $EncodedSecondParameter
	}
	#echo $Url
	try {
		[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11
		[Net.ServicePointManager]::ServerCertificateValidationCallback = { $true } 
		$webclient = New-Object system.net.webclient
		$webclient.Credentials = New-Object System.Net.NetworkCredential($Credentials.UserName, $credentials.GetNetworkCredential().Password)
		$webclient.DownloadString($Url)
	}
	catch {
		Write-Host $Error[0].Exception.Message
	}
};
Function Get-NxqlExport {
	<#
.SYNOPSIS
Returns output from multiple Nexthink Engines

.DESCRIPTION
Returns formatted output from multiple Nexthink Engines, based on provided Engine list (Hashtable).
Output of function can be saved to file or variable
Uses Invoke-Nxql function to retrieve data from each engine one by one.

.PARAMETER Query
String that contains NXQL query. 

.PARAMETER credential
Nexthink account authorised to extract data from engines

.PARAMETER webapiPort
WebAPI port: 
	for SaaS - "443"
	for on-premise "1671" 


.PARAMETER EngineList
List of connected engines in Hashtable, output of Get-EngineList can be used.

.EXAMPLE
Get-NxqlExport -Query "NXQL query" `
			   -credentials <variable_created_by "Get-Credential"> `
			   -webapiPort "443" `
			   -EngineList <output from Get-EngineList>

.OUTPUTS
String

.NOTES
    Author:  Stanislaw Horna
#>
	param (
		[Parameter(Mandatory = $true)]
		[String] $Query,
		[Parameter(Mandatory = $true)]
		[System.Management.Automation.PSCredential] $credentials,
		[Parameter(Mandatory = $true)]
		[String]$webapiPort,
		[Parameter(Mandatory = $true)]
		$EngineList
	)
	$counter = 0
	foreach ($engine in $EngineList) {
		if ($counter -eq 0) {
			$out = (Invoke-Nxql -ServerName $engine.address `
					-PortNumber $webapiPort `
					-credentials $credentials `
					-Query $Query)
			$out.Split("`n")[0..($out.Split("`n").count - 2)]
		}
		else {
			$out = (Invoke-Nxql -ServerName $engine.address `
					-PortNumber $webapiPort `
					-credentials $credentials `
					-Query $Query)
			$out.Split("`n")[1..($out.Split("`n").count - 2)]
		}
		$counter = 1
	}
};
function Invoke-HashTableSort {
	<#
.SYNOPSIS
Returns sorted hashtable.

.DESCRIPTION
Sorts hash table that contains array in value and returns new one.

.PARAMETER Hashtable
Hastable to sort

.PARAMETER Value_index
Index of value located in array by which entries will be sorted

.PARAMETER Descending
If included values will be sorted descending otherwise ascending

.EXAMPLE


.OUTPUTS
Hashtable

.NOTES
    Author:  Stanislaw Horna
#>
	param (
		[Parameter(Mandatory = $true)]
		[System.Collections.Hashtable]$Hashtable,
		[Parameter(Mandatory = $false)]
		[int]$Value_index,
		[Parameter(Mandatory = $false)]
		[switch]$Descending
	)
	if ($Value_index) {
		if ($Descending) {
			$hashSorted = [ordered] @{}
			$Hashtable.GetEnumerator() | Sort-Object { $_.Value[$Value_index] } -Descending | ForEach-Object { $hashSorted[$_.Key] = $_.Value }
		}
		else {
			$hashSorted = [ordered] @{}
			$Hashtable.GetEnumerator() | Sort-Object { $_.Value[$Value_index] } | ForEach-Object { $hashSorted[$_.Key] = $_.Value }
		}
	}
	else {
		if ($Descending) {
			$hashSorted = [ordered] @{}
			$Hashtable.GetEnumerator() | Sort-Object { $_.Value } -Descending | ForEach-Object { $hashSorted[$_.Key] = $_.Value }
		}
		else {
			$hashSorted = [ordered] @{}
			$Hashtable.GetEnumerator() | Sort-Object { $_.Value } | ForEach-Object { $hashSorted[$_.Key] = $_.Value }
		}
	}
	$hashSorted
}
Function Invoke-Popup {
	<#
.SYNOPSIS
Display windows 10 pop-up message

.DESCRIPTION
Display Windows 10 pop-up message based on the title and description provided
Powershell icon will be visible in the pop-up

.PARAMETER title
Message title 1 line

.PARAMETER description
Message description multiple lines

.EXAMPLE
Invoke-Popup -title "Report "ready" `
			 -description "Automatically created report is ready"


.INPUTS
String

.OUTPUTS
None

.NOTES
    Author:  Stanislaw Horna
#>
	param (
		[Parameter(Mandatory = $true)]
		[string] $title,
		[Parameter(Mandatory = $true)]
		[String] $description
	)
	$global:endmsg = New-Object System.Windows.Forms.Notifyicon
	$p = (Get-Process powershell | Sort-Object -Property CPU -Descending | Select-Object -First 1).Path
	$endmsg.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($p)
	$endmsg.BalloonTipTitle = $title
	$endmsg.BalloonTipText = $description
	$endmsg.Visible = $true
	$endmsg.ShowBalloonTip(10)
};
function Invoke-ExcelFileUpdate {
	<#
.SYNOPSIS
Updates Excel report template

.DESCRIPTION
Function copies report template to destination location,
updates all table queries and pivot tables which exists in a given file,
breaks all external connections

.PARAMETER SourceFile
Template file location full path

.PARAMETER DestinationFile
Destination file location (with name) 

.EXAMPLE
Invoke-ExceFileUpdate -SourceFile C:\temp\template.xlsx `
					  -DestinationFile C:\Users\User\Documents\report.xlsx

.INPUTS
File paths as a strings

.OUTPUTS
None

.NOTES
    Author:  Stanislaw Horna
#>
	param (
		[Parameter(Mandatory = $true)]
		[string] $SourceFile,
		[Parameter(Mandatory = $true)]
		[String] $DestinationFile
	)
	# Test if file already exist
	if ((Test-Path -Path $DestinationFile) -eq $true) {
		Remove-item -path $DestinationFile -Confirm:$false -Force
	}
	# Copy template to destination location
	Copy-Item -Path $SourceFile -Destination $DestinationFile | Out-Null
	# Open Excel App
	$excel = New-Object -ComObject Excel.Application
	$excel.Visible = $false
	$excel.DisplayAlerts = $false
	Start-Sleep -Seconds 2
	# Open Excel file
	$work = $excel.Workbooks.Open($DestinationFile)
	Start-Sleep -Seconds 2
	$connections = $work.connections
	# Refresh existing Table Queries
	$work.RefreshAll()
	Start-Sleep -Seconds 10
	while ($connections | ForEach-Object { if ($_.OLEDBConnection.Refreshing) { $true } }) {
		Start-Sleep -Milliseconds 500
	}
	Start-Sleep -Seconds 10
	# Get list of available sheets in file
	$list_sheets = $work.Worksheets | Select-Object name, index
	# For Each sheet update Pivot table if exist
	foreach ($sheet in $list_sheets) {
		$sh = $work.Worksheets.Item($sheet.Name)
		$pivots = $sh.PivotTables()
		for ($i = 1; $i -le $pivots.Count; $i++ ) {
			$pivots.Item($i).RefreshTable() | Out-Null
		}
	}
	# Get number of Table Queries
	$num_of_queries_to_delete = ($work.Queries).Count
	# Brake All Table Queries
	for ($i = 1; $i -le $num_of_queries_to_delete; $i++) {
		$work.Queries.Item(1).Delete()
	}
	# Save all done work and close file and Application
	$work.Save()
	$work.Close()
	$excel.Quit()
};
function Export-ExcelToCsv {
	<#
.SYNOPSIS
Saves Excel sheet to csv format

.DESCRIPTION
Exports selected Excel sheet to csv format. As working directory C:\temp\ is used
Function generated random string name for temp file, at the end file is removed

.PARAMETER SourceFile
Full path to Excel source file from which data will be exported

.PARAMETER Sheet_name
Excel Worksheet name which will be exported

.PARAMETER Refresh
Switch if existing table queries should be refreshed before exporting data

.EXAMPLE
$csv = Export-ExcelToCsv -SourceFile C:\temp\report.xslx -Sheet_name "Data"

.INPUTS
String

.OUTPUTS
PSCustomObject

.NOTES
    Author:  Stanislaw Horna
#>
	param (
		[Parameter(Mandatory = $true)]
		[string] $SourceFile,
		[Parameter(Mandatory = $false)]
		[string]$Sheet_name,
		[Parameter(Mandatory = $false)]
		[switch]$Refresh		
	)
	# Open Excel Application
	$excel = New-Object -ComObject Excel.Application
	$excel.Visible = $false
	$excel.DisplayAlerts = $false
	Start-Sleep -Seconds 2
	$work = $excel.Workbooks.Open($SourceFile)
	Start-Sleep -Seconds 2
	if ($Refresh) {
		$connections = $work.connections
		# Refresh all tables related to the source file
		$work.RefreshAll()
		# Wait until all tables will be refreshed
		Start-Sleep -Seconds 2
		while ($connections | ForEach-Object { if ($_.OLEDBConnection.Refreshing) { $true } }) {
			Start-Sleep -Milliseconds 500
		}
		Start-Sleep -seconds 2
	}
	[String]$temp_name = (65..90) | Get-Random -Count 5 | ForEach-Object { [char]$_ }
	$temp_name = $temp_name.Replace(" ", "")
	$temp_name = "C:\temp\" + $temp_name + ".csv"
	$work.Sheets.Item($Sheet_name).SaveAs($temp_name, 6)
	$work.Close()
	$excel.Quit()
	$csv = Import-Csv -Path $temp_name
	Remove-Item -Path $temp_name -Force
	$csv
};
function Get-AverageAppStartupTime {
	<#
.SYNOPSIS
Calculates Average startup time based on Nexthink RA results

.DESCRIPTION
Calculates Average Startup Time form Nexthink RA "Get Startup impact" from given column name

.PARAMETER SourceTable
Imported csv file with RA results

.PARAMETER ColumnName
Column Name where startup impact data is located

.EXAMPLE
Get-AverageAppStartupTime -SourceTable $csv_table -ColumnName "HignImpactApplications"


.INPUTS
PSCustomObject
String

.OUTPUTS
Hashtable

.NOTES
    Author:  Stanislaw Horna,
			 Pawel Bielinski
#>
	param (
		[Parameter(Mandatory = $true)]
		$SourceTable,
		[Parameter(Mandatory = $true)]
		[String]$ColumnName
	)
	$hash_all = @{}
	for ($j = 0; $j -lt $SourceTable.Count; $j++) {
		if (($SourceTable[$j].$ColumnName)[0] -ne "-") {
			# Extract line for all high impact application on device and remove ", " for applications where it is used in name
			$line = (($SourceTable[$j].$ColumnName) -replace ", ", "") -split ","
			$listed_apps = @()
			for ($i = 0; $i -lt $line.Count; $i++) {
				# Get app name and extract stats as array (miliseconds, MB)
				$appName = $line[$i].Substring(0, $line[$i].LastIndexOf("(") - 1)
				$appStats = $line[$i].Substring($appName.Length + 1).Split("(")[1].Split("ms;")[0]
				# If app was listed multiple times on one device include only the first instance
				if ($appName -notin $listed_apps) {
					# If application exist in hashtable, increment values, else add as new entry
					if ($appName -in $hash_all.Keys) {
						$hash_all.$appName[0] += [int]$appStats
						$hash_all.$appName[1]++
					}
					else {
						$hash_all.Add($appName, @([int]$appStats, 1))
					}
					$listed_apps += $appName
				}
			}
		}
	}
	foreach ($app in $hash_all.Keys) {
		$hash_all.$app[0] = [math]::Round((($hash_all.$app[0] / $hash_all.$app[1]) / 1000), 3)
	}
	$hash_all
};
function Get-AverageGpoTime {
	<#
.SYNOPSIS
Calculates Average GPO processing time based on Nexthink RA results

.DESCRIPTION
Calculates GPO procesing time for each gpo category based on Nexthink RA

.PARAMETER SourceTable
Imported csv file with RA results

.PARAMETER ColumnName
Column Name where GPO startup impact data is located

.EXAMPLE
Get-AverageGpoTime -SourceTable $csv_table -ColumnName "UserGPOtime"

.INPUTS
PSCustomObject
String

.OUTPUTS
Hashtable

.NOTES
    Author:  Stanislaw Horna,
			 Pawel Bielinski
#>
	param (
		[Parameter(Mandatory = $true)]
		$SourceTable,
		[Parameter(Mandatory = $true)]
		[String]$ColumnName
	)
	$hash_all = @{}
	for ($j = 0; $j -lt $SourceTable.Count; $j++) {
		if ((($SourceTable[$j].$ColumnName)[0] -ne "-") -and (($SourceTable[$j].$ColumnName)[0].Length -ne 0)) {
			# Extract line for all GPO categories on device
			$line = (($SourceTable[$j].$ColumnName) -replace ", ", ",").Split(",")
			$listed_category = @()
			for ($i = 0; $i -lt $line.Count; $i++) {
				# Get category name and extract stats as array (miliseconds)
				$CategoryName = $line[$i].Substring(0, $line[$i].LastIndexOf("(") - 1)
				$CategoryStats = $line[$i].Substring($CategoryName.Length + 2).Split(" ")[0]
				# If category was listed multiple times on one device sum time but do not increment number of devices
				if ($CategoryName -notin $listed_category) {
					# If category exist in hashtable, increment values, else add as new entry
					if ($CategoryName -in $hash_all.Keys) {
						$hash_all.$CategoryName[0] += [int]$CategoryStats
						$hash_all.$CategoryName[1]++
					}
					else {
						$hash_all.Add($CategoryName, @([int]$CategoryStats, 1))
					}
					$listed_category += $CategoryName
				}
				else {
					$hash_all.$CategoryName[0] += [int]$CategoryStats
				}
			}
		}
	}
	foreach ($category in $hash_all.Keys) {
		$hash_all.$category[0] = [math]::Round((($hash_all.$category[0] / $hash_all.$category[1]) / 1000), 3)
	}
	$hash_all
};
function Remove-Duplicates {
	<#
.SYNOPSIS
Remove duplicates - work the same as Excel function

.DESCRIPTION
Group entries by given column, sort by given column and select first result from sorting

.PARAMETER SourceTable
Csv imported table

.PARAMETER ColumnNameGroup
Column name where only unique values should remain

.PARAMETER ColumnNameSort
Column name by which values should be sorted befor removing duplicated vaules
Can be empty

.PARAMETER DateTime
Select if ColumnNameSort should be sorted as date time

.EXAMPLE



.INPUTS


.OUTPUTS
PSCustomObject

.NOTES
    Author:  Stanislaw Horna
			 Pawel Bielinski
#>
	param (
		[Parameter(Mandatory = $true)]
		$SourceTable,
		[Parameter(Mandatory = $true)]
		[String]$ColumnNameGroup,
		[Parameter(Mandatory = $false)]
		[String]$ColumnNameSort,
		[Parameter(Mandatory = $false)]
		[switch]$Descending,
		[Parameter(Mandatory = $false)]
		[switch]$DateTime
	)
	$Hash = @{}
	$ErrorActionPreference = 'SilentlyContinue'
	Switch ($true) {
		($ColumnNameSort -and $DateTime -and $Descending) {
			$SourceTable | Sort-Object -property { [System.DateTime]::ParseExact($_.$ColumnNameSort, "yyyy-MM-dd'T'HH:mm:ss", $null)} -Descending `
			| ForEach-Object {
    				if (-not($Hash.ContainsKey($_.$ColumnNameGroup))) {
							$Hash.Add($_.$ColumnNameGroup, $_)
    					}
					}
		}
		($ColumnNameSort -and $DateTime){
			$SourceTable | Sort-Object -property { [System.DateTime]::ParseExact($_.$ColumnNameSort, "yyyy-MM-dd'T'HH:mm:ss", $null)} `
			| ForEach-Object {
    				if (-not($Hash.ContainsKey($_.$ColumnNameGroup))) {
							$Hash.Add($_.$ColumnNameGroup, $_)
    					}
					}
		}
		($ColumnNameSort -and $Descending){
			$SourceTable | Sort-Object -property $ColumnNameSort -Descending `
			| ForEach-Object {
    				if (-not($Hash.ContainsKey($_.$ColumnNameGroup))) {
							$Hash.Add($_.$ColumnNameGroup, $_)
    					}
					}
		}
		($ColumnNameSort){
			$SourceTable | Sort-Object -property $ColumnNameSort -Descending `
			| ForEach-Object {
    				if (-not($Hash.ContainsKey($_.$ColumnNameGroup))) {
							$Hash.Add($_.$ColumnNameGroup, $_)
    					}
					}
		}
		Default{
			$SourceTable  `
			| ForEach-Object {
    				if (-not($Hash.ContainsKey($_.$ColumnNameGroup))) {
							$Hash.Add($_.$ColumnNameGroup, $_)
    					}
					}
		}

	}
	$ErrorActionPreference = 'Continue'
	$Hash.Values
};
function Export-HashTableToCsv {
	<#
.SYNOPSIS
Converts Hashtable to csv

.DESCRIPTION
Converts hashtable to csv format, works with hash where value is 1 dimmension array,
result can be saved to file or variable

.PARAMETER Hashtable
Hashtable which will be converted

.PARAMETER Headers
Array of headers in an specified order

.PARAMETER Path
Specify only if result should be save to file

.EXAMPLE
$csv = Export-HashTableToCsv -Hashtable $hash `
	-Headers @("Device name", "Average process count", "CPU time") | ConvertFrom-Csv


.INPUTS
Hashtable
Array

.OUTPUTS
PSCustomObject

.NOTES
    Author:  Stanislaw Horna
#>
	param (
		[Parameter(Mandatory = $true)]
		[System.Collections.Hashtable]$Hashtable,
		[Parameter(Mandatory = $true)]
		$Headers,
		[Parameter(Mandatory = $false)]
		[String]$Path
	)

	$out = "`"" + $Headers[0] + "`","
	for ($i = 1; $i -lt $Headers.Count - 1; $i++) {
		$out = $out + "`"" + $Headers[$i] + "`"," 
	}
	$out = $out + "`"" + $Headers[($Headers.Count) - 1] + "`""
	if ($Path) {
		Set-Content -Path $Path -Value $out
		foreach ($item in $Hashtable.Keys) {
			$out = "`"" + $item + "`","
			for ($i = 0; $i -lt (($Hashtable.$item.Count) - 1); $i++) {
				$out = $out + "`"" + $Hashtable.$item[$i] + "`","
			}
			$out = $out + "`"" + $Hashtable.$item[($Hashtable.$item.Count) - 1] + "`""
			Add-content $Path $out
			# $out
		}
	}
	else {
		$out
		foreach ($item in $Hashtable.Keys) {
			$out = "`"" + $item + "`","
			for ($i = 0; $i -lt (($Hashtable.$item.Count) - 1); $i++) {
				$out = $out + "`"" + $Hashtable.$item[$i] + "`","
			}
			$out = $out + "`"" + $Hashtable.$item[($Hashtable.$item.Count) - 1] + "`""
			# Add-content $Path $out
			$out
		}
	}

}
function New-RandomHashTable {
	<#
.SYNOPSIS
Generates Hash table with array in value

.NOTES
    Author:  Stanislaw Horna
#>
	param(
		[Parameter(Mandatory = $false)]
		[int]$NumberOfEntries	
	)
	if (-not $NumberOfEntries) {
		$NumberOfEntries = 100
	}
	$hash = @{}
	for ($i = 0; $i -lt $NumberOfEntries; $i++) {
		[String]$RanStrName = (48..57) + (65..90) | Get-Random -Count 5 | ForEach-Object { [char]$_ }
		$RanStrName = $RanStrName.Replace(" ", "")
		$RanStrName = "PC-" + $RanStrName
		if ($RanStrName -notin $hash.Keys) {
			$RanNumVal1 = Get-Random -Maximum 10000 -Minimum 1
			$RanNumVal2 = Get-Random -Maximum 10000 -Minimum 1
			$hash.Add($RanStrName, @($RanNumVal1, $RanNumVal2))
		}
		else {
			$i--
		}
	}
	$hash	
}
function New-DataSummary {
	<#
.SYNOPSIS
Creates data summary in Hashtable

.DESCRIPTION
Creates summary like Excel pivot table, based on table provided and selected columns

.PARAMETER SourceTable
Source data table

.PARAMETER RowsColumn
Simmilar to Excel PivotTable column, on which rows are created 

.PARAMETER AverageColumn
Vaule which is needed as a average for each unique value from RowsColumn

.INPUTS
Table
String

.OUTPUTS
Hashtable

.NOTES
    Author:  Stanislaw Horna
#>
	param (
		[Parameter(Mandatory = $true)]
		$SourceTable,
		[Parameter(Mandatory = $true)]
		[String]$RowsColumn,
		[Parameter(Mandatory = $true)]
		[String]$AverageColumn
	)
	$hash_all = @{}
	for ($j = 0; $j -lt $SourceTable.Count; $j++) {
		# Get category name and extract stats as array (miliseconds)
		$RowName = $SourceTable[$j].$RowsColumn
		if (($RowName -ne "-") -and ($RowName.length -gt 2)) {
		
			$RowStats = $SourceTable[$j].$AverageColumn
			# If category was listed multiple times on one device sum time but do not increment number of devices
			if ($RowName -in $hash_all.Keys) {
				$hash_all.$RowName[0] += [int]$RowStats
				$hash_all.$RowName[1]++
			}
			else {
				$hash_all.Add($RowName, @([int]$RowStats, 1))
			}
		}
	}
	foreach ($Item in $hash_all.Keys) {
		$hash_all.$Item[0] = [math]::Round((($hash_all.$Item[0] / $hash_all.$Item[1])), 3)
	}
	$hash_all
}
function Convert-FromCsvToHashtable {
	param (
		[Parameter(ValueFromPipeline)]
		$SourceTable,
		[Parameter(Mandatory = $true)]
		[String]$ColumnID,
		[Parameter(Mandatory = $false)]
		[String]$Path,
		[Parameter(Mandatory = $false)]
		[String]$Delimiter
	)
	if ($Path) {
		$SourceTable = Import-Csv -Path $Path -Delimiter $Delimiter
	}
	$Hash = $SourceTable | Group-Object -AsHashTable -Property $ColumnID
	$Hash
}
function Convert-FromHashtableToCsv {
	param (
		[Parameter(Mandatory = $true)]
		$SourceHashtable,
		[Parameter(Mandatory = $false)]
		[String]$Path
	)
	$csv = $SourceHashtable.GetEnumerator() | ForEach-Object { $SourceHashtable[$_.Key].GetEnumerator() }
	if ($Path) {
		$csv | Export-Csv -Path $Path -NoTypeInformation
	}
	else {
		$csv
	}
}

