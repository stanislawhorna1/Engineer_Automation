Add-Type -AssemblyName System.Web
Add-Type -AssemblyName System.Windows.Forms

function Get-EngineList {
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
}; Set-Alias inxql Invoke-Nxql

Function Get-NxqlExport {
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
					-credentials $credentials, `
					-Query $Query)
			$out.Split("`n")[0..($out.Split("`n").count - 2)]
		}
		else {
			$out = (Invoke-Nxql -ServerName $engine.address `
					-PortNumber $webapiPort `
					-credentials $credentials, `
					-Query $Query)
			$out.Split("`n")[1..($out.Split("`n").count - 2)]
		}
		$counter = 1
	}
};

function Invoke-HashTableSort {
	param (
		[Parameter(Mandatory = $true)]
		[System.Collections.Hashtable]$Hashtable,
		[Parameter(Mandatory = $false)]
		[int]$Value_index,
		[Parameter(Mandatory = $false)]
		[switch]$Descending
	)
	if ($index) {
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
	Start-Sleep -Seconds 2
	while ($connections | ForEach-Object { if ($_.OLEDBConnection.Refreshing) { $true } }) {
		Start-Sleep -Milliseconds 500
	}
	Start-Sleep -Seconds 2
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
	[String]$temp_name = (65..90) | Get-Random -Count 5 | ForEach-Object {[char]$_}
	$temp_name = $temp_name.Replace(" ","")
	$temp_name = "C:\temp\" + $temp_name + ".csv"
	$work.Sheets.Item($Sheet_name).SaveAs($temp_name, 6)
	$work.Close()
	$excel.Quit()
	$csv = Import-Csv -Path $temp_name
	Remove-Item -Path $temp_name -Force
	$csv
}

function Get-AverageAppStartupTime {
	param (
		[Parameter(Mandatory = $true)]
		[System.Management.Automation.PSCustomObject]$SourceTable,
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
}

function Remove-duplicates {
	param (
		[Parameter(Mandatory = $true)]
		[System.Management.Automation.PSCustomObject]$SourceTable,
		[Parameter(Mandatory = $true)]
		[String]$ColumnNameGroup,
		[Parameter(Mandatory = $false)]
		[String]$ColumnNameSort,
		[Parameter(Mandatory = $false)]
		[switch]$Descending
	)
	if($ColumnNameSort){
		if (Descending) {
			$SourceTable = `
			($SourceTable | Group-Object $ColumnNameGroup | `
			ForEach-Object { $_.Group |Sort-Object $ColumnNameSort -Descending | `
			Select-Object -First 1 })
		}else{
			$SourceTable = `
			($SourceTable | Group-Object $ColumnNameGroup | `
			ForEach-Object { $_.Group |Sort-Object $ColumnNameSort | `
			Select-Object -First 1 })
		}
	}else{
		$SourceTable = ($SourceTable | Group-Object $ColumnNameGroup | ForEach-Object { $_.Group | Select-Object -First 1 })
	}
	$SourceTable
}



	# Create file name based on column
	$app_type = $col.Split("/")[1]
	$calc_f_name = $calc_folder + $app_type + ".csv"
	# Write colums headers to file
	Set-Content -Path $calc_f_name -Value "`"$app_type`",`"Average_start_time[s]`",`"Number_of_devices`""
	# Calculate average value and write to file
	foreach ($app in $hash_all.Keys) {
		$hash_all.$app[0] = ($hash_all.$app[0] / $hash_all.$app[1]) / 1000
		$avg = [math]::Round($hash_all.$app[0], 3)
		$num_dev = $hash_all.$app[1]
		$out = "`"" + $app + "`",`"" + $avg + "`",`"" + $num_dev + "`""
		Add-content $calc_f_name $out
	}