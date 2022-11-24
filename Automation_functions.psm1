Add-Type -AssemblyName System.Web
Add-Type -AssemblyName System.Windows.Forms

function Get-EngineList {
	param (
		[Parameter(Mandatory = $true)]
		[string] $portal,
		[Parameter(Mandatory = $true)]
		[SecureString] $credentials
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
		[string]$UserName,
		[Parameter(Mandatory = $true)]
		[string]$UserPassword,
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
		$webclient.Credentials = New-Object System.Net.NetworkCredential($UserName, $UserPassword)
		$webclient.DownloadString($Url)
	}
	catch {
		Write-Host $Error[0].Exception.Message
	}
}; Set-Alias inxql Invoke-Nxql

Function Get-NxqlExport {
	param (
		[Parameter(Mandatory = $true)]
		[string] $query,
		[Parameter(Mandatory = $true)]
		[SecureString] $credentials,
		[Parameter(Mandatory = $true)]
		$webapiPort
	)
	$counter = 0
	foreach ($engine in $activeEngines) {
		if ($counter -eq 0) {
			$out = (Invoke-Nxql -ServerName $engine.address -PortNumber $webapiPort -UserName $username -UserPassword $password -Query $Query)
			$out.Split("`n")[0..($out.Split("`n").count - 2)]
		}
		else {
			$out = (Invoke-Nxql -ServerName $engine.address -PortNumber $webapiPort -UserName $username -UserPassword $password -Query $Query)
			$out.Split("`n")[1..($out.Split("`n").count - 2)]
		}
		$counter = 1
	}
};

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














# Extract Application names and stats from output
$csv = Import-Csv -Path $Output_folder/$nxql_file_name -Delimiter "`t"
foreach ($col in $column_name) {
	$hash_all = @{}
	for ($j = 0; $j -lt $csv.Count; $j++) {
		if (($csv[$j].$col)[0] -ne "-") {
			# Extract line for all high impact application on device and remove ", " for applications where it is used in name
			$line = (($csv[$j].$col) -replace ", ", "") -split ","
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
	# Sort Startup Apps descending by number of devices where it was detected
	$hashSorted = [ordered] @{}
	$hash_all.GetEnumerator() | Sort-Object { $_.Value[1] } -Descending | ForEach-Object { $hashSorted[$_.Key] = $_.Value }
	$hash_all = $hashSorted
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
}


####################################################################
##  UPDATE EXCEL FILE
####################################################################


$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$work = $excel.Workbooks.Open($Excel_file)
$connections = $work.connections
$work.RefreshAll()
while($connections | ForEach-Object {if($_.OLEDBConnection.Refreshing){$true}}){
	Start-Sleep -Milliseconds 500
}
Start-Sleep -seconds 10

foreach($s in $sheet_with_pivots){
	$sheet = $work.Worksheets.Item($s)
	$pivots = $sheet.PivotTables()
	for($i=1; $i -le $pivots.Count; $i++ ){
		$pivots.Item($i).RefreshTable() 
	}
}
$num_of_queries_to_delete = ($work.Queries).Count
Write-Host $num_of_queries_to_delete
for ($i = 1; $i -le $num_of_queries_to_delete; $i++) {
	$work.Queries.Item(1).Delete()
	Write-Host $i
}


$work.Save()
$work.Close()
$excel.Quit()
Remove-Variable excel
Remove-Variable work


$date = (Get-Date).ToString("yyyy-MM-dd")
$destination_excel_file = $location + "\Output\" + $date + " NexThink Reports.xlsb"

copy-item -path $Excel_file -Destination $destination_excel_file

####################################################################
##  UPDATE EXCEL FILE
####################################################################


####################################################################
##  RETRIEVE OLD DATA FROM EXCEL FILE
####################################################################
# Open Excel Application
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$work = $excel.Workbooks.Open($Excel_file)
$connections = $work.connections
# Refresh all tables related to the source file
$work.RefreshAll()
# Wait until all tables will be refreshed
while($connections | ForEach-Object {if($_.OLEDBConnection.Refreshing){$true}}){
	Start-Sleep -Milliseconds 500
}
Start-Sleep -seconds 10
$excel.DisplayAlerts = $false
if((test-path $Output_folder\merge)-eq $true){
Remove-item -path $Output_folder\merge -Confirm:$false -Recurse -Force
}
new-item -Type Directory $Output_folder\merge

if((test-path "C:\temp\MS_Teams_update_report\merge\new.csv")-eq $true){
Remove-item -path "C:\temp\MS_Teams_update_report\merge\new.csv"
}
$work.Sheets.Item("MS Teams Reinstall").SaveAs("C:\temp\MS_Teams_update_report\merge\new.csv", 6)
$work.Close()
$Output_reports = $location + "\Output\"
$Latest_excel = (Get-ChildItem -Path $Output_reports | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
$work = $excel.Workbooks.Open($Latest_excel)
if((test-path "C:\temp\MS_Teams_update_report\merge\old.csv")-eq $true){
Remove-item -path "C:\temp\MS_Teams_update_report\merge\old.csv"
}
$work.Sheets.Item("MS Teams campaign").SaveAs("C:\temp\MS_Teams_update_report\merge\old.csv", 6)
$work.Close()

####################################################################
##  RETRIEVE OLD DATA FROM EXCEL FILE
####################################################################

####################################################################
##  Merge CSV files
####################################################################

if((test-path "$Output_folder\Ready_output.csv")-eq $true){
Remove-item -path "$Output_folder\Ready_output.csv"
}
$new = Import-Csv -Path $Output_folder\merge\new.csv
$old = Import-Csv -Path $Output_folder\merge\old.csv
$ready_out = $old + $new
$ready_out = ($ready_out | Group-Object device/name | ForEach-Object {$_.Group | Select-Object -First 1 })
$ready_out |Export-Csv -Path $Output_folder\Ready_output.txt -NoTypeInformation

####################################################################
##  Merge CSV files
####################################################################