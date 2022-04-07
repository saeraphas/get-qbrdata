<#
.SYNOPSIS
	This script collects data from a Windows AD domain controller and builds a number of reports intended for customer review as part of periodic housekeeping, license audits, true-ups, etc.

.DESCRIPTION
	This script depends on WMF 4 and the ActiveDirectory Powershell module,  and is intended to run from a domain controller, though any domain member server or workstation with the RSAT-AD-PowerShell module installed should work.
	
.EXAMPLE
	.\Get-QBRData.ps1

.NOTES
    Author:             Douglas Hammond (douglas@douglashammond.com)
	License: 			This script is distributed under "THE BEER-WARE LICENSE" (Revision 42):
						As long as you retain this notice you can do whatever you want with this stuff.
						If we meet some day, and you think this stuff is worth it, you can buy me a beer in return.
#>

#Requires -Version 4.0
#Requires -Module activedirectory
Write-Progress -Id 0 -Activity "Initializing variables."

$contact = "Nexigen Communications, LLC"
$contactlink = "https://www.nexigen.com"
#$contactlogo 	= "https://www.nexigen.com/wp-content/themes/nexigen/library/images/nexigen-logo.svg"
#$contactlogo 	= "https://www.nexigen.com/files/2021/09/logo-min.png"
$contactlogo = "https://149698627.v2.pressablecdn.com/wp-content/uploads/2021/12/Logo-White-BG.png"
$contactmail = "mailto:help@nexigen.com"
$date = (Get-Date -DisplayHint Date).DateTime | Out-String
$outputpath = "C:\Nexigen\"
$outputprefix = "nex-sbr-"
$reportingby = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$reportingfrom = ([System.Net.Dns]::GetHostByName(($env:computerName))).Hostname
$zipoutput = $true
$headerdetail = "Report data generated by user $reportingby on $date from server $reportingfrom."
$footerdetail = "For questions or additional information, please <a href=`"$contactmail`">contact $contact</a>."

#super basic CSS to pretty up HTML reports, reports use -replace to customize strings
$HTMLPre = @"
<head>
<style>
.report {
  padding: 3%
}
.logo {
  width: 180px;
  height: auto;
  max-height: 150px;
}
.logo img {
  width: 100%;
  height: auto;
  max-height: 150px;
}
h1 {
  padding-left: 0;
  color: blue;
  font-family: -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Helvetica Neue",Arial,sans-serif,"Apple Color Emoji","Segoe UI Emoji","Segoe UI Symbol";
  font-size: 150%;
  color: #005a87;
  webkit-font-smoothing: antialiased;
  margin-bottom: .25em;
}
p {
  padding-left: 2em;
  color: #343a40;
  font-family: "Open Sans",sans-serif;
  font-size: 100%;
  webkit-font-smoothing: antialiased;
  margin-top: .5em;
  margin-bottom: .5em;
}
table {
  border: 1px solid black;
  color: #454E54;
  font-family: SFMono-Regular,Menlo,Monaco,Consolas,"Liberation Mono","Courier New",monospace;
  font-size: 75%; 
  webkit-font-smoothing: antialiased;  
}
tr:nth-of-type(odd) {
  background-color: #e5e5e5;
}
tr:first-child {
  color: #fff;
  background-color: #007bff;
}
</style>
</head>
<div class="report">
<div class="logo">
<logo><a href="$contactlink"><img src="$contactlogo" alt="$contact Logo" /></a></logo>
</div>
<h1>REPORTTITLE</h1>
<p>REPORTSUBTITLE</p>
<p>$headerdetail</p>
<hr>
"@

$HTMLPost = @"
<hr><p>$footerdetail</p></div>
"@

function New-Report() {
	Param($ReportName, $Title, $Subtitle, $ReportData)
	Write-Progress -Id 0 -Activity "Collecting report data." -CurrentOperation "Collecting $Title."
	$HTMLPrefixed = $HTMLPre -replace "REPORTTITLE", "$Title" -replace "REPORTSUBTITLE", "$Subtitle"
	$ReportOutput = $outputpath + $outputprefix + $ReportName
	$XLSreport = $outputpath + $outputprefix + "combined.xlsx"
	if (!$reportdata) {
		$reportdata = @()
		$row = New-Object PSObject
		$row | Add-Member -MemberType NoteProperty -Name "Result" -Value "This report is empty."
		$reportdata += $row
	}
	$reportdata | ConvertTo-Html -Title "$title" -PreContent $HTMLPrefixed -post $HTMLPost | out-file -filepath "$reportoutput.html"

	if ($skipExcel -ne $true) {
		$reportdata | Export-Excel `
			-Path $XLSreport `
			-WorkSheetname "$title" `
			-ClearSheet `
			-BoldTopRow `
			-Autosize `
			-FreezePane 2 `
			-Autofilter `
	
 }
 else {
		$reportdata | export-csv "$reportoutput.csv" -notypeinformation		
	}
	$reportdata = $null
}

function Get-Cert( $computer = $env:computername ) {
	$ro = [System.Security.Cryptography.X509Certificates.OpenFlags]"ReadOnly"
	$lm = [System.Security.Cryptography.X509Certificates.StoreLocation]"LocalMachine"
	#    $store=new-object System.Security.Cryptography.X509Certificates.X509Store("\\$computer\root",$lm)
	$store = new-object System.Security.Cryptography.X509Certificates.X509Store("\\$computer\My", $lm)
	$store.Open($ro)
	$store.Certificates
}

Write-Progress -Id 0 -Activity "Checking prerequisites."

If (!(test-path $outputpath)) { New-Item -ItemType Directory -Force -Path $outputpath }

#try to install required modules
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
If (!(Get-Module -ListAvailable -Name activedirectory)) { try { Install-Module activedirectory -scope CurrentUser -Force } catch { Write-Error "An error occurred adding the ActiveDirectory Powershell module. Unable to continue."; exit } }
If (!(Get-Module -ListAvailable -Name NuGet)) { try { Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force } catch { Write-Warning "An error occurred adding the NuGet provider." } }
If (!(Get-Module -ListAvailable -Name ImportExcel)) { try { Install-Module ImportExcel -scope CurrentUser -Force } catch { Write-Warning "An error occurred adding the ImportExcel Powershell module. Excel-formatted reports will not be available."; $skipExcel = $true } }

#bail out if we can't load them
If (Get-Module -ListAvailable -Name activedirectory) { try { import-module activedirectory } catch { Write-Error "An error occurred importing the ActiveDirectory Powershell module. Unable to continue."; exit } }
If (Get-Module -ListAvailable -Name ImportExcel) { try { import-module ImportExcel } catch { Write-Warning "An error occurred importing the ImportExcel Powershell module. Excel-formatted reports will not be available."; $skipExcel = $true } }

Write-Progress -Id 0 -Activity "Collecting report data."

# Get AD user accounts and logon dates
$ReportName = "usersaudit"
$Title = "User Account Audit Report"
$Subtitle = "All enabled and disabled accounts in this domain. </br>Last logon date is reported by a single domain controller and may not be 100% accurate."
$reportdata = Get-ADUser -Filter * -Properties Name, Description, lastlogondate, passwordlastset, enabled | select-object -property name, distinguishedname, lastlogondate, @{N='Days Since Last Logon'; E={(new-timespan -start $(Get-date $_.LastLogondate) -end (get-date)).days}}, passwordlastset, enabled | Sort-Object -Property enabled, name, lastlogondate
New-Report -ReportName $ReportName -Title $Title -Subtitle $Subtitle -ReportData $reportdata

# Get inactive users 
$ReportName = "inactiveusers"
$Title = "Inactive Users Report"
$inactivitythreshold = 365
$Subtitle = "User accounts that have not logged on to Active Directory in ~$($inactivitythreshold) days or more."
$inactivitypad = $inactivitythreshold + 15 #pad this date by 15 days because this attrtibute is only replicated periodically
$inactivitydate = (get-date).AddDays(-$inactivitypad) 
$reportdata = Get-ADUser -Filter { (LastLogonDate -lt $inactivitydate) -and (enabled -eq $true) } -properties LastLogonDate, passwordlastset | Select-Object Name, LastLogonDate, passwordlastset | Sort-Object -Property name, lastlogondate
$reportoutput = $outputpath + $outputprefix + "inactiveusers.$outputtype"
New-Report -ReportName $ReportName -Title $Title -Subtitle $Subtitle -ReportData $reportdata

# Get inactive servers
$ReportName = "inactiveservers"
$Title = "Offline Servers Report"
$Subtitle = "Servers that do not respond to ICMP or SMB. This report may be empty."
$ServersOnline = @()
$ServersOffline = @()
$Servers = Get-ADComputer -Filter { OperatingSystem -Like '*Windows Server*' } -Properties OperatingSystem, enabled | Where-Object { $_.Enabled -eq $True } | Select-Object -ExpandProperty Name
$TimeoutMillisec = 3000

Foreach ($Server in $Servers) {
	$PingStatus = Get-WmiObject -Class Win32_PingStatus -Filter "(Address='$server') and timeout=$TimeoutMillisec"
	$SMBStatus = start-job { test-path -path "\\$args\c$" } -ArgumentList $Server | wait-job -timeout 1 | Receive-Job
	# Construct an object
	$myobj = "" | Select-Object "Server", "PingStatus", "SMBStatus"

	# Fill the object
	$myobj.Server = $Server
	$myobj.PingStatus = $PingStatus.StatusCode
	$myobj.SMBStatus = $SMBStatus

	# Add the object to the out-array
	If (($PingStatus.StatusCode -eq 0) -or ($SMBStatus -eq $true)) { $ServersOnline += $myobj } else { $ServersOffline += $myobj }

	# Clear the object
	$myobj = $null
	$pingstatus = $null
	$smbstatus = $null
}
$reportdata = $ServersOffline | Select-Object -Property Server, PingStatus, SMBStatus | Sort-Object -Property Server
New-Report -ReportName $ReportName -Title $Title -Subtitle $Subtitle -ReportData $reportdata

# Get inactive computers as selected output type
$ReportName = "inactivepcs"
$Title = "Inactive Computers Report"
$Subtitle = "Computer accounts that have not logged on to Active Directory in ~180 days or more."
$reportdata = search-adaccount -accountinactive -computersonly -timespan "195" | Where-Object { $_.enabled } | select-object -property name, distinguishedname, lastlogondate | Sort-Object -Property lastlogondate, name
New-Report -ReportName $ReportName -Title $Title -Subtitle $Subtitle -ReportData $reportdata

# Get domain admins
$ReportName = "domainadmins"
$Title = "Domain Administrators Report"
$Subtitle = "Accounts with Domain Administrator permissions."
$reportdata = Get-ADGroupMember -Identity 'Domain Admins' | Get-ADObject -Properties Name, distinguishedname, objectclass, Description | select-object -property name, distinguishedname, objectclass, description | Sort-Object -Property name
New-Report -ReportName $ReportName -Title $Title -Subtitle $Subtitle -ReportData $reportdata

# Get server disk space 
$ReportName = "diskfreespace"
$Title = "Server Storage Report"
$Subtitle = "Storage utilizaion on Windows Servers."
# but not if we're the local system account
if ($reportingby -ne "NT AUTHORITY\SYSTEM") {
	#	$Servers 	= Get-ADComputer -Filter { OperatingSystem -Like '*Windows Server*' } -Properties OperatingSystem,enabled | Where { $_.Enabled -eq $True} | select -ExpandProperty Name
	$Servers = $ServersOnline | Select-Object -ExpandProperty Server
	$reportdata = Get-WmiObject Win32_LogicalDisk -ComputerName $Servers -Filter "DriveType='3'" -ErrorAction SilentlyContinue | Select-Object PsComputerName, DeviceID, @{N = "Disk Size (GB) "; e = { [math]::Round($($_.Size) / 1073741824, 2) } }, @{N = "Free Space (GB)"; e = { [math]::Round($($_.FreeSpace) / 1073741824, 2) } }, @{N = "Used Space (%)"; e = { [math]::Round($($_.Size - $_.FreeSpace) / $_.Size * 100, 1) } }, @{N = "Used Space (GB)"; e = { [math]::Round($($_.Size - $_.FreeSpace) / 1073741824, 2) } } 
	New-Report -ReportName $ReportName -Title $Title -Subtitle $Subtitle -ReportData $reportdata
}
else {
	Write-Warning "Skipped collecting $Title. This report cannot run as $reportingby."
}

# Get service accounts 
$ReportName = "serviceaccounts"
$Title = "Service Accounts Report"
$Subtitle = "Windows Services using a custom Log On As account. </br>This report may be empty."
# but not if we're the local system account
if ($reportingby -ne "NT AUTHORITY\SYSTEM") {
	#	$Servers 	= Get-ADComputer -Filter { OperatingSystem -Like '*Windows Server*' } -Properties OperatingSystem,enabled | Where { $_.Enabled -eq $True} | select -ExpandProperty Name
	$Servers = $ServersOnline | Select-Object -ExpandProperty Server
	$reportdata = Get-WmiObject Win32_Service -ComputerName $Servers -Filter "not StartMode='Disabled'" -ErrorAction SilentlyContinue | Select-Object PsComputerName, Name, StartName | Where-Object -Property StartName -notlike "" | Where-Object -Property StartName -notmatch "LocalSystem" | Where-Object -Property StartName -notmatch "LocalService" | Where-Object -Property StartName -notmatch "NetworkService" | Sort-Object -Property pscomputername
	New-Report -ReportName $ReportName -Title $Title -Subtitle $Subtitle -ReportData $reportdata

}
else {
	Write-Warning "Skipped collecting $Title. This report cannot run as $reportingby."
}

# Get static nameservers on server interfaces
$ReportName = "nameservers"
$Title = "Static DNS servers"
$Subtitle = "DNS server addresses in use on Windows Servers. </br>This report may be empty."
# but not if we're the local system account
if ($reportingby -ne "NT AUTHORITY\SYSTEM") {
	#	$Servers 	= Get-ADComputer -Filter { OperatingSystem -Like '*Windows Server*' } -Properties OperatingSystem,enabled | Where { $_.Enabled -eq $True} | select -ExpandProperty Name
	$Servers = $ServersOnline | Select-Object -ExpandProperty Server
	$reportdata = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $Servers -Filter "IPEnabled=TRUE" -ErrorAction SilentlyContinue | Where-Object { $_.DNSServerSearchOrder -ne $null } | Select-Object PsComputerName, @{Name = 'Nameservers'; Expression = { [string]::join("; ", ($_.DnsServerSearchOrder)) } } | Sort-Object -Property pscomputername
	New-Report -ReportName $ReportName -Title $Title -Subtitle $Subtitle -ReportData $reportdata
}
else {
	Write-Warning "Skipped collecting $Title. This report cannot run as $reportingby."
}

# Get file and print shares 
$ReportName = "fileshares"
$Title = "Network shares"
$Subtitle = "SMB shares on Windows Servers."
# but not if we're the local system account
if ($reportingby -ne "NT AUTHORITY\SYSTEM") {
	#	$Servers 	= Get-ADComputer -Filter { OperatingSystem -Like '*Windows Server*' } -Properties OperatingSystem,enabled | Where { $_.Enabled -eq $True} | select -ExpandProperty Name
	$Servers = $ServersOnline | Select-Object -ExpandProperty Server
	$reportdata = Get-WmiObject -Class Win32_Share -ComputerName $Servers -ErrorAction SilentlyContinue | Select-Object PsComputerName, Name, Path, Description | Sort-Object -Property pscomputername 
	New-Report -ReportName $ReportName -Title $Title -Subtitle $Subtitle -ReportData $reportdata
}
else {
	Write-Warning "Skipped collecting $Title. This report cannot run as $reportingby."
}

# Get SSL certificates
$ReportName = "sslcertificates"
$Title = "SSL Certificates"
$Subtitle = "SSL Certificates on servers in this domain. This report may be empty. "
# but not if we're the local system account
if ($reportingby -ne "NT AUTHORITY\SYSTEM") {
	#	$Servers 	= Get-ADComputer -Filter { OperatingSystem -Like '*Windows Server*' } -Properties OperatingSystem,enabled | Where { $_.Enabled -eq $True} | select -ExpandProperty Name
	$Servers = $ServersOnline | Select-Object -ExpandProperty Server
	#	$reportdata = Get-Cert $Servers -ErrorAction SilentlyContinue | ?{$_.Subject -ne $_.Issuer} | ?{$_.NotAfter -gt (Get-Date)} | ?{$_.NotAfter -lt (Get-Date).AddDays(365)} | format-list -property thumbprint,NotAfter,Subject,Issuer
	$reportdata = @()
	foreach ( $Server in $Servers ) {
		$certificates = $null
		#		try {$certificates = Get-Cert "$Server" -ErrorAction SilentlyContinue | ?{$_.Subject -ne $_.Issuer} | ?{$_.NotAfter -gt (Get-Date)} | ?{$_.NotAfter -lt (Get-Date).AddDays(365)} | Select-Object -property thumbprint,NotAfter,Subject,Issuer} catch {Write-Warning "An error occurred collecting $ReportName data from $Server."; Continue}
		try { $certificates = Get-Cert "$Server" -ErrorAction SilentlyContinue | Select-Object -property Subject, Issuer, thumbprint, NotAfter } catch { Write-Warning "An error occurred collecting $ReportName data from $Server."; Continue }
		foreach ($certificate in $certificates) {
			$certHash = $null
			$certHash = [ordered]@{
				'Server'     = $Server
				'Subject'    = ($certificate.subject | Out-String).Trim()
				'ExpiryDate' = $certificate.notafter
				'ExpiryDays' = (New-Timespan -Start $(Get-Date) -End $($certificate.notafter)).Days
				'Thumbprint' = ($certificate.thumbprint | Out-String).Trim()
				'Issuer'     = ($certificate.issuer | Out-String).Trim()
			}
			$certObject = $null
			$certObject = New-Object PSObject -Property $certHash
			$reportdata += $certObject
		}	
	}
	$reportdata = $reportdata | sort-object -Property ExpiryDays
	New-Report -ReportName $ReportName -Title $Title -Subtitle $Subtitle -ReportData $reportdata
}
else {
	Write-Warning "Skipped collecting $Title. This report cannot run as $reportingby."
}

# Get custom Active Directory Groups and their users 
# this will error out on groups over 5000 users until I rewrite it to use Get-ADUser -LDAPFilter
$ReportName = "domaingroups"
$Title = "Active Directory Groups Report"
$Subtitle = "Groups specific to this organization and their members. </br>Default Built-in groups are excluded."
$Groups = Get-ADGroup -Filter { GroupCategory -eq "Security" -and GroupScope -eq "Global" } -Properties isCriticalSystemObject, distinguishedname | Where-Object { !($_.IsCriticalSystemObject) } | select-object DistinguishedName, Name
#$reportdata 	= foreach( $Group in $Groups ){Get-ADGroupMember -Identity $Group | foreach {[pscustomobject]@{GroupName = $Group.Name;Name = $_.Name}}}
$reportdata = foreach ( $Group in $Groups ) { Get-ADUser -LDAPFilter "(&(objectCategory=user)(memberof=$($group.distinguishedname)))" | ForEach-Object { [pscustomobject]@{GroupName = $Group.Name; Name = $_.Name } } }
New-Report -ReportName $ReportName -Title $Title -Subtitle $Subtitle -ReportData $reportdata

#get EOL PC list and last known IP address
#note: win10 build list from here https://docs.microsoft.com/en-us/windows/release-information/
$ReportName = "eospcs"
$Title = "End-of-Support PCs Report"
$Subtitle = "Computer accounts in Active Directory with end-of-support operating systems. </br>Old Win10 builds and feature updates are also included."
$reportdata = Get-ADComputer -Filter 'operatingsystem -notlike "*server*" -and enabled -eq "true"' -Properties Name, Operatingsystem, OperatingSystemVersion, LastLogonDate, IPv4Address | Where-Object { $_.OperatingSystem -imatch "Windows 10|Windows Vista|Windows XP|95|94|Windows 8|2000|2003|Windows NT|Windows 7" -and $_.OperatingSystemVersion -inotmatch "6.3.9600|6.1.7601|19044|19043|19042|19041" } | Select-Object -Property Name, Operatingsystem, OperatingSystemVersion, LastLogonDate, IPv4Address | Sort-Object -Property operatingsystemversion, name
New-Report -ReportName $ReportName -Title $Title -Subtitle $Subtitle -ReportData $reportdata

#get EOL server list and last known IP address
#note: win10 build list from here https://docs.microsoft.com/en-us/windows/release-information/
$ReportName = "eosservers"
$Title = "End-of-Support Servers Report"
$Subtitle = "Server accounts in Active Directory with end-of-support operating systems"
$reportdata = Get-ADComputer -Filter 'operatingsystem -like "*server*" -and enabled -eq "true"' -Properties Name, Operatingsystem, OperatingSystemVersion, LastLogonDate, IPv4Address | Where-Object { $_.OperatingSystem -imatch "Windows NT|2000|2003|2008" } | Select-Object -Property Name, Operatingsystem, OperatingSystemVersion, LastLogonDate, IPv4Address | Sort-Object -Property operatingsystemversion, name
New-Report -ReportName $ReportName -Title $Title -Subtitle $Subtitle -ReportData $reportdata

Write-Progress -Id 0 -Activity "Collecting report data." -Status "Complete."

If ($zipoutput = $true) {
	#create scratch directory and move output files there
	Write-Progress -Id 1 -Activity "Compressing report data."
	Write-Progress -Id 1 -Activity "Compressing report data." -Status "Creating ZIP working directory."
	$scratchpath = $outputpath + "scratch\"
	If (!(Test-Path -LiteralPath $scratchpath)) { New-Item -Path $scratchpath -ItemType Directory -ErrorAction Stop | Out-Null }
	Get-ChildItem -Path $outputpath $outputprefix*.* | Where-Object { ! $_.PSIsContainer } | Where-Object { $_.Extension -ne ".zip" } | Move-Item -Destination $scratchpath 

	Write-Progress -Id 1 -Activity "Compressing report data." -Status "Adding files to ZIP."
	#zip scratch to output using powershell v4 method
	#	$destinationZipFileName = $outputpath + "QBRData.zip"
	$destinationZipFileName = $outputpath + $outputprefix + "report-bundle-" + $((get-date).tostring("yyyy-MM-dd")) + ".zip"
	If (Test-Path -LiteralPath $destinationZipFileName) {
		Write-Warning "ZIP file $destinationZipFileName already exists. Replacing old file."
		Remove-Item -Path $destinationZipFileName -Force
	}
	[Reflection.Assembly]::LoadWithPartialName("System.IO.Compression.FileSystem") | Out-Null
	[System.IO.Compression.ZipFile]::CreateFromDirectory($scratchpath, $destinationZipFileName) | Out-Null
	Write-Progress -Id 1 -Activity "Compressing report data." -Status "ZIP file $destinationZipFileName creation finished."

	If (Test-Path -LiteralPath $destinationZipFileName) {
		#remove the scratch directories
		Write-Progress -Id 1 -Activity "Compressing report data." -Status "Removing ZIP working directory."
		Remove-Item -Path $scratchpath -recurse -force
	
		Write-Progress -Id 1 -Activity "Compressing report data." -Status "Complete."
	}

}

Write-Host "Done."