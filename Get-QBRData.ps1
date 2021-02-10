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

$owner 			= "Nexigen Communications, LLC"
$ownerlink 		= "https://www.nexigen.com"
$ownerlogo 		= "https://www.nexigen.com/wp-content/themes/nexigen/library/images/nexigen-logo.svg"
$ownermail 		= "mailto:help@nexigen.com"
$date 			= (Get-Date -DisplayHint Date).DateTime | Out-String
$outputpath 	= "C:\Nexigen\"
$outputprefix 	= "nex-sbr-"
$reportingby 	= [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$reportingfrom 	= ([System.Net.Dns]::GetHostByName(($env:computerName))).Hostname
$zipoutput 		= $true
$headerdetail 	= "Report data generated by $reportingby on $date from $reportingfrom."
$footerdetail 	= "For questions or additional information, please <a href=`"$ownermail`">contact $owner</a>."

#super basic CSS to pretty up HTML reports, reports use -replace to customize strings
$HTMLPre = @"
<head>
<style>
.report {
  padding: 3%
}
.logo {
  width: 200px;
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
  font-family: verdana;
  font-size: 150%;
  line-height: .5em;
}
h2 {
  padding-left: 2em;
  color: blue;
  font-family: verdana;
  font-size: 100%;
  line-height: .5em;
}
p {
  color: gray;
  font-family: courier;
  font-size: 100%;  
}
</style>
</head>
<div class="report">
<div class="logo">
<logo><a href="$ownerlink"><img src="$ownerlogo" alt="$owner Logo" /></a></logo>
</div>
<h1>REPORTTITLE</h1>
<h2>REPORTSUBTITLE</h2>
<p>$headerdetail</P>
<hr>
"@

$HTMLPost 		= @"
<hr><p>$footerdetail</p></div>
"@

function New-Report(){
	Param($ReportName, $Title, $Subtitle, $ReportData)
	Write-Progress -Id 0 -Activity "Collecting report data." -CurrentOperation "Collecting $Title."
	$HTMLPrefixed = $HTMLPre -replace "REPORTTITLE", "$Title" -replace "REPORTSUBTITLE", "$Subtitle"
	$ReportOutput = $outputpath + $outputprefix + $ReportName
	if (!$reportdata){
		$reportdata = @()
			$row = New-Object PSObject
			$row | Add-Member -MemberType NoteProperty -Name "Result" -Value "This report is empty."
			$reportdata += $row
		}
	$reportdata | ConvertTo-Html -Title "$title" -PreContent $HTMLPrefixed -post $HTMLPost | out-file -filepath "$reportoutput.html"
	$reportdata | export-csv "$reportoutput.csv" -notypeinformation
}

If(!(test-path $outputpath)){New-Item -ItemType Directory -Force -Path $outputpath }

import-module activedirectory

Write-Progress -Id 0 -Activity "Collecting report data."

# Get AD user accounts and logon dates
$ReportName 	= "usersaudit"
$Title 			= "User Account Audit Report"
$Subtitle 		= "User accounts in this domain. Last logon date is reported by single domain controller and may not be 100% accurate."
$reportdata 	= Get-ADUser -Filter 'enabled -eq "true"' -Properties Name,Description,lastlogondate,passwordlastset | select-object -property name,distinguishedname,lastlogondate,passwordlastset | Sort-Object -Property lastlogondate,name
New-Report -ReportName $ReportName -Title $Title -Subtitle $Subtitle -ReportData $reportdata

# Get inactive users 
$ReportName 	= "inactiveusers"
$Title 			= "Inactive Users Report"
$Subtitle 		= "User accounts that have not logged on to Active Directory in 180 days or more."
$reportdata 	= search-adaccount -accountinactive -usersonly -timespan "195" | where {$_.enabled} | select-object -property name,distinguishedname,lastlogondate | Sort-Object -Property lastlogondate,name
$reportoutput 	= $outputpath + $outputprefix + "inactiveusers.$outputtype"
New-Report -ReportName $ReportName -Title $Title -Subtitle $Subtitle -ReportData $reportdata

# Get inactive computers as selected output type
$ReportName 	= "inactivepcs"
$Title 			= "Inactive Computers Report"
$Subtitle 		= "Computer accounts that have not logged on to Active Directory in 180 days or more."
$reportdata 	= search-adaccount -accountinactive -computersonly -timespan "195" | where {$_.enabled} | select-object -property name,distinguishedname,lastlogondate | Sort-Object -Property lastlogondate,name
New-Report -ReportName $ReportName -Title $Title -Subtitle $Subtitle -ReportData $reportdata

# Get domain admins
$ReportName 	= "domainadmins"
$Title 			= "Domain Administrators Report"
$Subtitle 		= "Active accounts with Domain Administrator permissions"
$reportdata 	= Get-ADGroupMember -Identity 'Domain Admins' | Get-ADObject -Properties Name,distinguishedname,objectclass,Description | select-object -property name,distinguishedname,objectclass,description | Sort-Object -Property name
New-Report -ReportName $ReportName -Title $Title -Subtitle $Subtitle -ReportData $reportdata

# Get server disk space 
$ReportName 	= "diskfreespace"
$Title 			= "Server Storage Report"
$Subtitle 		= "Hard drive space on servers"
# but not if we're the local system account
if ($reportingby -ne "NT AUTHORITY\SYSTEM")
{
	$Servers 	= Get-ADComputer -Filter { OperatingSystem -Like '*Windows Server*' } -Properties OperatingSystem,enabled | Where { $_.Enabled -eq $True} | select -ExpandProperty Name
	$reportdata = Get-WmiObject Win32_LogicalDisk -ComputerName $Servers -Filter "DriveType='3'" -ErrorAction SilentlyContinue | Select-Object PsComputerName, DeviceID, @{N="Disk Size (GB) ";e={[math]::Round($($_.Size) / 1073741824,2)}}, @{N="Free Space (GB)";e={[math]::Round($($_.FreeSpace) / 1073741824,2)}}, @{N="Used Space (%)";e={[math]::Round($($_.Size - $_.FreeSpace) / $_.Size * 100,1)}}, @{N="Used Space (GB)";e={[math]::Round($($_.Size - $_.FreeSpace) / 1073741824,2)}} 
	New-Report -ReportName $ReportName -Title $Title -Subtitle $Subtitle -ReportData $reportdata
}else {
Write-Warning "Skipped collecting $Title. This report cannot run as $reportingby."
}

# Get service accounts 
$ReportName 	= "serviceaccounts"
$Title 			= "Service Accounts Report"
$Subtitle 		= "Windows Services using a custom Log On As account. This report may be empty."
# but not if we're the local system account
if ($reportingby -ne "NT AUTHORITY\SYSTEM")
{
	$Servers 	= Get-ADComputer -Filter { OperatingSystem -Like '*Windows Server*' } -Properties OperatingSystem,enabled | Where { $_.Enabled -eq $True} | select -ExpandProperty Name
	$reportdata = Get-WmiObject Win32_Service -ComputerName $Servers -Filter "not StartMode='Disabled'" -ErrorAction SilentlyContinue | Select-Object PsComputerName, Name, StartName | Where -Property StartName -notlike "" | Where -Property StartName -notmatch "LocalSystem" | Where -Property StartName -notmatch "LocalService" | Where -Property StartName -notmatch "NetworkService" | Sort-Object -Property pscomputername
	New-Report -ReportName $ReportName -Title $Title -Subtitle $Subtitle -ReportData $reportdata

}else {
Write-Warning "Skipped collecting $Title. This report cannot run as $reportingby."
}

# Get static nameservers on server interfaces
$ReportName 	= "nameservers"
$Title 			= "Static DNS servers"
$Subtitle 		= "Windows Servers using static DNS addresses. This report may be empty."
# but not if we're the local system account
if ($reportingby -ne "NT AUTHORITY\SYSTEM")
{
	$Servers 	= Get-ADComputer -Filter { OperatingSystem -Like '*Windows Server*' } -Properties OperatingSystem,enabled | Where { $_.Enabled -eq $True} | select -ExpandProperty Name
	$reportdata = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $Servers -Filter "IPEnabled=TRUE" -ErrorAction SilentlyContinue | where {$_.DNSServerSearchOrder -ne $null} | Select-Object PsComputerName,@{Name='Nameservers';Expression={[string]::join("; ", ($_.DnsServerSearchOrder))}} | Sort-Object -Property pscomputername
	New-Report -ReportName $ReportName -Title $Title -Subtitle $Subtitle -ReportData $reportdata
}else {
Write-Warning "Skipped collecting $Title. This report cannot run as $reportingby."
}

# Get custom Active Directory Groups and their users 
# this will error out on groups over 5000 users until I rewrite it to use Get-ADUser -LDAPFilter
$ReportName 	= "domaingroups"
$Title 			= "Active Directory Groups Report"
$Subtitle 		= "Groups specific to this organization and their members. Default Built-in groups are excluded."
$Groups 		= Get-ADGroup -Filter { GroupCategory -eq "Security" -and GroupScope -eq "Global"  } -Properties isCriticalSystemObject | Where-Object { !($_.IsCriticalSystemObject)}
$reportdata 	= foreach( $Group in $Groups ){Get-ADGroupMember -Identity $Group | foreach {[pscustomobject]@{GroupName = $Group.Name;Name = $_.Name}}}
New-Report -ReportName $ReportName -Title $Title -Subtitle $Subtitle -ReportData $reportdata

#get EOL PC list and last known IP address
#note: win10 build list from here https://docs.microsoft.com/en-us/windows/release-information/
$ReportName 	= "eospcs"
$Title 			= "End-of-Support PCs Report"
$Subtitle 		= "Computer accounts in Active Directory with end-of-support operating systems"
$reportdata 	= Get-ADComputer -Filter 'operatingsystem -notlike "*server*" -and enabled -eq "true"' -Properties Name,Operatingsystem,OperatingSystemVersion,LastLogonDate,IPv4Address | Where {$_.OperatingSystem -imatch "Windows 10|Windows Vista|Windows XP|95|94|Windows 8|2000|2003|Windows NT|Windows 7" -and $_.OperatingSystemVersion -inotmatch "6.3.9600|6.1.7601|19042|19041|18363|17763|17134|14393"} | Select-Object -Property Name,Operatingsystem,OperatingSystemVersion,LastLogonDate,IPv4Address | Sort-Object -Property operatingsystemversion,name
New-Report -ReportName $ReportName -Title $Title -Subtitle $Subtitle -ReportData $reportdata

Write-Progress -Id 0 -Activity "Collecting report data." -Status "Complete."

If ($zipoutput = $true){
	#create scratch directory and move output files there
	Write-Progress -Id 1 -Activity "Compressing report data."
	Write-Progress -Id 1 -Activity "Compressing report data." -Status "Creating ZIP working directory."
	$scratchpath = $outputpath + "scratch\"
	If (!(Test-Path -LiteralPath $scratchpath)){New-Item -Path $scratchpath -ItemType Directory -ErrorAction Stop | Out-Null}
	Get-ChildItem -Path $outputpath $outputprefix*.* | Move-Item -Destination $scratchpath 

	Write-Progress -Id 1 -Activity "Compressing report data." -Status "Adding files to ZIP."
	#zip scratch to output using powershell v4 method
	$destinationZipFileName = $outputpath + "QBRData.zip"
	If (Test-Path -LiteralPath $destinationZipFileName){
		Write-Warning "ZIP file $destinationZipFileName already exists. Replacing old file."
		Remove-Item -Path $destinationZipFileName -Force
		}
	[Reflection.Assembly]::LoadWithPartialName("System.IO.Compression.FileSystem") | Out-Null
	[System.IO.Compression.ZipFile]::CreateFromDirectory($scratchpath, $destinationZipFileName) | Out-Null
	Write-Progress -Id 1 -Activity "Compressing report data." -Status "ZIP file $destinationZipFileName creation finished."

	If (Test-Path -LiteralPath $destinationZipFileName){
	#remove the scratch directories
	Write-Progress -Id 1 -Activity "Compressing report data." -Status "Removing ZIP working directory."
	Remove-Item -Path $scratchpath -recurse -force
	
	Write-Progress -Id 1 -Activity "Compressing report data." -Status "Complete."
	}

}

Write-Host "Done."