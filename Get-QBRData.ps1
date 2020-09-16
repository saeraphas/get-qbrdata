<#
.SYNOPSIS
	Pull housekeeping reports from a domain controller for periodic review. 
.DESCRIPTION
	This script produces reports intended for customer review as part of 
	periodic housekeeping, license audits, true-ups, etc. 
	
	These reports can be generated in HTML or CSV format - most of the HTML 
	reports print out nicely and work well for print>strikethrough>scanback.
	CSV is for doing fancy stuff. 	
	
	This script depends on the ActiveDirectory module and is intended to run
	from a domain controller, though any member server or workstation with the
	RSAT-AD-PowerShell module installed will work. 
	#Install-WindowsFeature RSAT-AD-PowerShell
	
	It has been tested on Windows Server 2008R2 with WMF 4 and higher though
	other platforms may also work. I'm not a real developer. 

	When run with a local SYSTEM account, this script produces
	- inactive users report
	- inactive computers report
	- domain admin group membership report
	- custom AD group membership report
		
	When run with a domain admin, this script additionally produces
	- server storage space report 
	- service accounts report
		
.EXAMPLE
	.\Get-QBRData.ps1

.NOTES
    Author:             Douglas Hammond 
    Changelog:
        2020-08-28		Added script header for first github version. 
        2020-09-02		Added check for NT AUTHORITY\SYSTEM to prevent ugly errors on storage report. 
		2020-09-02		Added check to remove old output ZIP before creating a new one. 
		2020-09-14		Added services report, Expanded customization of HTML header
		2020-09-16		Fixed services report, output works as intended now.
#>

#DQDQBRDC - Douglas' Quick and Dirty QBR Data Collector
#Requires -Version 4.0
#Requires -Module activedirectory

$owner 			= "Nexigen Communications, LLC"
$ownerlink 		= "https://www.nexigen.com"
$ownerlogo 		= "https://www.nexigen.com/wp-content/themes/nexigen/library/images/nexigen-logo.svg"
$ownermail 		= "mailto:help@nexigen.com"
$date 			= (Get-Date -DisplayHint Date).DateTime | Out-String
$outputpath 	= "C:\Nexigen\"
$outputprefix 	= "nex-sbr-"
$outputtype 	= "HTML" #CSV or HTML
$inactiveusers 	= $outputpath + $outputprefix + "inactiveusers.$outputtype"
$inactivepcs 	= $outputpath + $outputprefix + "inactivepcs.$outputtype"
$domainadmins 	= $outputpath + $outputprefix + "domainadmins.$outputtype"
$ADGroupReport	= $outputpath + $outputprefix + "adgroups.$outputtype"
$diskfreespace 	= $outputpath + $outputprefix + "diskfreespace.$outputtype"
$serviceaccts 	= $outputpath + $outputprefix + "serviceaccounts.$outputtype"
#$reportingby 	= $env:UserName
$reportingby 	= [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$reportingfrom 	= ([System.Net.Dns]::GetHostByName(($env:computerName))).Hostname
$zipoutput 		= $true
$headerdetail 	= "Report data generated by $reportingby on $date from $reportingfrom."
$footerdetail 	= "For questions or additional information, please <a href=`"$ownermail`">contact $owner</a>."


#what if i use a here-string instead?
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

$HTMLPost 		= "<hr><p>$footerdetail</p></div>"

If(!(test-path $outputpath)){New-Item -ItemType Directory -Force -Path $outputpath }

import-module activedirectory

# Get inactive users as selected output type
$Title = "Inactive Users Report"
$Subtitle = "User accounts that have not logged on to Active Directory in 180 days or more."
Write-Host "Collecting $Title. This should only take a moment."
switch ($outputtype)
{
	"CSV" {
		search-adaccount -accountinactive -usersonly -timespan "195" | select-object -property name,distinguishedname,lastlogondate,enabled | export-csv $inactiveusers -notypeinformation
	}
	"HTML" {
		$HTMLPrefixed = $HTMLPre -replace "REPORTTITLE", "$Title" -replace "REPORTSUBTITLE", "$Subtitle"
		search-adaccount -accountinactive -usersonly -timespan "195" | select-object -property name,distinguishedname,lastlogondate,enabled | Where { $_.Enabled -eq $True} | Sort-Object -Property lastlogondate | ConvertTo-Html -Title "$title" -PreContent $HTMLPrefixed -post $HTMLPost | out-file -filepath $inactiveusers
	}
	default{
		Write-Error "No action defined for this output type."
	}
}

# Get inactive computers as selected output type
$Title = "Inactive Computers Report"
$Subtitle = "Computer accounts that have not logged on to Active Directory in 180 days or more."
Write-Host "Collecting $Title. This should only take a moment."
switch ($outputtype)
{
	"CSV" {
		search-adaccount -accountinactive -computersonly -timespan "195" | select-object -property name,distinguishedname,lastlogondate,enabled | export-csv $inactivepcs -notypeinformation
	}
	"HTML" {
		$HTMLPrefixed = $HTMLPre -replace "REPORTTITLE", "$Title" -replace "REPORTSUBTITLE", "$Subtitle"
		search-adaccount -accountinactive -computersonly -timespan "195" | select-object -property name,distinguishedname,lastlogondate,enabled | Where { $_.Enabled -eq $True} | Sort-Object -Property lastlogondate | ConvertTo-Html -Title "$title" -PreContent $HTMLPrefixed -post $HTMLPost | out-file -filepath $inactivepcs
	}
	default{
		Write-Error "No action defined for this output type."
	}
}

# Get domain admins as selected output type
$Title = "Domain Administrators Report"
$Subtitle = "Active accounts with Domain Administrator permissions"
Write-Host "Collecting $Title. This should only take a moment."
switch ($outputtype)
{
	"CSV" {
		Get-ADGroupMember -Identity 'Domain Admins' | Get-ADUser -Properties Name,Description,lastlogondate,enabled | select-object -property name,distinguishedname,lastlogondate,enabled | export-csv $domainadmins -notypeinformation
	}
	"HTML" {
		$HTMLPrefixed = $HTMLPre -replace "REPORTTITLE", "$Title" -replace "REPORTSUBTITLE", "$Subtitle"
		Get-ADGroupMember -Identity 'Domain Admins' | Get-ADUser -Properties Name,Description,lastlogondate,enabled | select-object -property name,distinguishedname,lastlogondate,enabled | Sort-Object -Property lastlogondate | ConvertTo-Html -Title "$title" -PreContent $HTMLPrefixed -post $HTMLPost | out-file -filepath $domainadmins
	}
	default{
		Write-Error "No action defined for this output type."
	}
}

# Get server disk space as selected output type
$Title = "Server Storage Report"
$Subtitle = "Hard drive space on servers"
Write-Host "Collecting $Title. This may take a while."
# but not if we're the local system account
if ($reportingby -ne "NT AUTHORITY\SYSTEM")
{
	switch ($outputtype)
	{
		"CSV" {
			$Servers = Get-ADComputer -Filter { OperatingSystem -Like '*Windows Server*' } -Properties OperatingSystem,enabled | Where { $_.Enabled -eq $True} | select -ExpandProperty Name; $output = foreach ($server in $servers){if (test-connection -computername $server -count 1 -quiet){Get-WmiObject Win32_LogicalDisk -ComputerName $server -Filter DriveType=3 | Select-Object @{'Name'='Server Name'; 'Expression'={$server}}, DeviceID, @{'Name'='Size (GB)'; 'Expression'={[math]::truncate($_.size / 1GB)}}, @{'Name'='Freespace (GB)'; 'Expression'={[math]::truncate($_.freespace / 1GB)}}}}; $output | export-csv $diskfreespace -notypeinformation
		}
		"HTML" {
			$HTMLPrefixed = $HTMLPre -replace "REPORTTITLE", "$Title" -replace "REPORTSUBTITLE", "$Subtitle"
			$Servers = Get-ADComputer -Filter { OperatingSystem -Like '*Windows Server*' } -Properties OperatingSystem,enabled | Where { $_.Enabled -eq $True} | select -ExpandProperty Name
			Get-WmiObject Win32_LogicalDisk -ComputerName $Servers -Filter "DriveType='3'" -ErrorAction SilentlyContinue | Select-Object PsComputerName, DeviceID, @{N="Disk Size (GB) ";e={[math]::Round($($_.Size) / 1073741824,2)}}, @{N="Free Space (GB)";e={[math]::Round($($_.FreeSpace) / 1073741824,2)}}, @{N="Used Space (%)";e={[math]::Round($($_.Size - $_.FreeSpace) / $_.Size * 100,1)}}, @{N="Used Space (GB)";e={[math]::Round($($_.Size - $_.FreeSpace) / 1073741824,2)}} | ConvertTo-Html -Title "$title" -PreContent $HTMLPrefixed -post $HTMLPost | out-file -filepath $diskfreespace
		}
		default{
			Write-Error "No action defined for this output type."
		}
	}
}else {
Write-Warning "Skipped collecting $Title. This report cannot run as $reportingby."
}

# Get service accounts report as selected output type
$Title = "Service Accounts Report"
$Subtitle = "Windows Services using a custom Log On As account. This report may be empty."
Write-Host "Collecting $Title. This may take a while."
# but not if we're the local system account
if ($reportingby -ne "NT AUTHORITY\SYSTEM")
{
	switch ($outputtype)
	{
		"CSV" {
# not tested yet
			$Servers = Get-ADComputer -Filter { OperatingSystem -Like '*Windows Server*' } -Properties OperatingSystem,enabled | Where { $_.Enabled -eq $True} | select -ExpandProperty Name; $output = foreach ($server in $servers){if (test-connection -computername $server -count 1 -quiet){Get-WmiObject Win32_Service -ComputerName $Servers -Filter "not StartMode='Disabled'" ErrorAction SilentlyContinue | Where -Property StartName -notlike "" | Where -Property StartName -notmatch "LocalSystem" | Where -Property StartName -notmatch "LocalService" | Where -Property StartName -notmatch "NetworkService" | Select-Object Name, StartName}; $output | export-csv $serviceaccounts -notypeinformation	}
		}
		"HTML" {
			$HTMLPrefixed = $HTMLPre -replace "REPORTTITLE", "$Title" -replace "REPORTSUBTITLE", "$Subtitle"
			$Servers = Get-ADComputer -Filter { OperatingSystem -Like '*Windows Server*' } -Properties OperatingSystem,enabled | Where { $_.Enabled -eq $True} | select -ExpandProperty Name
			Get-WmiObject Win32_Service -ComputerName $Servers -Filter "not StartMode='Disabled'" -ErrorAction SilentlyContinue | Select-Object PsComputerName, Name, StartName | Where -Property StartName -notlike "" | Where -Property StartName -notmatch "LocalSystem" | Where -Property StartName -notmatch "LocalService" | Where -Property StartName -notmatch "NetworkService" | ConvertTo-Html -Title "$title" -PreContent $HTMLPrefixed -post $HTMLPost | out-file -filepath $serviceaccts
		}
		default{
			Write-Error "No action defined for this output type."
		}
	}
}else {
Write-Warning "Skipped collecting $Title. This report cannot run as $reportingby."
}

# Get custom Active Directory Groups and their users as selected output type
$Title = "Active Directory Groups Report"
$Subtitle = "Groups specific to this organization and their members. Default Built-in groups are excluded."
Write-Host "Collecting $Title. This may take a while."
switch ($outputtype)
{
	"CSV" {
		$Groups = Get-ADGroup -Filter { GroupCategory -eq "Security" -and GroupScope -eq "Global"  } -Properties isCriticalSystemObject | Where-Object { !($_.IsCriticalSystemObject)}; $Results = foreach( $Group in $Groups ){Get-ADGroupMember -Identity $Group | foreach {[pscustomobject]@{GroupName = $Group.Name; Name = $_.Name}}}; $Results | export-csv $ADGroupReport -notypeinformation
	}
	"HTML" {
		$HTMLPrefixed = $HTMLPre -replace "REPORTTITLE", "$Title" -replace "REPORTSUBTITLE", "$Subtitle"
		$Groups = Get-ADGroup -Filter { GroupCategory -eq "Security" -and GroupScope -eq "Global"  } -Properties isCriticalSystemObject | Where-Object { !($_.IsCriticalSystemObject)}
		$Results = foreach( $Group in $Groups ){
		    Get-ADGroupMember -Identity $Group | foreach {
		        [pscustomobject]@{
		            GroupName = $Group.Name
		            Name = $_.Name
		            }
		        }
		    }
		$Results | ConvertTo-Html -Title "$title" -PreContent $HTMLPrefixed -post $HTMLPost | out-file -filepath $ADGroupReport
	}
	default{
	Write-Error "No action defined for this output type."
	}
}
Write-Host "Report generation finished."

If ($zipoutput = $true){
	#create scratch directory and move output files there
	Write-Host "Creating ZIP working directory."
	$scratchpath = $outputpath + "scratch\"
	If (!(Test-Path -LiteralPath $scratchpath)){New-Item -Path $scratchpath -ItemType Directory -ErrorAction Stop | Out-Null}
	Get-ChildItem -Path $outputpath $outputprefix*.$outputtype | Move-Item -Destination $scratchpath 

	Write-Host "Adding files to ZIP."
	#zip scratch to output using powershell v4 method
	$destinationZipFileName = $outputpath + "QBRData.zip"
	If (Test-Path -LiteralPath $destinationZipFileName){
		Write-Warning "ZIP file $destinationZipFileName already exists. Removing."
		Remove-Item -Path $destinationZipFileName -Force
		}
	[Reflection.Assembly]::LoadWithPartialName("System.IO.Compression.FileSystem") | Out-Null
	[System.IO.Compression.ZipFile]::CreateFromDirectory($scratchpath, $destinationZipFileName) | Out-Null

	If (Test-Path -LiteralPath $destinationZipFileName){
	Write-Host "ZIP file $destinationZipFileName creation finished."
	}

	#remove the scratch directories
	Write-Host "Removing ZIP working directory."
	Remove-Item -Path $scratchpath -recurse -force


}

Write-Host "Done."