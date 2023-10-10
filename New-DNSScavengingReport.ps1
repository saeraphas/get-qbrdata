<#
.SYNOPSIS
	This script collects DNS scavenging data and provides a method to manually scavenge old records.

.DESCRIPTION
	Nascent stale DNS report and manual scavenger. Still developing.
    Requires the DnsServer module.

    Supposedly this will install it on Windows Server, but it hasn't worked yet for me. 
    Add-WindowsCapability -Online -Name "Rsat.Dns.Tools~~~~0.0.1.0"

.EXAMPLE
	.\New-DNSScavengingReport.ps1

.NOTES
    Author:             Douglas Hammond (douglas@douglashammond.com)
	License: 			This script is distributed under "THE BEER-WARE LICENSE" (Revision 42):
						As long as you retain this notice you can do whatever you want with this stuff.
						If we meet some day, and you think this stuff is worth it, you can buy me a beer in return.
#>
#requires -Modules DnsServer

Param (
    [Parameter(ValueFromPipelineByPropertyName)]
    [int] $StaleRecordThresholdDays,
    [switch] $Scavenge
)

$Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

function CheckPrerequisites($PrerequisiteModulesTable) {
    $PrerequisiteModules = $PrerequisiteModulesTable | ConvertFrom-Csv
    $ProgressActivity = "Checking for prerequisite modules."
    ForEach ( $PrerequisiteModule in $PrerequisiteModules ) {
        $moduleName = $($PrerequisiteModule.Name)
        $ProgressOperation = "Checking for module $moduleName."
        Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation
        $minimumVersion = $($PrerequisiteModule.minimumversion)
        $installedversion = $(Get-Module -ListAvailable -Name $moduleName | Select-Object -first 1).version
        If (!($installedversion)) {
            try { Install-Module $moduleName -Repository PSGallery -AllowClobber -scope CurrentUser -Force -RequiredVersion $minimumversion } catch { Write-Warning "An error occurred installing $moduleName."; exit }
        }
        elseif ([version]$installedversion -lt [version]$minimumversion) {
            Write-Warning "The installed version of $moduleName is lower than the required version $minimumversion."
            #try { Install-Module $moduleName -Repository PSGallery -AllowClobber -scope CurrentUser -Force -RequiredVersion $minimumversion } catch { Write-Error "An error occurred installing $moduleName."; exit }
            #try { Update-Module -Name $moduleName -Repository PSGallery -Force -RequiredVersion $minimumversion } catch { Write-Error "An error occurred updating $moduleName."; exit }
            #try { Uninstall-Module $moduleName -AllVersions } catch { Write-Error "An error occurred removing $moduleName. You may need to manually remove old versions using admin privileges."; exit }
        }
    }
    Write-Progress -Activity $ProgressActivity -Completed
}

$PrerequisiteModulesTable = @'
Name,MinimumVersion
ImportExcel,7.0.0
'@
CheckPrerequisites($PrerequisiteModulesTable)

#define the output path
$DateString = ((get-date).tostring("yyyy-MM-dd"))
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$ReportPath = "$DesktopPath\Reports"
#get NETBIOS domain name
try { $ADDomain = (Get-WMIObject Win32_NTDomain).DomainName } catch { Write-Warning "An error occurred getting the AD domain name." }
if ($ADDomain.length -ge 1) { $Customer = $ADDomain } else { $Customer = "Nexigen" }
$ReportType = "DNS-Scavenging"
$XLSreport = "$ReportPath\$Customer\$Customer-$ReportType-$DateString.xlsx"

$PSNewLine = [System.Environment]::Newline

#$Scavenge = $false #set $true to remove stale records
if (!($StaleRecordThresholdDays)) { $StaleRecordThresholdDays = -28 } #must be a negative number

$ZoneNames = try { Get-DnsServerZone } catch { Write-Warning "Error getting DNS zones."; exit }

#list of valid record types, because I can't figure out how to pull out a validate set. Server 2012 barfs on TLSA.
#https://learn.microsoft.com/en-us/powershell/module/dnsserver/get-dnsserverresourcerecord?view=windowsserver2022-ps#-rrtype
$RecordTypes = @('HInfo', 'Afsdb', 'Atma', 'Isdn', 'Key', 'Mb', 'Md', 'Mf', 'Mg', 'MInfo', 'Mr', 'Mx', 'NsNxt', 'Rp', 'Rt', 'Wks', 'X25', 'A', 'AAAA', 'CName', 'Ptr', 'Srv', 'Txt', 'Wins', 'WinsR', 'Ns', 'Soa', 'NasP', 'NasPtr', 'DName', 'Gpos', 'Loc', 'DhcId', 'Naptr', 'RRSig', 'DnsKey', 'DS', 'NSec', 'NSec3', 'NSec3Param', 'Tlsa')
$StaleRecordsReport = @()

foreach ($ZoneName in $ZoneNames) {
    #Get-DnsServerZoneAging -Name $($ZoneName.ZoneName) #not used yet. 
    foreach ($RecordType in $RecordTypes) {
        $WarningMessage = "An error occurred looking up $RecordType records for $ZoneName." + $PSNewLine + "This RRType may not be supported in this version of Windows Server."
        $StaleRecords = try { Get-DnsServerResourceRecord -ZoneName $($ZoneName.ZoneName) -RRtype $RecordType | Where-Object { $_.Timestamp } | Where-Object { $_.Timestamp -lt ((Get-Date).AddDays($StaleRecordThresholdDays)) } } catch { Write-Warning $WarningMessage; $StaleRecords = $false }
        Write-Verbose "Zone $($ZoneName.ZoneName) has $($StalePTRRecords.count) stale $RecordType records."
        if ($($StaleRecords).count -ge 1) {
            if ($Scavenge) { $StaleRecords | remove-DnsServerResourceRecord -ZoneName $ZoneName.ZoneName -force }
            $StaleRecordsReport += $StaleRecords
        }
    }
}

Write-Host "Total Stale Records: $($StaleRecordsReport.count)."
#$StaleRecordsReport | Export-CSV -NTI -path c:\nexigen\stale-dns-all.csv
$StaleRecordsReport | Select-Object -Property RecordType, Hostname, Distinguishedname, Timestamp | Export-Excel `
    -Path $XLSreport `
    -WorkSheetname "$ReportType" `
    -ClearSheet `
    -BoldTopRow `
    -Autosize `
    -FreezePane 2 `
    -Autofilter 

Write-Output "Finished in $($Stopwatch.Elapsed.TotalSeconds) seconds."
Write-Output "Report output path is $XLSreport."
$Stopwatch.Stop()