<#
.SYNOPSIS
	Summarizes QBR Bundles. 

.DESCRIPTION
	Imports the specified quick report XLSX and outputs a summary list of common review points.

.EXAMPLE
	.\New-QBRBundleSummary.ps1

.NOTES
    Author:             Douglas Hammond (douglas@douglashammond.com)
	License: 			This script is distributed under "THE BEER-WARE LICENSE" (Revision 42):
						As long as you retain this notice you can do whatever you want with this stuff.
						If we meet some day, and you think this stuff is worth it, you can buy me a beer in return.
#>

Param (
    [Parameter(ValueFromPipelineByPropertyName)]
    [string] $ReportPath
)

$Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

$reportExists = Test-Path $ReportPath
if (!($reportExists)) { Write-Warning "Specified report file $ReportFile does not exist or could not be read. Exiting."; exit } else {

    $usersdata = Import-Excel $ReportPath -WorkSheetName "Users - Audit" | Where-Object { $_.Result -ne "This report is empty." }
    
    #count user accounts with 30d+ inactive
    [array]$UsersInactive30Days = @() #strong typing in case there's exactly 1 result
    $UsersInactive30Days = $usersdata | Where-Object { $_.'Enabled' -eq "TRUE" -and $_.'Days Since Last Logon' -ge 30 }
    Write-Output "Counted $($UsersInactive30Days.count) user accounts not signed in for 30d+."

    #count user accounts with 180d+ inactive
    [array]$UsersInactive180Days = @()
    $UsersInactive180Days = $usersdata | Where-Object { $_.'Enabled' -eq "TRUE" -and $_.'Days Since Last Logon' -ge 180 }
    Write-Output "Counted $($UsersInactive180Days.count) user accounts not signed in for 180d+."

    #count users with no logon date
    [array]$UsersInactiveUnknownDays = @()
    $UsersInactiveUnknownDays = $usersdata | Where-Object { $_.'Enabled' -eq "TRUE" -and $null -eq $_.'Days Since Last Logon' }
    Write-Output "Counted $($UsersInactiveUnknownDays.count) user accounts not signed in for an unknown number of days."

    $InactiveServersData = Import-Excel $ReportPath -WorkSheetName "Servers - Offline" | Where-Object { $_.Result -ne "This report is empty." }

    #count offline server accounts
    [array]$InactiveServers = @()
    $InactiveServers = $InactiveServersData
    Write-Output "Counted $($InactiveServers.count) servers not reachable by PING or SMB."

    $InactiveComputersData = Import-Excel $ReportPath -WorkSheetName "Endpoints - Inactive" | Where-Object { $_.Result -ne "This report is empty." }

    #count computer accounts with 180d+ inactive
    [array]$InactiveComputers = @()
    $InactiveComputers = $InactiveComputersData
    Write-Output "Counted $($InactiveComputers.count) computer accounts not signed in for 180d+."

    $DomainAdminsData = Import-Excel $ReportPath -WorkSheetName "Domain Administrators" | Where-Object { $_.Result -ne "This report is empty." }

    #count domain admins
    [array]$DomainAdmins = @()
    $DomainAdmins = $DomainAdminsData
    Write-Output "Counted $($DomainAdmins.count) Domain Administrators."

    #count domain admins with blank description
    [array]$DomainAdmins = @()
    $DomainAdminsWithoutDescription = $DomainAdminsData | Where-Object { $null -eq $_.'description' }
    Write-Output "Counted $($DomainAdminsWithoutDescription.count) Domain Administrators with no description set."

    $DiskUtilizationData = Import-Excel $ReportPath -WorkSheetName "Servers - Storage Utilization" | Where-Object { $_.Result -ne "This report is empty." }

    #count volumes with less than 10GB free space or greater than 90% used space
    [array]$LowDiskSpace = @()
    $LowDiskSpace = $DiskUtilizationData | Where-Object { $_.'Free Space (GB)' -le 10 -or $_.'Used Space (%)' -gt 90 }
    Write-Output "Counted $($LowDiskSpace.count) volumes with low disk space."

    $NameServersData = Import-Excel $ReportPath -WorkSheetName "Servers - Interface DNS" | Where-Object { $_.Result -ne "This report is empty." }

    #count unique DNS configs
    [array]$UniqueNameservers = @()
    $UniqueNameservers = $NameServersData | Select-Object -Property NameServers -Unique
    Write-Output "Counted $($UniqueNameservers.count) unique nameserver configurations."

    $SSLCertificateData = Import-Excel $ReportPath -WorkSheetName "Servers - SSL Certificates" | Where-Object { $_.Result -ne "This report is empty." }

    #count SSL certificates expiring within 90d.
    [array]$ExpiringCertificates = @()
    $ExpiringCertificates = $SSLCertificateData | Where-Object { $_.'ExpiryDays' -ge 0 -and $_.'ExpiryDays' -lt 90 }
    Write-Output "Counted $($ExpiringCertificates.count) SSL certificates expiring within 90 days."

    $OOSWorkstationsData = Import-Excel $ReportPath -WorkSheetName "Endpoints - End-of-Support" | Where-Object { $_.Result -ne "This report is empty." }

    #count end-of-support workstation
    [array]$OOSWorkstations = @()
    $OOSWorkstations = $OOSWorkstationsData | Where-Object { -not $_.'Result' -eq "This report is empty." }
    Write-Output "Counted $($OOSWorkstations.count) workstations with end-of-support OS."

    $OOSServersData = Import-Excel $ReportPath -WorkSheetName "Servers - End-of-Support" | Where-Object { $_.Result -ne "This report is empty." }

    #count end-of-support servers
    [array]$OOSServers = @()
    $OOSServers = $OOSServersData | Where-Object { -not $_.'Result' -eq "This report is empty." }
    Write-Output "Counted $($OOSServers.count) servers with end-of-support OS."

    $BitLockersData = Import-Excel $ReportPath -WorkSheetName "Endpoints - BitLocker Recovery" | Where-Object { -not $_.Result -eq "This report is empty." }

    #count workstations without BitLocker
    [array]$NoBitLocker = @()
    $NoBitLocker = $BitLockersData | Where-Object { $_.'Key Exists In AD' -eq "false" }
    Write-Output "Counted $($NoBitLocker.count) workstations with no BitLocker key in AD."

}

Write-Output "Finished in $($Stopwatch.Elapsed.TotalSeconds) seconds."
$Stopwatch.Stop()