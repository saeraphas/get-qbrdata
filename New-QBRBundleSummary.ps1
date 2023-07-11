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

$reportExists = Test-Path $ReportPath
if (!($reportExists)) { Write-Warning "Specified report file $ReportFile does not exist or could not be read. Exiting."; exit } else {

    $usersdata = Import-Excel $ReportPath -WorkSheetName "User Account Audit Report"
    
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

    $InactiveServersData = Import-Excel $ReportPath -WorkSheetName "Offline Servers Report"

    #count offline server accounts
    [array]$InactiveServers = @()
    $InactiveServers = $InactiveServersData
    Write-Output "Counted $($InactiveServers.count) servers not reachable by PING or SMB."

    $InactiveComputersData = Import-Excel $ReportPath -WorkSheetName "Inactive Computers Report"

    #count computer accounts with 180d+ inactive
    [array]$InactiveComputers = @()
    $InactiveComputers = $InactiveComputersData
    Write-Output "Counted $($InactiveComputers.count) computer accounts not signed in for 180d+."

    $DomainAdminsData = Import-Excel $ReportPath -WorkSheetName "Domain Administrators Report"

    #count domain admins
    [array]$DomainAdmins = @()
    $DomainAdmins = $DomainAdminsData
    Write-Output "Counted $($DomainAdmins.count) Domain Administrators."

    #count domain admins with blank description
    [array]$DomainAdmins = @()
    $DomainAdminsWithoutDescription = $DomainAdminsData | Where-Object { $null -eq $_.'description' }
    Write-Output "Counted $($DomainAdminsWithoutDescription.count) Domain Administrators with no description set."

    $DiskUtilizationData = Import-Excel $ReportPath -WorkSheetName "Server Storage Report"

    #count volumes with less than 10GB free space or greater than 90% used space
    [array]$LowDiskSpace = @()
    $LowDiskSpace = $DiskUtilizationData | Where-Object { $_.'Free Space (GB)' -le 10 -or $_.'Used Space (%)' -gt 90 }
    Write-Output "Counted $($LowDiskSpace.count) volumes with low disk space."

    $NameServersData = Import-Excel $ReportPath -WorkSheetName "Static DNS Servers"

    #count unique DNS configs
    [array]$UniqueNameservers = @()
    $UniqueNameservers = $NameServersData | Select-Object -Property NameServers -Unique
    Write-Output "Counted $($UniqueNameservers.count) unique nameserver configurations."

    $SSLCertificateData = Import-Excel $ReportPath -WorkSheetName "SSL Certificates"

    #count SSL certificates expiring within 90d.
    [array]$ExpiringCertificates = @()
    $ExpiringCertificates = $SSLCertificateData | Where-Object { $_.'ExpiryDays' -le 90 }
    Write-Output "Counted $($ExpiringCertificates.count) SSL certificates expiring within 90 days."

    $OOSWorkstationsData = Import-Excel $ReportPath -WorkSheetName "End-of-Support PCs Report"

    #count end-of-support servers
    [array]$OOSServers = @()
    $OOSServers = $OOSServersData | Where-Object { -not $_.'Result' -eq "This report is empty." }
    Write-Output "Counted $($OOSServers.count) servers with end-of-support OS."

    #count end-of-support workstation
    [array]$OOSWorkstations = @()
    $OOSWorkstations = $OOSWorkstationsData | Where-Object { -not $_.'Result' -eq "This report is empty." }
    Write-Output "Counted $($OOSWorkstations.count) workstations with end-of-support OS."

    $OOSServersData = Import-Excel $ReportPath -WorkSheetName "End-of-Support Servers Report"

    Write-Output "Finished."
}