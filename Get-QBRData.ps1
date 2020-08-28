#DQDQBRDC - Douglas' Quick and Dirty QBR Data Collector
#Requires -Module activedirectory

$owner 			= "Nexigen Communications, LLC"
$date 			= Get-Date -DisplayHint Date | Out-String
$outputpath 	= "C:\Nexigen\"
$outputtype 	= "HTML" #CSV or HTML
$inactiveusers 	= $outputpath + "qbr-inactiveusers.$outputtype"
$inactivepcs 	= $outputpath + "qbr-inactivepcs.$outputtype"
$domainadmins 	= $outputpath + "qbr-domainadmins.$outputtype"
$diskfreespace 	= $outputpath + "qbr-diskfreespace.$outputtype"
$ADGroupReport	= $outputpath + "qbr-adgroups.$outputtype"
$reportingby 	= $env:UserName
$reportingfrom 	= ([System.Net.Dns]::GetHostByName(($env:computerName))).Hostname
$zipoutput 		= $true
$headerdetail 	= "Report data generated by $reportingby on $date from $reportingfrom."
$footerdetail 	= "For questions or additional information, please contact $owner."
#not sure why this nested string expansion doesn't work right
#$HTMLPre 		= "<p><font size=`"6`">$Title</font><br>$Subtitle</p><P><font size=`"2`">$headerdetail</font></P><hr>"
$HTMLPost 		= "<hr><font size=`"2`">$footerdetail</font>"

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
		search-adaccount -accountinactive -usersonly -timespan "195" | select-object -property name,distinguishedname,lastlogondate,enabled | Where { $_.Enabled -eq $True} | Sort-Object -Property lastlogondate | ConvertTo-Html -Title "$title" -PreContent "<p><font size=`"6`">$Title</font><br>$Subtitle</p><P><font size=`"2`">$headerdetail</font></P><hr>" -post $HTMLPost | out-file -filepath $inactiveusers
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
		search-adaccount -accountinactive -computersonly -timespan "195" | select-object -property name,distinguishedname,lastlogondate,enabled | Where { $_.Enabled -eq $True} | Sort-Object -Property lastlogondate | ConvertTo-Html -Title "$title" -PreContent "<p><font size=`"6`">$Title</font><br>$Subtitle</p><P><font size=`"2`">$headerdetail</font></P><hr>" -post $HTMLPost | out-file -filepath $inactivepcs
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
		Get-ADGroupMember -Identity 'Domain Admins' | Get-ADUser -Properties Name,Description,lastlogondate,enabled | select-object -property name,distinguishedname,lastlogondate,enabled | Sort-Object -Property lastlogondate | ConvertTo-Html -Title "$title" -PreContent "<p><font size=`"6`">$Title</font><br>$Subtitle</p><P><font size=`"2`">$headerdetail</font></P><hr>" -post $HTMLPost | out-file -filepath $domainadmins
	}
	default{
		Write-Error "No action defined for this output type."
	}
}

# Get server disk space as selected output type
$Title = "Server Storage Report"
$Subtitle = "Hard drive space on servers"
Write-Host "Collecting $Title. This may take a while."
switch ($outputtype)
{
	"CSV" {
		$Servers = Get-ADComputer -Filter { OperatingSystem -Like '*Windows Server*' } -Properties OperatingSystem,enabled | Where { $_.Enabled -eq $True} | select -ExpandProperty Name; $output = foreach ($server in $servers){if (test-connection -computername $server -count 1 -quiet){Get-WmiObject Win32_LogicalDisk -ComputerName $server -Filter DriveType=3 | Select-Object @{'Name'='Server Name'; 'Expression'={$server}}, DeviceID, @{'Name'='Size (GB)'; 'Expression'={[math]::truncate($_.size / 1GB)}}, @{'Name'='Freespace (GB)'; 'Expression'={[math]::truncate($_.freespace / 1GB)}}}}; $output | export-csv $diskfreespace -notypeinformation
	}
	"HTML" {
		$Servers = Get-ADComputer -Filter { OperatingSystem -Like '*Windows Server*' } -Properties OperatingSystem,enabled | Where { $_.Enabled -eq $True} | select -ExpandProperty Name
		Get-WmiObject Win32_LogicalDisk -ComputerName $Servers -Filter "DriveType='3'" -ErrorAction SilentlyContinue | Select-Object PsComputerName, DeviceID, @{N="Disk Size (GB) ";e={[math]::Round($($_.Size) / 1073741824,2)}}, @{N="Free Space (GB)";e={[math]::Round($($_.FreeSpace) / 1073741824,2)}}, @{N="Used Space (%)";e={[math]::Round($($_.Size - $_.FreeSpace) / $_.Size * 100,1)}}, @{N="Used Space (GB)";e={[math]::Round($($_.Size - $_.FreeSpace) / 1073741824,2)}} | ConvertTo-Html -Title "$title" -PreContent "<p><font size=`"6`">$Title</font><br>$Subtitle</p><P><font size=`"2`">$headerdetail</font></P><hr>" -post $HTMLPost | out-file -filepath $diskfreespace
	}
	default{
		Write-Error "No action defined for this output type."
	}
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
		$Groups = Get-ADGroup -Filter { GroupCategory -eq "Security" -and GroupScope -eq "Global"  } -Properties isCriticalSystemObject | Where-Object { !($_.IsCriticalSystemObject)}
		$Results = foreach( $Group in $Groups ){
		    Get-ADGroupMember -Identity $Group | foreach {
		        [pscustomobject]@{
		            GroupName = $Group.Name
		            Name = $_.Name
		            }
		        }
		    }
		$Results | ConvertTo-Html -Title "$title" -PreContent "<p><font size=`"6`">$Title</font><br>$Subtitle</p><P><font size=`"2`">$headerdetail</font></P><hr>" -post $HTMLPost | out-file -filepath $ADGroupReport
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
	Get-ChildItem -Path $outputpath qbr-*.$outputtype | Move-Item -Destination $scratchpath 

	Write-Host "Adding files to ZIP."
	#zip scratch to output using powershell v4 method
	$destinationZipFileName = $outputpath + "QBRData.zip"
	[Reflection.Assembly]::LoadWithPartialName("System.IO.Compression.FileSystem") | Out-Null
	[System.IO.Compression.ZipFile]::CreateFromDirectory($scratchpath, $destinationZipFileName) | Out-Null

	#remove the scratch directories
	Write-Host "Removing ZIP working directory."
	Remove-Item -Path $scratchpath -recurse -force

	Write-Host "ZIP file $destinationZipFileName creation finished."
}

Write-Host "Done."