#New-RMMReportMangler
#requires -module ImportExcel
Param (
    [Parameter(ValueFromPipelineByPropertyName)]
    [string] $ReportPath
)
$Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

#check for Fortinet EDR since it breaks everything useful
$FortinetEDRPresent = Get-Service -Name "FortiEDR Collector Service" -ErrorAction SilentlyContinue
if ($FortinetEDRPresent) { Write-Warning "Fortinet EDR service present on this machine. Network connections may be blocked." }

$HasNewPS = $PSVersionTable.PSVersion -ge [version]::new(7, 4)
$HasPowerHTML = Get-Module -ListAvailable -Name PowerHTML -ErrorAction SilentlyContinue

$reportExists = Test-Path $ReportPath
if (!($reportExists)) { Write-Warning "Specified report file $ReportFile does not exist or could not be read. Exiting."; exit } else {

    Import-Module ImportExcel

    # Open Excel File
    $excel = open-excelpackage $ReportPath

    # Remove cover sheet
    $CoverSheetName = "Report Header"
    if ($excel.Workbook.Worksheets[$CoverSheetName]) { 
        $excel.Workbook.Worksheets.Delete($CoverSheetName) 
        Write-Verbose "Removed cover sheet."
    }

    # Set Worksheet
    $sheet = $excel.Workbook.Worksheets["Network Hardware"]

    # Trim extra blank row and columns
    $cell = $sheet.Cells[1, 1]
    if ($cell.value -ne 'Customer Name') {
        $sheet.DeleteRow(1)
        $sheet.DeleteColumn(18)
        $sheet.DeleteColumn(17)
        $sheet.DeleteColumn(16)
        $sheet.DeleteColumn(1)
        Write-Verbose "Trimmed padding."
    }

    # Get customer name
    # not using this; customer names contain invalid characters
    #$CustomerName = $sheet.Cells[2, 1].Value
    #Write-Output $CustomerName
    #$ReportPath = "$CustomerName - $ReportPath"

    # Save File
    try {
        Close-ExcelPackage $excel -SaveAs $ReportPath
    }
    catch {
        Write-Warning "Error saving $ReportPath; check that the file is not already open in Excel. Exiting."; exit
    }

    
    #Get the index of the columns with the data we need to apply conditional formatting to
    $RecordsCount = $(Import-Excel $ReportPath -WorksheetName "Network Hardware").count
    [int]$HighlightRangeUpper = $RecordsCount + 1

    # Open Excel File 
    $excel = open-excelpackage $ReportPath

    # Set Worksheet
    $sheet = $excel.Workbook.Worksheets["Network Hardware"]

    # Check whether the sheet contains data in the correct place.
    $cell = $sheet.Cells[1, 1]
    if ($cell.value -ne 'Customer Name') { Write-Warning "Unexpected value in cell A1. Exiting."; exit } else {
        Write-Verbose "Report data aligned to A1."

        $columnCount = $sheet.Dimension.Columns
    
        $firstrow = foreach ($column in 1..$columnCount) {
            $sheet.Cells[1, $column].Value
        }
    
        #Get the column letters with the data we need to apply conditional formatting to
        $DeviceNameColumn = Get-ExcelColumnName $($firstrow.IndexOf('Device Name') + 1) | Select-Object -ExpandProperty ColumnName
        $DeviceModelColumn = Get-ExcelColumnName $($firstrow.IndexOf('Make / Model') + 1) | Select-Object -ExpandProperty ColumnName
        $CPUDescriptionColumn = Get-ExcelColumnName $($firstrow.IndexOf('CPU Description') + 1) | Select-Object -ExpandProperty ColumnName
        $RAMColumn = Get-ExcelColumnName $($firstrow.IndexOf('RAM (MB)') + 1) | Select-Object -ExpandProperty ColumnName
        $DiskColumn = Get-ExcelColumnName $($firstrow.IndexOf('Total Disk (GB)') + 1) | Select-Object -ExpandProperty ColumnName
        $OSColumn = Get-ExcelColumnName $($firstrow.IndexOf('OS and Service Pack') + 1) | Select-Object -ExpandProperty ColumnName
        $WarrantyColumn = Get-ExcelColumnName $($firstrow.IndexOf('Warranty Expiry') + 1) | Select-Object -ExpandProperty ColumnName

        #Highlight Win11 incompatible CPU
        # this one is clumsy, it requires a second worksheet containing only the incompatible CPUs
        $CPUConditionalFormatExpression = "=NOT(ISERROR(VLOOKUP(`$$($DeviceNameColumn)2, 'Win11 Incompatible'!`$$($DeviceNameColumn):`$$($DeviceNameColumn), 1, FALSE)))"
        Add-ConditionalFormatting -WorkSheet $sheet -Address "$($CPUDescriptionColumn)2:$($CPUDescriptionColumn)$HighlightRangeUpper" -RuleType Expression -ConditionValue $CPUConditionalFormatExpression -ForeGroundColor DarkRed -BackgroundColor LightPink

        #Highlight low memory
        #column name is "RAM (MB)"
        Add-ConditionalFormatting -WorkSheet $sheet -Address "$($RAMColumn)2:$($RAMColumn)$HighlightRangeUpper" -RuleType LessThan -ConditionValue "8192" -ForeGroundColor DarkRed -BackgroundColor LightPink

        #Highlight low disk
        #column name is "Total Disk (GB)"
        $DiskConditionalFormattingExpressionFail = "=$($DiskColumn)2<=128"
        $DiskConditionalFormattingExpressionWarn = "=(AND($($DiskColumn)2>128,$($DiskColumn)2<=256))"
        Add-ConditionalFormatting -WorkSheet $sheet -Address "$($DiskColumn)2:$($DiskColumn)$HighlightRangeUpper" -RuleType Expression -ConditionValue $DiskConditionalFormattingExpressionFail -ForeGroundColor DarkRed -BackgroundColor LightPink
        Add-ConditionalFormatting -WorkSheet $sheet -Address "$($DiskColumn)2:$($DiskColumn)$HighlightRangeUpper" -RuleType Expression -ConditionValue $DiskConditionalFormattingExpressionWarn -ForeGroundColor DarkYellow -BackgroundColor LightYellow

        #highlight end-of-life OS
        #column name is "OS and Service Pack" 
        $OSConditionalFormattingExpression = "=NOT(OR(ISNUMBER(SEARCH(`"10 Pro`", $($OSColumn)2)), ISNUMBER(SEARCH(`"10 Enterprise`", $($OSColumn)2)), ISNUMBER(SEARCH(`"10 Business`", $($OSColumn)2)), ISNUMBER(SEARCH(`"11 Pro`", $($OSColumn)2)), ISNUMBER(SEARCH(`"11 Enterprise`", $($OSColumn)2)), ISNUMBER(SEARCH(`"11 Business`", $($OSColumn)2)), ISNUMBER(SEARCH(`"Server 2016`", $($OSColumn)2)), ISNUMBER(SEARCH(`"Server 2019`", $($OSColumn)2)), ISNUMBER(SEARCH(`"Server 2022`", $($OSColumn)2))))"
        Add-ConditionalFormatting -WorkSheet $sheet -Address "$($OSColumn)2:$($OSColumn)$HighlightRangeUpper" -RuleType Expression -ConditionValue $OSConditionalFormattingExpression -ForeGroundColor DarkRed -BackgroundColor LightPink
        
        #highlight warranty expiry (exclude virtual devices)
        #column name is "Warranty Expiry" 
        $WarrantyConditionalFormattingExpressionFail = "=AND(TODAY()-$($WarrantyColumn)2>365, ISERROR(SEARCH(`"irtual`", $($DeviceModelColumn)2)), ISERROR(SEARCH(`"VMware`", $($DeviceModelColumn)2)))"
        $WarrantyConditionalFormattingExpressionWarn = "=AND(TODAY()-$($WarrantyColumn)2>0,TODAY()-$($WarrantyColumn)2<=365)"
        Add-ConditionalFormatting -WorkSheet $sheet -Address "$($WarrantyColumn)2:$($WarrantyColumn)$HighlightRangeUpper" -RuleType Expression -ConditionValue $WarrantyConditionalFormattingExpressionFail -ForeGroundColor DarkRed -BackgroundColor LightPink
        Add-ConditionalFormatting -WorkSheet $sheet -Address "$($WarrantyColumn)2:$($WarrantyColumn)$HighlightRangeUpper" -RuleType Expression -ConditionValue $WarrantyConditionalFormattingExpressionWarn -ForeGroundColor DarkYellow -BackgroundColor LightYellow      

        # Save File
        Close-ExcelPackage $excel -SaveAs $ReportPath
    }

    $RMMData = Import-Excel $ReportPath -WorksheetName "Network Hardware"
    
    #count systems with low memory
    [array]$LowMemory = @()
    #$LowMemory = $RMMData | Where-Object { $_.'Device Class' -notlike "*Server*" -and $_.'RAM (MB)' -lt "8192" }
    $LowMemory = $RMMData | Where-Object { $_.'RAM (MB)' -lt "8192" }
    Write-Output "Counted $($LowMemory.count) systems with less than 8GB RAM."

    #count systems with low disk
    [array]$LowDisk = @()
    $LowDisk = $RMMData | Where-Object { $_.'Total Disk (GB)' -gt "128" -and $_.'Total Disk (GB)' -le "256" }
    Write-Output "Counted $($LowDisk.count) systems with less than 256GB storage."
    $LowDisk = $RMMData | Where-Object { $_.'Total Disk (GB)' -le "128" }
    Write-Output "Counted $($LowDisk.count) systems with less than 128GB storage."

    $Endpoints = $RMMData | Where-Object { $_.'Device Class' -notlike "*Server*" }
    
    #count endpoints within warranty
    [array]$WarrantyOK = @()
    $WarrantyOK = $RMMData | Where-Object { $_.'Device Class' -notlike "*Server*" -and $null -ne $_.'Warranty Expiry' -and $(New-TimeSpan -Start $_.'Warranty Expiry').Days -lt 0 }
    Write-Output "Counted $($WarrantyOK.count) endpoints under warranty."
    $Percent = [math]::Round($($($WarrantyOK.count) / $($Endpoints.count) * 100), 2)
    Write-Output "$Percent % of total endpoints."

    #count endpoints with expired warranty within the last year
    [array]$WarrantyExpired = @()
    $WarrantyExpired = $RMMData | Where-Object { $_.'Device Class' -notlike "*Server*" -and $null -ne $_.'Warranty Expiry' -and $(New-TimeSpan -Start $_.'Warranty Expiry').Days -ge 0 -and $(New-TimeSpan -Start $_.'Warranty Expiry').Days -lt 365 }
    Write-Output "Counted $($WarrantyExpired.count) endpoints with expired warranty less than a year ago."
    $Percent = [math]::Round($($($WarrantyExpired.count) / $($Endpoints.count) * 100), 2)
    Write-Output "$Percent % of total endpoints."
    
    #count endpoints with expired warranty over 1y
    [array]$WarrantyExpired = @()
    $WarrantyExpired = $RMMData | Where-Object { $_.'Device Class' -notlike "*Server*" -and $null -ne $_.'Warranty Expiry' -and $(New-TimeSpan -Start $_.'Warranty Expiry').Days -ge 365 }
    Write-Output "Counted $($WarrantyExpired.count) endpoints with expired warranty more than a year ago."
    $Percent = [math]::Round($($($WarrantyExpired.count) / $($Endpoints.count) * 100), 2)
    Write-Output "$Percent % of total endpoints."
   
    #count endpoints with no warranty data
    [array]$WarrantyNoData = @()
    $WarrantyNoData = $RMMData | Where-Object { $_.'Device Class' -notlike "*Server*" -and $null -eq $_.'Warranty Expiry' }
    Write-Output "Counted $($WarrantyNoData.count) endpoints with no warranty data."
    $Percent = [math]::Round($($($WarrantyNoData.count) / $($Endpoints.count) * 100), 2)
    Write-Output "$Percent % of total endpoints."

    #count servers with EOL OS
    [array]$EOLServers = @()
    $EOLServers = $RMMData | Where-Object { $_.'Device Class' -like "*Server*" -and $_.'OS and Service Pack' -notlike "*2016*" -and $_.'OS and Service Pack' -notlike "*2019*" -and $_.'OS and Service Pack' -notlike "*2022*" }
    Write-Output "Counted $($EOLServers.count) servers with end-of-support operating systems."

    if (!($HasPowerHTML -and $HasNewPS)) { Write-Warning "Skipping Win11 CPU compatibility count because prereqs not met." } else {
        #count CPUs not supported by win11 (needs PS 7)
        $AMDURL = "https://learn.microsoft.com/en-us/windows-hardware/design/minimum/supported/windows-11-supported-amd-processors"
        $IntelURL = "https://learn.microsoft.com/en-us/windows-hardware/design/minimum/supported/windows-11-supported-intel-processors"
        $SupportedAMDCPUs = Get-HtmlTable $AMDURL
        $SupportedIntelCPUs = Get-HtmlTable $IntelURL

        $SupportedAMDCPUModelStrings = @()
        Foreach ($Model in $SupportedAMDCPUs) {
            $modelString = $($Model.Model | Out-String ).Trim()
            $modelObject = [PSCustomObject]@{ 'Model' = $modelString }
            $SupportedAMDCPUModelStrings += $modelObject
        }    

        $SupportedIntelCPUModelStrings = @()
        Foreach ($Model in $SupportedIntelCPUs) {
            $modelString = $($Model.Model | Out-String ).Trim()
            $modelObject = [PSCustomObject]@{ 'Model' = $modelString }
            $SupportedIntelCPUModelStrings += $modelObject
        }    

        Function CheckWin11CPUSupport($CPUString) {
            Switch ($CPUString) {
                { $_ -like "Intel*" } { if (($($SupportedIntelCPUModelStrings.Model) | ForEach-Object { $CPUstring.contains($_) }) -notcontains $true) { return $true } else { return $false } }
                { $_ -like "AMD*" } { if (($($SupportedAMDCPUModelStrings.Model) | ForEach-Object { $CPUstring.contains($_) }) -notcontains $true) { return $true } else { return $false } }
                default { return $false }
            } 
        }
    }
    $DevicesWithUnSupportedCPUs = $RMMData | Where-Object { CheckWin11CPUSupport($_.'CPU Description') }
    Write-Output "Counted $($DevicesWithUnSupportedCPUs.count) devices with CPUs that do not support Windows 11."
    $DevicesWithUnSupportedCPUs | Export-Excel -Path $Reportpath -ClearSheet -BoldTopRow -Autosize -FreezePane 2 -Autofilter -WorkSheetname "Win11 Incompatible" 

}

Write-Output "Finished in $($Stopwatch.Elapsed.TotalSeconds) seconds."
Write-Output "Report output path is $ReportPath."
$Stopwatch.Stop()
