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
    $ws = $excel.Workbook.Worksheets["Network Hardware"]

    # Trim extra blank row and columns
    $cell = $ws.Cells[1, 1]
    if ($cell.value -ne 'Customer Name') {
        $ws.DeleteRow(1)
        $ws.DeleteColumn(18)
        $ws.DeleteColumn(17)
        $ws.DeleteColumn(16)
        $ws.DeleteColumn(1)
        Write-Verbose "Trimmed padding."
    }

    # Save File
    Close-ExcelPackage $excel -SaveAs $ReportPath

    $RecordsCount = $(Import-Excel $ReportPath -WorksheetName "Network Hardware").count
    [int]$HighlightRangeUpper = $RecordsCount + 1

    # Open Excel File
    $excel = open-excelpackage $ReportPath

    # Set Worksheet
    $ws = $excel.Workbook.Worksheets["Network Hardware"]

    # Check whether the sheet contains data in the correct place.
    $cell = $ws.Cells[1, 1]
    if ($cell.value -ne 'Customer Name') { Write-Warning "Unexpected value in cell A1. Exiting."; exit } else {
        Write-Verbose "Report data aligned to A1."

        #Highlight low memory
        Add-ConditionalFormatting -WorkSheet $ws -Address "I2:I$HighlightRangeUpper" -RuleType LessThan -ConditionValue "8192" -ForeGroundColor DarkRed -BackgroundColor LightPink

        #Highlight low disk
        Add-ConditionalFormatting -WorkSheet $ws -Address "J2:J$HighlightRangeUpper" -RuleType Expression -ConditionValue '=J2<=128' -ForeGroundColor DarkRed -BackgroundColor LightPink
        Add-ConditionalFormatting -WorkSheet $ws -Address "J2:J$HighlightRangeUpper" -RuleType Expression -ConditionValue '=(AND(J2>128,J2<=256))' -ForeGroundColor DarkYellow -BackgroundColor LightYellow

        #highlight end-of-life OS
        Add-ConditionalFormatting -WorkSheet $ws -Address "K2:K$HighlightRangeUpper" -RuleType Expression -ConditionValue '=NOT(OR(ISNUMBER(SEARCH("10 Pro", K2)), ISNUMBER(SEARCH("10 Enterprise", K2)), ISNUMBER(SEARCH("10 Business", K2)), ISNUMBER(SEARCH("11 Pro", K2)), ISNUMBER(SEARCH("11 Enterprise", K2)), ISNUMBER(SEARCH("11 Business", K2)), ISNUMBER(SEARCH("Server 2016", K2)), ISNUMBER(SEARCH("Server 2019", K2)), ISNUMBER(SEARCH("Server 2022", K2))))' -ForeGroundColor DarkRed -BackgroundColor LightPink
        
        #highlight warranty expiry 
        Add-ConditionalFormatting -WorkSheet $ws -Address "N2:N$HighlightRangeUpper" -RuleType Expression -ConditionValue '=TODAY()-N2>365' -ForeGroundColor DarkRed -BackgroundColor LightPink
        Add-ConditionalFormatting -WorkSheet $ws -Address "N2:N$HighlightRangeUpper" -RuleType Expression -ConditionValue '=AND(TODAY()-N2>30,TODAY()-N2<=365)' -ForeGroundColor DarkYellow -BackgroundColor LightYellow      

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
        $IntelURL = "https://learn.microsoft.com/en-us/windows-hardware/design/minimum/supported/windows-11-supported-intel-processors"
        $AMDURL = "https://learn.microsoft.com/en-us/windows-hardware/design/minimum/supported/windows-11-supported-amd-processors"
        $SupportedCPUs = @()
        Foreach ($URL in @($IntelURL, $AMDURL)) {
            $SupportedCPUs += Get-HtmlTable $URL
        }
        $SupportedCPUModelStrings = @()
        Foreach ($Model in $SupportedCPUs) {
            $modelString = $($Model.Model | Out-String ).Trim()
            $modelObject = [PSCustomObject]@{ 'Model' = $modelString }
            $SupportedCPUModelStrings += $modelObject
        }    

        Function CheckWin11CPUSupport($CPUString) {
            if (($($SupportedCPUModelStrings.Model) | ForEach-Object { $CPUstring.contains($_) }) -notcontains $true) { return $true } else { return $false }
        }
        $DevicesWithSupportedCPUs = $RMMData | Where-Object { CheckWin11CPUSupport($_.'CPU Description') }
        Write-Output "Counted $($DevicesWithSupportedCPUs.count) devices with CPUs that do not support Windows 11."
    }

    Write-Output "Finished in $($Stopwatch.Elapsed.TotalSeconds) seconds."
    Write-Output "Report output path is $ReportPath."
    $Stopwatch.Stop()
}