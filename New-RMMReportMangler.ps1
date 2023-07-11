#New-RMMReportMangler
Param (
    [Parameter(ValueFromPipelineByPropertyName)]
    [string] $ReportPath
)

$reportExists = Test-Path $ReportPath
if (!($reportExists)) { Write-Warning "Specified report file $ReportFile does not exist or could not be read. Exiting."; exit } else {

    Import-Module ImportExcel

    # Open Excel File
    $excel = open-excelpackage $ReportPath

    # Remove cover sheet
    $CoverSheetName = "Report Header"
    if ($excel.Workbook.Worksheets[$CoverSheetName]) { $excel.Workbook.Worksheets.Delete($CoverSheetName) }

    # Set Worksheet
    $ws = $excel.Workbook.Worksheets["Network Hardware"]

    # Trim extra blank row and column
    $cell = $ws.Cells[1, 1]
    if ($cell.value -ne 'Customer Name') {
        $ws.DeleteRow(1)
        $ws.DeleteColumn(18)
        $ws.DeleteColumn(17)
        $ws.DeleteColumn(16)
        $ws.DeleteColumn(1)
    }

    # Check whether the sheet 
    $cell = $ws.Cells[1, 1]
    if ($cell.value -eq 'Customer Name') {
        Write-Output "OK so far."
        # Save File
        Close-ExcelPackage $excel -SaveAs $ReportPath
    }

}