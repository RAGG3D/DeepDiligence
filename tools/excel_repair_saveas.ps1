param(
    [Parameter(Mandatory = $true)]
    [string]$Path
)

$ErrorActionPreference = "Stop"

$dir = [System.IO.Path]::GetDirectoryName($Path)
$stem = [System.IO.Path]::GetFileNameWithoutExtension($Path)
$tmp = [System.IO.Path]::Combine($dir, "$stem.normalized.xlsx")
if (Test-Path $tmp) { Remove-Item $tmp -Force }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $null

try {
    # CorruptLoad=1 is xlRepairFile. This normalizes strict OOXML edits through
    # Excel itself so the user never sees an Open and Repair prompt.
    $workbook = $excel.Workbooks.Open($Path, 0, $false, 5, "", "", $false, 1, "", $false, $false, 0, $false, $false, 1)
    # OOXML generators replace whole formula regions.  Excel can otherwise
    # retain a stale calc chain and leave downstream terminal-year values at
    # zero.  Rebuild every dependency before persisting cached results.
    $excel.Calculation = -4105 # xlCalculationAutomatic
    $excel.CalculateBeforeSave = $true
    foreach ($name in @(
        "FY DATA", "Scenarios", "Pipeline", "RIS", "RBS", "Schedules",
        "FSA", "RCFS", "VALUATION", "Catalyst"
    )) {
        if ($workbook.Worksheets.Item($name) -ne $null) {
            $workbook.Worksheets.Item($name).Calculate()
        }
    }
    $workbook.ForceFullCalculation = $false
    try { $workbook.FullCalculationOnLoad = $false } catch { }
    $workbook.SaveAs($tmp, 51)
    $workbook.Close($false)
    $workbook = $null
    Move-Item -Force $tmp $Path
    Write-Host "Excel repair-save normalized: $Path"
}
finally {
    if ($workbook -ne $null) {
        $workbook.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
