param(
    [Parameter(Mandatory = $true)]
    [string]$Path
)

$ErrorActionPreference = "Stop"
$xlCalculationAutomatic = -4105
$xlDone = 0
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $null

try {
    $workbook = $excel.Workbooks.Open($Path, 0, $false)
    $excel.Calculation = $xlCalculationAutomatic
    $excel.CalculateBeforeSave = $true
    $excel.CalculateFullRebuild()
    $deadline = [DateTime]::UtcNow.AddMinutes(5)
    while ($excel.CalculationState -ne $xlDone -and [DateTime]::UtcNow -lt $deadline) {
        Start-Sleep -Milliseconds 250
    }
    if ($excel.CalculationState -ne $xlDone) {
        throw "Excel calculation did not finish within five minutes"
    }
    # Calculate generated sheets in dependency order.  Excel's full rebuild can
    # leave full-column SUMIFS formulas stale after their source sheet was
    # replaced at the OOXML level, even though Evaluate() returns the right
    # result.  Explicit sheet calculation refreshes the persisted cell cache.
    foreach ($name in @(
        "FY DATA", "Scenarios", "Pipeline", "RIS", "RBS", "Schedules",
        "FSA", "RCFS", "VALUATION", "Catalyst"
    )) {
        if ($workbook.Worksheets.Item($name) -ne $null) {
            $workbook.Worksheets.Item($name).Calculate()
        }
    }
    $excel.CalculateFullRebuild()
    $deadline = [DateTime]::UtcNow.AddMinutes(5)
    while ($excel.CalculationState -ne $xlDone -and [DateTime]::UtcNow -lt $deadline) {
        Start-Sleep -Milliseconds 250
    }
    if ($excel.CalculationState -ne $xlDone) {
        throw "Excel calculation did not finish within five minutes"
    }
    $workbook.ForceFullCalculation = $false
    try { $workbook.FullCalculationOnLoad = $false } catch { }
    $workbook.Save()
    Write-Host "Excel full recalculation saved: $Path"
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
