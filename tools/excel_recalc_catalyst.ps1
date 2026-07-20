param(
    [Parameter(Mandatory = $true)]
    [string]$Path
)

$ErrorActionPreference = "Stop"
$xlCalculationManual = -4135
$xlCalculationAutomatic = -4105
$xlDone = 0
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $null

try {
    $workbook = $excel.Workbooks.Open($Path, 0, $false)
    $excel.Calculation = $xlCalculationManual
    $valuation = $workbook.Worksheets.Item("VALUATION")
    $catalyst = $workbook.Worksheets.Item("Catalyst")

    # Calculate only the new two-input scenario data table.  A workbook-wide
    # FullRebuild also recomputes every historical sensitivity table and can be
    # unnecessarily slow for a large DCF model.
    $lastScenarioRow = 9
    for ($r = 10; $r -le 1000; $r++) {
        if ($catalyst.Cells.Item($r,2).Value2 -eq $null) { break }
        $lastScenarioRow = $r
    }
    $catalyst.Range("C9:C"+$lastScenarioRow).Calculate()
    $catalyst.Calculate()

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
    $workbook.ForceFullCalculation = $false
    try { $workbook.FullCalculationOnLoad = $false } catch { }
    $workbook.Save()
    Write-Host "Excel Catalyst data table recalculated: $Path"
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
