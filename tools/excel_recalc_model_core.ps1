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
    foreach ($name in @(
        "FY DATA", "Scenarios", "Pipeline", "RIS", "RBS", "Schedules",
        "FSA", "RCFS", "VALUATION"
    )) {
        $workbook.Worksheets.Item($name).Calculate()
    }
    $catalyst = $workbook.Worksheets.Item("Catalyst")
    $last = 10
    while ($last -lt 1000 -and $catalyst.Cells.Item($last,2).Value2 -ne $null) { $last++ }
    $last--
    if ($last -lt 9) { $last = 9 }
    $catalyst.Range("C9:C"+$last).Calculate()
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
    Write-Host "Excel core model recalculated through Catalyst: $Path"
}
finally {
    if ($workbook -ne $null) {
        $workbook.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
