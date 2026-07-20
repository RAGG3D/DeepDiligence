param(
    [Parameter(Mandatory = $true)]
    [string]$Path
)

$ErrorActionPreference = "Stop"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $null

try {
    $workbook = $excel.Workbooks.Open($Path, 0, $false)
    $sheet = $workbook.Worksheets.Item("Scenarios")
    $lastRow = $sheet.UsedRange.Rows.Count
    $patched = 0
    for ($row = 1; $row -le $lastRow; $row++) {
        if ($sheet.Cells.Item($row, 1).Value2 -ne 4) { continue }
        if ([string]$sheet.Cells.Item($row, 2).Value2 -notmatch "Absolute") { continue }
        if ($sheet.Cells.Item($row, 4).Value2 -ne $null) { continue }
        $approved = $false
        for ($col = 5; $col -le 24; $col++) {
            $cell = $sheet.Cells.Item($row, $col)
            if ($cell.Value2 -eq 5) { $approved = $true; continue }
            if ($approved -and $cell.Value2 -eq $null) {
                $cell.Value2 = 5
                $patched++
            }
        }
    }
    $sheet.Calculate()
    $workbook.Save()
    Write-Host "Persisted Stage 5 approval states: $patched cells"
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
