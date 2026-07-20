param(
    [Parameter(Mandatory=$true)][string]$Path
)

$ErrorActionPreference = "Stop"
$excel = $null
$book = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $book = $excel.Workbooks.Open($Path, 0, $false)

    foreach ($sheetName in @("Catalyst", "Test-ASCO2026")) {
        $sheet = $null
        try { $sheet = $book.Worksheets.Item($sheetName) } catch { continue }

        # Rename only the main-table joint-product column.  Table 3 keeps its
        # target-level Conv. labels because those are the RJConv. inputs.
        for ($col = 2; $col -le 20; $col++) {
            $header = [string]$sheet.Cells.Item(7, $col).Value2
            $left = if ($col -gt 2) { [string]$sheet.Cells.Item(7, $col - 1).Value2 } else { "" }
            if (($header -eq "Conv." -or $header -eq "RJConv.") -and $left -eq "Upside") {
                $sheet.Cells.Item(7, $col).Value2 = "RJConv."
            }
        }

        if ($sheetName -eq "Test-ASCO2026") {
            $usedLast = $sheet.UsedRange.Row + $sheet.UsedRange.Rows.Count - 1
            for ($row = 1; $row -le $usedLast; $row++) {
                $label = [string]$sheet.Cells.Item($row, 2).Value2
                if ($label -eq "Highest-Conv. Scenario" -or $label -eq "Highest-RJConv. Scenario") {
                    $sheet.Cells.Item($row, 2).Value2 = "Highest-RJConv. Scenario"
                    $sheet.Cells.Item($row, 3).Value2 = "RJConv."
                }
            }
        }
    }

    $book.Save()
}
finally {
    if ($book -ne $null) { $book.Close($true) | Out-Null }
    if ($excel -ne $null) { $excel.Quit() }
    if ($book -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
    if ($excel -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
