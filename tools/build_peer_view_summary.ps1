param(
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [Parameter(Mandatory = $true)]
    [string]$Ticker,
    [Parameter(Mandatory = $true)]
    [string]$RowsJson
)

$ErrorActionPreference = "Stop"

$rows = Get-Content -Raw -Path $RowsJson | ConvertFrom-Json
$headers = @(
    "Indication", "Drug", "Ticker/Owner", "Rating", "Status/Phase",
    "Mechanism", "Treatment Line", "N", "ORR", "CR", "PFS", "OS",
    "Safety", "Date/Source"
)
$fields = @(
    "indication", "drug", "ticker", "rating", "status",
    "mechanism", "line", "n", "orr", "cr", "pfs", "os",
    "safety", "source"
)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
$excel.EnableEvents = $false
try { $excel.Calculation = -4135 } catch {}
$workbook = $null

try {
    $workbook = $excel.Workbooks.Open($Path, 0, $false)
    $ws = $workbook.Worksheets.Item("Peer View")
    $ws.Range("D5:Q160").UnMerge()
    $ws.Range("D5:Q160").ClearContents()
    $ws.Range("D5:Q160").Interior.Pattern = -4142
    $ws.Range("D5:Q160").Font.Bold = $false

    $ws.Cells.Item(5, 4).Value2 = "$($Ticker.ToUpper()) Peer View: Pipeline vs Indication Peers"
    $ws.Range("D5").Font.Bold = $true
    $ws.Range("D5").Font.Size = 12

    for ($i = 0; $i -lt $headers.Count; $i++) {
        $cell = $ws.Cells.Item(7, 4 + $i)
        $cell.Value2 = $headers[$i]
    }
    $ws.Range("D7:Q7").Font.Bold = $true
    $ws.Range("D7:Q7").Interior.Color = 14277081

    $r = 8
    foreach ($row in @($rows)) {
        for ($i = 0; $i -lt $fields.Count; $i++) {
            $value = $row.($fields[$i])
            if ($null -eq $value -or "$value" -eq "") { continue }
            $cell = $ws.Cells.Item($r, 4 + $i)
            $text = "$value"
            if ($text -match "^-?\d+(\.\d+)?$") {
                $cell.Value2 = [double]$text
            } else {
                $cell.Value2 = $text
            }
        }
        if ("$($row.is_company)" -eq "1") {
            $ws.Range("D${r}:Q${r}").Font.Bold = $true
            $ws.Range("D${r}:Q${r}").Interior.Color = 15921906
        }
        $r += 1
    }

    $lastRow = [Math]::Max(8, $r - 1)
    $ws.Range("L8:M$lastRow").NumberFormat = "0%"
    $ws.Range("N8:O$lastRow").NumberFormat = "0.0"
    $ws.Range("P8:P$lastRow").NumberFormat = "0%"
    $ws.Range("D:Q").WrapText = $true
    $ws.Columns("D:D").ColumnWidth = 12
    $ws.Columns("E:F").ColumnWidth = 20
    $ws.Columns("G:H").ColumnWidth = 14
    $ws.Columns("I:I").ColumnWidth = 44
    $ws.Columns("J:O").ColumnWidth = 14
    $ws.Columns("P:Q").ColumnWidth = 24
    $workbook.Save()
    Write-Host "Peer View summary written with Excel COM: $($rows.Count) rows"
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
