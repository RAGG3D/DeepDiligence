param(
    [Parameter(Mandatory = $true)]
    [string]$Path
)

$ErrorActionPreference = "Stop"

$xlNone = -4142
$xlColorIndexAutomatic = -4105
$blue = 16711680
$white = 16777215

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $null

try {
    $workbook = $excel.Workbooks.Open($Path, 0, $false)

    foreach ($ws in @($workbook.Worksheets)) {
        try {
            foreach ($comment in @($ws.Comments)) { $comment.Delete() }
            foreach ($threaded in @($ws.CommentsThreaded)) { $threaded.Delete() }
        } catch {}
    }

    $pipelineWs = $null
    try { $pipelineWs = $workbook.Worksheets.Item("Pipeline") } catch {}
    if ($pipelineWs -ne $null) {
        $ws = $pipelineWs
        $lastRow = $ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1
        if ($lastRow -lt 9) { $lastRow = 9 }
        for ($row = 9; $row -le $lastRow; $row++) {
            $a = "$($ws.Range("A$row").Text)".Trim()
            $c = "$($ws.Range("C$row").Text)".Trim()
            $dText = "$($ws.Range("D$row").Text)".Trim()
            $dFormula = "$($ws.Range("D$row").Formula)".Trim()
            $d = "$dText $dFormula".Trim()

            if ($a -eq "X" -and $c -eq "" -and $d -ne "") {
                $ws.Range("E${row}:AH${row}").NumberFormat = '"Stage "0;("Stage "0);""'
            }

            if ($c -ne "") {
                $ws.Range("C$row").Interior.Color = $white
                $ws.Range("C$row").Font.Color = $blue
            }

            $ws.Range("E$row").Interior.Pattern = $xlNone
            $ws.Range("E$row").Font.ColorIndex = $xlColorIndexAutomatic

            if ($d -like "*Market Share*") {
                $rng = $ws.Range("F${row}:AH${row}")
                $rng.NumberFormat = "0%"
                $rng.Font.Bold = $false
            }
            if ($d -like "*Revenue*" -or $d -like "*COGS*") {
                $rng = $ws.Range("D${row}:AH${row}")
                $rng.Font.Bold = $true
                $ws.Range("F${row}:AH${row}").NumberFormat = "#,##0.0;(#,##0.0);0.0"
            }
        }
        $ws.Columns("C:E").AutoFit() | Out-Null
        $ws.Columns("E:E").ColumnWidth = 12

    }

    foreach ($name in @("FY DATA", "FY DATA K USD")) {
        try {
            $ws = $workbook.Worksheets.Item($name)
            $ws.Columns("D:D").NumberFormat = "General"
            $ws.Columns("D:D").ColumnWidth = 68
            $ws.Columns("D:D").Font.ColorIndex = $xlColorIndexAutomatic
        } catch {}
    }

    try {
        $ws = $workbook.Worksheets.Item("Peer View")
        $ws.Columns("D:Q").ColumnWidth = 16
        $ws.Columns("I:I").ColumnWidth = 38
        $ws.Columns("P:Q").ColumnWidth = 22
        $ws.Range("D7:Q7").Font.Bold = $true
        $ws.Range("D7:Q7").Interior.Color = 14277081
        $ws.Range("D:Q").WrapText = $true
        $ws.Activate()
        $ws.Range("D8").Select()
        $excel.ActiveWindow.FreezePanes = $true
    } catch {}

    $excel.CalculateFullRebuild()
    $workbook.Save()
    Write-Host "Final Excel polish complete"
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
