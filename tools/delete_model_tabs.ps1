param(
    [Parameter(Mandatory = $true)]
    [string]$Path,

    [Parameter(Mandatory = $true)]
    [string]$SheetsCsv,

    [string]$TickerLabel = "",
    [string]$WelcomeCellsCsv = "",
    [string]$GreenCellsCsv = ""
)

$ErrorActionPreference = "Stop"
$Sheets = $SheetsCsv -split '\|'

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false

    $workbook = $excel.Workbooks.Open($Path, 0, $false)
    if ($WelcomeCellsCsv -ne "" -and $TickerLabel -ne "") {
        foreach ($item in ($WelcomeCellsCsv -split '\|')) {
            $bang = $item.LastIndexOf("!")
            if ($bang -le 0) { continue }
            $sheetName = $item.Substring(0, $bang)
            $address = $item.Substring($bang + 1)
            $workbook.Worksheets.Item($sheetName).Range($address).Value2 = $TickerLabel
        }
    }
    if ($GreenCellsCsv -ne "") {
        $pipeline = $workbook.Worksheets.Item("Pipeline")
        foreach ($address in ($GreenCellsCsv -split '\|')) {
            if ($address -ne "") {
                # Excel/VBA RGB(0,97,0) = 24832 (dark green).
                $pipeline.Range($address).Font.Color = 24832
            }
        }
    }
    foreach ($sheetName in $Sheets) {
        $sheet = $null
        try {
            $sheet = $workbook.Worksheets.Item($sheetName)
        } catch {
            $sheet = $null
        }
        if ($null -ne $sheet) {
            $sheet.Delete()
        }
    }
    $workbook.Save()
} finally {
    if ($null -ne $workbook) {
        $workbook.Close($true) | Out-Null
    }
    if ($null -ne $excel) {
        $excel.DisplayAlerts = $true
        $excel.Quit() | Out-Null
    }
    if ($null -ne $workbook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null 2>$null
    }
    if ($null -ne $excel) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null 2>$null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
