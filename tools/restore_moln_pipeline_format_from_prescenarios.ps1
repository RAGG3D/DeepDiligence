param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath,

    [Parameter(Mandatory = $true)]
    [string]$TargetPath
)

$ErrorActionPreference = "Stop"

$xlPasteFormats = -4122
$xlColorScale = 3
$xlLowestValue = 1
$xlHighestValue = 2
$xlExpression = 2
$xlNone = -4142
$white = 16777215

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$srcWb = $null
$dstWb = $null

function Copy-PageSetup($src, $dst) {
    $props = @(
        "Orientation", "PaperSize", "Zoom", "FitToPagesWide", "FitToPagesTall",
        "LeftMargin", "RightMargin", "TopMargin", "BottomMargin",
        "HeaderMargin", "FooterMargin", "CenterHorizontally", "CenterVertically",
        "PrintGridlines", "PrintHeadings", "Order", "BlackAndWhite",
        "Draft", "FirstPageNumber"
    )
    foreach ($p in $props) {
        try { $dst.PageSetup.$p = $src.PageSetup.$p } catch {}
    }
    try { $dst.PageSetup.PrintArea = $src.PageSetup.PrintArea } catch {}
    try { $dst.PageSetup.PrintTitleRows = $src.PageSetup.PrintTitleRows } catch {}
    try { $dst.PageSetup.PrintTitleColumns = $src.PageSetup.PrintTitleColumns } catch {}
    try { $dst.PageSetup.LeftHeader = $src.PageSetup.LeftHeader } catch {}
    try { $dst.PageSetup.CenterHeader = $src.PageSetup.CenterHeader } catch {}
    try { $dst.PageSetup.RightHeader = $src.PageSetup.RightHeader } catch {}
    try { $dst.PageSetup.LeftFooter = $src.PageSetup.LeftFooter } catch {}
    try { $dst.PageSetup.CenterFooter = $src.PageSetup.CenterFooter } catch {}
    try { $dst.PageSetup.RightFooter = $src.PageSetup.RightFooter } catch {}
}

function Add-MarketShareColorScale($range) {
    $scale = $range.FormatConditions.AddColorScale(2)
    $scale.ColorScaleCriteria.Item(1).Type = $xlLowestValue
    $scale.ColorScaleCriteria.Item(1).FormatColor.Color = $white
    $scale.ColorScaleCriteria.Item(2).Type = $xlHighestValue
    try {
        $scale.ColorScaleCriteria.Item(2).FormatColor.ThemeColor = 8
        $scale.ColorScaleCriteria.Item(2).FormatColor.TintAndShade = 0.799951170384838
    } catch {
        $scale.ColorScaleCriteria.Item(2).FormatColor.Color = 16247773
    }
}

try {
    $srcWb = $excel.Workbooks.Open($SourcePath, 0, $true)
    $dstWb = $excel.Workbooks.Open($TargetPath, 0, $false)

    $src = $srcWb.Worksheets.Item("Pipeline")
    $dst = $dstWb.Worksheets.Item("Pipeline")

    # Copy visible cell formatting for the live MOLN Pipeline area only.
    [void]$src.Range("A1:AH35").Copy()
    [void]$dst.Range("A1:AH35").PasteSpecial($xlPasteFormats)
    try { $excel.CutCopyMode = $false } catch {}

    # Match row heights and column layout for the same visible range.
    for ($c = 1; $c -le 34; $c++) {
        try { $dst.Columns.Item($c).ColumnWidth = $src.Columns.Item($c).ColumnWidth } catch {}
        try { $dst.Columns.Item($c).Hidden = $src.Columns.Item($c).Hidden } catch {}
        try { $dst.Columns.Item($c).OutlineLevel = $src.Columns.Item($c).OutlineLevel } catch {}
    }
    for ($r = 1; $r -le 35; $r++) {
        try { $dst.Rows.Item($r).RowHeight = $src.Rows.Item($r).RowHeight } catch {}
        try { $dst.Rows.Item($r).Hidden = $src.Rows.Item($r).Hidden } catch {}
        try { $dst.Rows.Item($r).OutlineLevel = $src.Rows.Item($r).OutlineLevel } catch {}
    }

    try { $dst.Tab.Color = $src.Tab.Color } catch {}
    try { $dst.StandardWidth = $src.StandardWidth } catch {}
    try { $dst.StandardHeight = $src.StandardHeight } catch {}
    Copy-PageSetup $src $dst

    # Rebuild conditional formatting semantically. The source workbook has stale
    # color scales on far blank rows, so only the valid live rows are recreated.
    try { $dst.Range("A1:AH35").FormatConditions.Delete() } catch {}
    foreach ($staleRange in @("S383:AH383", "S390:AH390", "S392:AH392", "S399:AH399", "S405:AH405", "S407:AH407")) {
        try { $dst.Range($staleRange).FormatConditions.Delete() } catch {}
    }

    [void]$src.Range("B1:B6").Copy()
    [void]$dst.Range("B1:B6").PasteSpecial($xlPasteFormats)
    try { $excel.CutCopyMode = $false } catch {}

    # If PasteSpecial did not transfer the source expression CF cleanly, create
    # the visible-equivalent rule explicitly.
    if ($dst.Range("B1:B6").FormatConditions.Count -eq 0) {
        $rule = $dst.Range("B1:B6").FormatConditions.Add($xlExpression, [Type]::Missing, '=B1<>""')
        try { $rule.Font.Bold = $true } catch {}
        try {
            $rule.Font.ThemeColor = 1
            $rule.Font.TintAndShade = 0.499984740745262
        } catch {}
        try {
            $rule.Interior.Pattern = 1
            $rule.Interior.ThemeColor = 5
            $rule.Interior.TintAndShade = 0.799920651875362
        } catch {}
        try { $rule.NumberFormat = '#,##0' } catch {}
    }

    $marketRows = @(11, 18, 20, 27, 29, 31)
    foreach ($row in $marketRows) {
        Add-MarketShareColorScale $dst.Range("F${row}:AH${row}")
    }

    $dstWb.Save()
    Write-Host "Restored Pipeline formatting from $SourcePath to $TargetPath"
}
finally {
    if ($srcWb -ne $null) {
        $srcWb.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($srcWb)
    }
    if ($dstWb -ne $null) {
        $dstWb.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($dstWb)
    }
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
