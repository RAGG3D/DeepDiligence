param(
    [Parameter(Mandatory = $true)]
    [string]$Path
)

$ErrorActionPreference = "Stop"

$xlChartTypeWaterfall = 119
$xlChartTypeStockOHLC = 89
$xlChartTypeColumnStacked = 52
$xlChartTypeBarStacked = 58
$xlColumns = 2
$xlNone = -4142
$msoShapeRectangle = 1
$msoFalse = 0
$msoTrue = -1
$msoLineDash = 4
$xlCalculationManual = -4135

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $null
$bootstrapWb = $null

try {
    # Calculation mode is application-wide.  Set it before opening the model so
    # Excel does not launch a full 56-scenario What-If recalculation merely to
    # create drawing objects.  Cached table values are populated by the explicit
    # repair/recalc step that precedes this chart-only worker.
    $bootstrapWb = $excel.Workbooks.Add()
    $excel.Calculation = $xlCalculationManual
    $excel.CalculateBeforeSave = $false
    $bootstrapWb.Close($false)
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($bootstrapWb)
    $bootstrapWb = $null
    $workbook = $excel.Workbooks.Open($Path, 0, $false)
    # Opening a workbook can reassert its saved calculation mode/properties.
    # Reset them on the live model as well so Save() remains drawing-only.
    $excel.Calculation = $xlCalculationManual
    $excel.CalculateBeforeSave = $false
    $workbook.ForceFullCalculation = $false
    try { $workbook.FullCalculationOnLoad = $false } catch { }

    $ws = $workbook.Worksheets.Item("VALUATION")

    foreach ($co in @($ws.ChartObjects())) {
        if ($co.Name -eq "Valuation Football Field" -or $co.Name -eq "Valuation Waterfall") {
            $co.Delete()
        }
    }
    foreach ($shape in @($ws.Shapes)) {
        $shapeText = ""
        try { $shapeText = "$($shape.TextFrame2.TextRange.Text)".Trim() } catch {}
        if ($shape.Name -eq "Valuation Football Field Frame" -or
            $shape.Name -eq "Valuation Waterfall Frame") {
            $shape.Delete()
        }
    }

    $lastWaterfall = 26
    for ($row = 26; $row -le 60; $row++) {
        $label = $ws.Range("F$row").Text
        $scenarioId = $ws.Range("I$row").Value2
        if (($label -ne $null -and "$label".Trim() -ne "") -or $scenarioId -ne $null) {
            $lastWaterfall = $row
        }
    }

    # Remove stale helper/output formatting.  Visible helper blocks are not
    # allowed; chart-specific helper series live in hidden AA:AI only.
    $ws.Range("L64:O73").ClearContents()
    # Clear the yellow input fill only on trailing empties BELOW the last live
    # waterfall row (dynamic per ticker), never on the live scenario-id cells.
    if ($lastWaterfall -lt 60) {
        $ws.Range(("I{0}:I60" -f ($lastWaterfall + 1))).Interior.Pattern = $xlNone
    }
    $ws.Range("AA1:AI120").Clear()
    $ws.Columns("AA:AI").Hidden = $false

    function Add-DashedFrame($name, $left, $top, $width, $height) {
        $frame = $ws.Shapes.AddShape($msoShapeRectangle, $left - 6, $top - 6, $width + 12, $height + 12)
        $frame.Name = $name
        [void]($frame.Fill.Visible = $msoFalse)
        [void]($frame.Line.Visible = $msoTrue)
        [void]($frame.Line.ForeColor.RGB = 8421504)
        [void]($frame.Line.Weight = 1)
        [void]($frame.Line.DashStyle = $msoLineDash)
        [void]$frame.ZOrder(1)
    }

    function Get-PlaceholderBounds($needle, $fallbackCell, $fallbackWidth, $fallbackHeight) {
        foreach ($shape in @($ws.Shapes)) {
            $shapeText = ""
            try { $shapeText = "$($shape.TextFrame2.TextRange.Text)".Trim() } catch {}
            if ($shapeText -eq $needle) {
                return @{
                    Left = [double]$shape.Left
                    Top = [double]$shape.Top
                    Width = [double]$shape.Width
                    Height = [double]$shape.Height
                }
            }
        }
        return @{
            Left = [double]$ws.Range($fallbackCell).Left
            Top = [double]$ws.Range($fallbackCell).Top
            Width = [double]$fallbackWidth
            Height = [double]$fallbackHeight
        }
    }

    $invariantCulture = [System.Globalization.CultureInfo]::InvariantCulture
    function Set-CellNumber($address, $value) {
        $ws.Range($address).Formula = [string]::Format($invariantCulture, "{0}", [double]$value)
    }

    $footballBox = Get-PlaceholderBounds "Valuation Football" "L11" 540 200
    $waterfallBox = Get-PlaceholderBounds "Waterflow Chart" "L28" 540 250
    foreach ($shape in @($ws.Shapes)) {
        $shapeText = ""
        try { $shapeText = "$($shape.TextFrame2.TextRange.Text)".Trim() } catch {}
        if ($shapeText -eq "Valuation Football" -or $shapeText -eq "Waterflow Chart") {
            $shape.Delete()
        }
    }

    Add-DashedFrame "Valuation Football Field Frame" $footballBox.Left $footballBox.Top $footballBox.Width $footballBox.Height

    # Football field: horizontal floating range bars from the existing F13:J21
    # table.  Hidden helper columns decompose Low→25th→75th→High into stacked
    # segments so it renders like a valuation football/box range, not an OHLC
    # stock chart.
    $ws.Range("AA12").Value2 = "Metric"
    $ws.Range("AB12").Value2 = "LowBase"
    $ws.Range("AC12").Value2 = "LowTo25"
    $ws.Range("AD12").Value2 = "25To75"
    $ws.Range("AE12").Value2 = "75ToHigh"
    $footballLast = 21
    for ($row = 13; $row -le $footballLast; $row++) {
        $outRow = $row
        $labelValue = $ws.Range("F" + [string]$row).Value2
        if ($labelValue -eq $null) { $labelValue = $ws.Range("F" + [string]$row).Text }
        $label = [string]$labelValue
        $q1 = $ws.Range("G" + [string]$row).Value2
        $high = $ws.Range("H" + [string]$row).Value2
        $low = $ws.Range("I" + [string]$row).Value2
        $q3 = $ws.Range("J" + [string]$row).Value2
        if ($q1 -eq $null) { $q1 = 0 }
        if ($high -eq $null) { $high = 0 }
        if ($low -eq $null) { $low = 0 }
        if ($q3 -eq $null) { $q3 = 0 }
        $q1 = [double]$q1; $high = [double]$high; $low = [double]$low; $q3 = [double]$q3
        if ($high -lt $low) { $tmp = $high; $high = $low; $low = $tmp }
        if ($q1 -lt $low) { $q1 = $low }
        if ($q3 -lt $q1) { $q3 = $q1 }
        if ($q3 -gt $high) { $q3 = $high }
        $ws.Range("AA" + [string]$outRow).Value2 = $label
        Set-CellNumber ("AB" + [string]$outRow) $low
        Set-CellNumber ("AC" + [string]$outRow) ([Math]::Max(0, $q1 - $low))
        Set-CellNumber ("AD" + [string]$outRow) ([Math]::Max(0, $q3 - $q1))
        Set-CellNumber ("AE" + [string]$outRow) ([Math]::Max(0, $high - $q3))
    }

    $co = $ws.ChartObjects().Add($footballBox.Left + 8, $footballBox.Top + 8, $footballBox.Width - 16, $footballBox.Height - 16)
    $co.Name = "Valuation Football Field"
    $chart = $co.Chart
    $chart.ChartType = $xlChartTypeBarStacked
    $chart.SetSourceData($ws.Range("AA12:AE$footballLast"), $xlColumns)
    [void]($chart.PlotVisibleOnly = $false)
    try {
        $chart.SeriesCollection(1).Format.Fill.Visible = $msoFalse
        $chart.SeriesCollection(1).Format.Line.Visible = $msoFalse
        $chart.SeriesCollection(2).Format.Fill.ForeColor.RGB = 12632256
        $chart.SeriesCollection(3).Format.Fill.ForeColor.RGB = 4473924
        $chart.SeriesCollection(4).Format.Fill.ForeColor.RGB = 12632256
    } catch {}
    try {
        $chart.HasTitle = $true
        $chart.ChartTitle.Text = "Valuation Football Field"
        $chart.Axes(1).TickLabels.Font.Size = 8
        $chart.Axes(2).TickLabels.NumberFormat = "0.0"
        $chart.HasLegend = $false
    } catch {}

    if ($lastWaterfall -ge 26) {
        Add-DashedFrame "Valuation Waterfall Frame" $waterfallBox.Left $waterfallBox.Top $waterfallBox.Width $waterfallBox.Height

        # Waterfall fallback: manually sourced stacked columns with an
        # invisible cumulative base and visible delta.  This avoids Excel COM
        # builds that silently create an empty native waterfall chart.
        $ws.Range("AG25").Value2 = "Component"
        $ws.Range("AH25").Value2 = "Base"
        $ws.Range("AI25").Value2 = "Delta"
        $running = 0.0
        for ($row = 26; $row -le $lastWaterfall; $row++) {
            $labelValue = $ws.Range("F" + [string]$row).Value2
            if ($labelValue -eq $null) { $labelValue = $ws.Range("F" + [string]$row).Text }
            $label = [string]$labelValue
            $delta = $ws.Range("G" + [string]$row).Value2
            if ($delta -eq $null) { $delta = 0 }
            $delta = [double]$delta
            $ws.Range("AG" + [string]$row).Value2 = $label
            if ($row -eq $lastWaterfall) {
                Set-CellNumber ("AH" + [string]$row) 0
                Set-CellNumber ("AI" + [string]$row) $delta
            } else {
                Set-CellNumber ("AH" + [string]$row) $running
                Set-CellNumber ("AI" + [string]$row) $delta
                $running += $delta
            }
        }

        $co2 = $ws.ChartObjects().Add($waterfallBox.Left + 8, $waterfallBox.Top + 8, $waterfallBox.Width - 16, $waterfallBox.Height - 16)
        $co2.Name = "Valuation Waterfall"
        $chart2 = $co2.Chart
        $chart2.ChartType = $xlChartTypeColumnStacked
        $chart2.SetSourceData($ws.Range("AG25:AI$lastWaterfall"), $xlColumns)
        [void]($chart2.PlotVisibleOnly = $false)
        $baseSeries = $chart2.SeriesCollection(1)
        $deltaSeries = $chart2.SeriesCollection(2)
        try {
            $baseSeries.Format.Fill.Visible = $msoFalse
            $baseSeries.Format.Line.Visible = $msoFalse
            $deltaSeries.Format.Fill.ForeColor.RGB = 4473924
        } catch {}
        try {
            $chart2.HasTitle = $true
            $chart2.ChartTitle.Text = "Valuation Waterfall"
            $chart2.Axes(1).TickLabels.Font.Size = 8
            $chart2.Axes(2).TickLabels.NumberFormat = "0.0"
            $chart2.HasLegend = $false
        } catch {}
    }

    $ws.Columns("AA:AI").Hidden = $true

    $workbook.Save()
    Write-Host "Valuation charts rebuilt: waterfall rows 26-$lastWaterfall"
}
finally {
    if ($bootstrapWb -ne $null) {
        $bootstrapWb.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($bootstrapWb)
    }
    if ($workbook -ne $null) {
        $workbook.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
