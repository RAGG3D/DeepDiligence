param(
    [Parameter(Mandatory = $true)]
    [string]$ReferencePath,

    [Parameter(Mandatory = $true)]
    [string]$TargetPath
)

$ErrorActionPreference = "Stop"

$xlNone = -4142
$xlSolid = 1
$xlRight = -4152
$xlCenter = -4108
$xlLowestValue = 1
$xlHighestValue = 2
$xlCalculationManual = -4135
$black = 0
$white = 16777215
$blue = 16711680
$lightYellow = 13434879

$stageFormat = '"Stage" #,##0'
$numberFormat = '#,##0'
$percentFormat = '#,##0%'
$priceFormat = '#,##0.00'

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try { $excel.Calculation = $xlCalculationManual } catch {}
try { $excel.CalculateBeforeSave = $false } catch {}
$refWb = $null
$targetWb = $null

function Set-NoFill($range) {
    try { $range.Interior.Pattern = $xlNone } catch {}
}

function Set-Font($range, [bool]$bold, [bool]$italic, [int]$color) {
    try { $range.Font.Bold = $bold } catch {}
    try { $range.Font.Italic = $italic } catch {}
    try { $range.Font.Color = $color } catch {}
}

function Set-FillFromReference($range, $referenceCell) {
    try { $range.Interior.Pattern = $referenceCell.Interior.Pattern } catch {}
    try {
        $range.Interior.ThemeColor = $referenceCell.Interior.ThemeColor
        $range.Interior.TintAndShade = $referenceCell.Interior.TintAndShade
    } catch {
        try { $range.Interior.Color = $referenceCell.Interior.Color } catch {}
    }
}

function Set-YearBase($range, [string]$numFmt, [bool]$bold) {
    try { $range.NumberFormat = $numFmt } catch {}
    Set-Font $range $bold $false $black
    try { $range.HorizontalAlignment = $xlRight } catch {}
}

function Set-StageFill($range, $referenceCell) {
    try {
        $range.Interior.Pattern = $xlSolid
        $range.Interior.ThemeColor = $referenceCell.Interior.ThemeColor
        $range.Interior.TintAndShade = $referenceCell.Interior.TintAndShade
    } catch {
        try { $range.Interior.Color = 16247773 } catch {}
    }
}

function Set-LabelNormal($cell) {
    try { $cell.NumberFormat = "General" } catch {}
    Set-NoFill $cell
    Set-Font $cell $false $false $black
}

function Set-LabelBold($cell) {
    try { $cell.NumberFormat = "General" } catch {}
    Set-NoFill $cell
    Set-Font $cell $true $false $black
}

function Set-LabelDrug($cell) {
    try { $cell.NumberFormat = "General" } catch {}
    Set-NoFill $cell
    Set-Font $cell $false $true $black
}

function Set-CBlueNoFill($cell) {
    Set-NoFill $cell
    Set-Font $cell $false $false $blue
}

function Add-MarketShareColorScale($range, $stageReferenceCell) {
    $scale = $range.FormatConditions.AddColorScale(2)
    $scale.ColorScaleCriteria.Item(1).Type = $xlLowestValue
    $scale.ColorScaleCriteria.Item(1).FormatColor.Color = $white
    $scale.ColorScaleCriteria.Item(2).Type = $xlHighestValue
    try {
        $scale.ColorScaleCriteria.Item(2).FormatColor.ThemeColor = $stageReferenceCell.Interior.ThemeColor
        $scale.ColorScaleCriteria.Item(2).FormatColor.TintAndShade = $stageReferenceCell.Interior.TintAndShade
    } catch {
        $scale.ColorScaleCriteria.Item(2).FormatColor.Color = 16247773
    }
}

try {
    $refWb = $excel.Workbooks.Open($ReferencePath, 0, $true)
    $targetWb = $excel.Workbooks.Open($TargetPath, 0, $false)
    try { $excel.Calculation = $xlCalculationManual } catch {}
    try { $excel.CalculateBeforeSave = $false } catch {}
    try { $targetWb.ForceFullCalculation = $false } catch {}
    try { $targetWb.PrecisionAsDisplayed = $false } catch {}
    $refWs = $refWb.Worksheets.Item("Pipeline")
    $ws = $targetWb.Worksheets.Item("Pipeline")

    $used = $ws.UsedRange
    $lastRow = $used.Row + $used.Rows.Count - 1
    if ($lastRow -lt 9) { $lastRow = 9 }

    # Normalize the target Pipeline span to its own full forecast horizon.
    $yearStartCol = "F"
    $yearEndCol = "AH"
    $stageRef = $refWs.Range("F9")

    # Clear all legacy/misplaced Pipeline conditional formatting first,
    # including stale color scales on rows in the 380s/400s.
    try { $ws.Cells.FormatConditions.Delete() } catch {}

    # Section header row: preserve target text/formulas, fix style only.
    $sectionRow = 6
    $sectionRange = $ws.Range("D${sectionRow}:${yearEndCol}${sectionRow}")
    $sectionReference = $refWs.Range("D6")
    $sectionRange.NumberFormat = "General"
    Set-FillFromReference $sectionRange $sectionReference
    Set-Font $sectionRange $true $false $white

    # The revenue line header is not a per-drug row, but it is part of the
    # visible Pipeline table header and should follow the CMPX style.
    $revenueHeaderRow = 8
    $revenueHeaderD = $ws.Range("D${revenueHeaderRow}")
    $revenueHeaderE = $ws.Range("E${revenueHeaderRow}")
    $revenueHeaderRefD = $refWs.Range("D8")
    $revenueHeaderRefE = $refWs.Range("E8")
    $revenueHeaderRefF = $refWs.Range("F8")
    $revenueHeaderD.NumberFormat = $revenueHeaderRefD.NumberFormat
    Set-NoFill $revenueHeaderD
    Set-Font $revenueHeaderD $true $true $black
    $revenueHeaderE.NumberFormat = $revenueHeaderRefE.NumberFormat
    Set-NoFill $revenueHeaderE
    Set-Font $revenueHeaderE $false $false $black
    $revenueHeaderYears = $ws.Range("${yearStartCol}${revenueHeaderRow}:${yearEndCol}${revenueHeaderRow}")
    $revenueHeaderYears.NumberFormat = $revenueHeaderRefF.NumberFormat
    Set-NoFill $revenueHeaderYears
    Set-Font $revenueHeaderYears $false $true $black

    # Preserve top helper-row styles where relevant, but keep this script
    # scoped to the revenue-forecasting block below.
    for ($row = 9; $row -le $lastRow; $row++) {
        $dText = "$($ws.Range("D$row").Text)".Trim()
        $dFormula = "$($ws.Range("D$row").Formula)".Trim()
        $dCombined = "$dText $dFormula"
        $cText = "$($ws.Range("C$row").Text)".Trim()
        $aText = "$($ws.Range("A$row").Text)".Trim()
        $bFormula = "$($ws.Range("B$row").Formula)".Trim()

        $yearRange = $ws.Range("${yearStartCol}${row}:${yearEndCol}${row}")
        $dCell = $ws.Range("D$row")
        $cCell = $ws.Range("C$row")

        if ($dCombined -like "*Market Share*") {
            Set-LabelNormal $dCell
            Set-CBlueNoFill $cCell
            Set-YearBase $yearRange $percentFormat $false
            Set-NoFill $yearRange
            Add-MarketShareColorScale $yearRange $stageRef
        } elseif ($dCombined -like "*List Price*") {
            Set-LabelNormal $dCell
            Set-YearBase $yearRange $priceFormat $false
            $yearRange.Interior.Pattern = $xlSolid
            $yearRange.Interior.Color = $lightYellow
        } elseif ($dCombined -like "* Revenue*") {
            Set-LabelBold $dCell
            Set-YearBase $yearRange $numberFormat $true
            Set-NoFill $yearRange
        } elseif ($dCombined -like "* COGS*") {
            Set-LabelBold $dCell
            Set-YearBase $yearRange $numberFormat $true
            Set-NoFill $yearRange
        } elseif ($dCombined -like "* TAM*") {
            Set-LabelNormal $dCell
            Set-CBlueNoFill $cCell
            Set-YearBase $yearRange $numberFormat $false
            Set-NoFill $yearRange
        } elseif ($aText -eq "X" -and $cText -eq "" -and $dFormula -notlike "=*" -and $dText -ne "") {
            Set-LabelDrug $dCell
            Set-YearBase $yearRange $stageFormat $false
            Set-StageFill $yearRange $stageRef
        } elseif ($dText -eq "" -and $bFormula -like "=B*") {
            $ws.Range("B$row").Interior.Pattern = $xlSolid
            try {
                $ws.Range("B$row").Interior.ThemeColor = 2
                $ws.Range("B$row").Interior.TintAndShade = 0
            } catch {}
            $yearRange.NumberFormat = "General"
        }
    }

    $targetWb.Save()
    Write-Host "Pipeline styling fixed in $TargetPath"
}
finally {
    if ($refWb -ne $null) {
        $refWb.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($refWb)
    }
    if ($targetWb -ne $null) {
        $targetWb.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($targetWb)
    }
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
