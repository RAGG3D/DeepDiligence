param(
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [Parameter(Mandatory = $true)]
    [string]$ReferencePath,
    [Parameter(Mandatory = $true)]
    [string]$PipelineReferencePath,
    [switch]$PipelineOnly,
    [switch]$ScenariosOnly
)

$ErrorActionPreference = "Stop"
$xlPasteFormats = -4122
$xlCalculationManual = -4135
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
$excel.EnableEvents = $false
$targetWb = $null
$refWb = $null
$pipelineRefWb = $null
$bootstrapWb = $null

function Copy-Formats($source, $sourceRange, $target, $targetRange) {
    [void]$target.Range($targetRange).ClearFormats()
    [void]$source.Range($sourceRange).Copy()
    [void]$target.Range($targetRange).PasteSpecial($xlPasteFormats)
}

function Copy-Dimensions($source, $target, [int]$rows, [int]$cols) {
    for ($r=1; $r -le $rows; $r++) {
        $target.Rows.Item($r).RowHeight = $source.Rows.Item($r).RowHeight
        $target.Rows.Item($r).Hidden = $source.Rows.Item($r).Hidden
    }
    for ($c=1; $c -le $cols; $c++) {
        $target.Columns.Item($c).ColumnWidth = $source.Columns.Item($c).ColumnWidth
        $target.Columns.Item($c).Hidden = $source.Columns.Item($c).Hidden
    }
}

function Copy-RowFormat($source, [int]$sourceRow, $target, [int]$targetRow, [string]$lastCol) {
    Copy-Formats $source ("A"+$sourceRow+":"+$lastCol+$sourceRow) $target ("A"+$targetRow+":"+$lastCol+$targetRow)
    $target.Rows.Item($targetRow).RowHeight = $source.Rows.Item($sourceRow).RowHeight
}

function Rating-Key([string]$rating) {
    if ($rating -match "BIC|Best") { return "BIC" }
    if ($rating -match "T1|Tier One") { return "T1" }
    return "AVG"
}

function Col-Letter([int]$n) {
    $result = ""
    while ($n -gt 0) {
        $n--
        $result = [char](65 + ($n % 26)) + $result
        $n = [math]::Floor($n / 26)
    }
    return $result
}

try {
    $bootstrapWb = $excel.Workbooks.Add()
    $excel.Calculation = $xlCalculationManual
    $bootstrapWb.Close($false)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($bootstrapWb)
    $bootstrapWb = $null
    $targetWb = $excel.Workbooks.Open($Path, 0, $false)
    # Opening the model can reassert its saved automatic calculation state.
    # Restore manual mode on the live workbook so PasteSpecial/Save remains a
    # format-only operation even with large What-If tables.
    $excel.Calculation = $xlCalculationManual
    $excel.CalculateBeforeSave = $false
    $targetWb.ForceFullCalculation = $false
    try { $targetWb.FullCalculationOnLoad = $false } catch { }
    $excel.CalculateBeforeSave = $false
    if (-not $PipelineOnly) {
        $refWb = $excel.Workbooks.Open($ReferencePath, 0, $true)
    }
    if (-not $ScenariosOnly) {
        $pipelineRefWb = $excel.Workbooks.Open($PipelineReferencePath, 0, $true)
    }

    if ($ScenariosOnly) {
        $src=$refWb.Worksheets.Item("Scenarios"); $dst=$targetWb.Worksheets.Item("Scenarios")
        $last=$dst.UsedRange.Row+$dst.UsedRange.Rows.Count-1
        $divider=$null
        for ($r=10; $r -le $last; $r++) {
            if ([string]$dst.Cells.Item($r,3).Value2 -eq "Catalyst Scenarios") {
                $divider=$r
                break
            }
        }
        if ($null -eq $divider) { throw "Catalyst Scenarios divider not found" }

        # Framework v3 already preserves every non-header semantic row.  Only
        # the generated Catalyst title rows and the spacer immediately before
        # each later title differ from the approved source.  Restricting the
        # repair to those rows avoids thousands of slow Excel COM reads.
        for ($r=$divider+1; $r -le $last; $r++) {
            $b=$dst.Cells.Item($r,2).Value2
            $c=[string]$dst.Cells.Item($r,3).Value2
            if (($b -is [double] -or $b -is [int]) -and $c) {
                Copy-RowFormat $src 20 $dst $r "AE"
                if ($r -gt $divider+1) { Copy-RowFormat $src 19 $dst ($r-1) "AE" }
            }
        }
        [void]$src.Range("AA10").Copy(); [void]$dst.Range("AA12").PasteSpecial($xlPasteFormats)
        [void]$src.Range("AA10").Copy(); [void]$dst.Range("AA15").PasteSpecial($xlPasteFormats)
        $targetWb.Save()
        Write-Host "Applied semantic Scenarios reference styles only: $ReferencePath"
        return
    }

    if ($PipelineOnly) {
        $src=$pipelineRefWb.Worksheets.Item("Pipeline"); $dst=$targetWb.Worksheets.Item("Pipeline")
        [void]$dst.Cells.FormatConditions.Delete()
        Copy-Formats $src "A1:AH8" $dst "A1:AH8"
        Copy-Dimensions $src $dst 8 34
        $last=$dst.UsedRange.Row+$dst.UsedRange.Rows.Count-1
        for ($r=9; $r -le $last; $r++) {
            $a=[string]$dst.Cells.Item($r,1).Value2
            $d=[string]$dst.Cells.Item($r,4).Value2
            $formula=[string]$dst.Cells.Item($r,4).Formula
            if (-not $d -and -not $formula) { $sourceRow=15 }
            elseif (($d+$formula) -like "*Market Share*") { $sourceRow=11 }
            elseif (($d+$formula) -like "* TAM*") { $sourceRow=10 }
            elseif (($d+$formula) -like "*List Price*") { $sourceRow=12 }
            elseif (($d+$formula) -like "*Revenue*") { $sourceRow=13 }
            elseif (($d+$formula) -like "*COGS*") { $sourceRow=14 }
            elseif ($a -eq "X") { $sourceRow=9 }
            else { $sourceRow=15 }
            Copy-RowFormat $src $sourceRow $dst $r "AH"
        }
        for ($cc=1; $cc -le 34; $cc++) {
            $dst.Columns.Item($cc).ColumnWidth=$src.Columns.Item($cc).ColumnWidth
            $dst.Columns.Item($cc).Hidden=$src.Columns.Item($cc).Hidden
        }
        $targetWb.Save()
        Write-Host "Applied locked Pipeline reference styles only: $PipelineReferencePath"
        return
    }

    # Coordinate-stable tabs: MOLN is the visual authority; values/formulas are
    # untouched because only PasteSpecial(xlPasteFormats) is used.
    $static = @(
        @("Earnings", "A1:H18", 18, 8),
        @("VALUATION", "A1:P73", 73, 16),
        @("FSA", "A1:Y132", 132, 25),
        @("RIS", "A1:Y98", 98, 25),
        @("RBS", "A1:Y67", 67, 25),
        @("RCFS", "A1:W69", 69, 23),
        @("Schedules", "A1:X73", 73, 24),
        @("FY DATA", "A1:K143", 143, 11),
        @("FY DATA K USD", "A1:L130", 130, 12),
        @("Historical Events", "A1:AR374", 374, 44),
        @("BBG DAPI", "A1:AV113", 113, 48)
    )
    foreach ($spec in $static) {
        $name=$spec[0]; $range=$spec[1]; $rows=[int]$spec[2]; $cols=[int]$spec[3]
        $src=$refWb.Worksheets.Item($name); $dst=$targetWb.Worksheets.Item($name)
        # PasteSpecial does not overwrite a destination's old formatting when
        # the reference cell uses Excel's default style.  Clear first so blank
        # cells, decorative tails and previously blue formula cells are also
        # normalized to the MOLN reference.
        [void]$dst.Range($range).ClearFormats()
        Copy-Formats $src $range $dst $range
        Copy-Dimensions $src $dst $rows $cols
    }
    # Excel What-If tables can retain the input-formula blue font through a
    # range PasteSpecial.  These two VALUATION formulas are black in MOLN.
    $targetWb.Worksheets.Item("VALUATION").Range("C6:C7").Font.Color = `
        $refWb.Worksheets.Item("VALUATION").Range("C6").Font.Color
    # The embedded Catalyst What-If table contains price/share outputs. The legacy
    # MOLN cell format called L4 "Abs. Revenue" and rounded M5 to an integer,
    # which obscured both the unit bug and legitimate per-share precision.
    $valuation=$targetWb.Worksheets.Item("VALUATION")
    $valuation.Range("L4").NumberFormat = '"Abs. Value"'
    $valuation.Range("M5:M60").NumberFormat = "0.00"
    $targetWb.Worksheets.Item("Catalyst").Range("C9:C520").NumberFormat = "0.00"
    # Styled blank tail rows in FY DATA K USD are not part of the reference
    # model and should not create phantom sections in new tickers.
    $fyK=$targetWb.Worksheets.Item("FY DATA K USD")
    if ($fyK.UsedRange.Rows.Count -gt 130) { [void]$fyK.Range("A131:L"+$fyK.UsedRange.Rows.Count).ClearFormats() }

    # Scenarios: semantic row styles allow any number of assets/outcomes.
    $src=$refWb.Worksheets.Item("Scenarios"); $dst=$targetWb.Worksheets.Item("Scenarios")
    Copy-Formats $src "A1:AE9" $dst "A1:AE9"
    Copy-Dimensions $src $dst 9 31
    $last=$dst.UsedRange.Row+$dst.UsedRange.Rows.Count-1
    for ($r=10; $r -le $last; $r++) {
        $a=$dst.Cells.Item($r,1).Value2; $b=$dst.Cells.Item($r,2).Value2
        $c=[string]$dst.Cells.Item($r,3).Value2; $d=[string]$dst.Cells.Item($r,4).Value2
        $nonempty=0
        for ($cc=1; $cc -le 31; $cc++) { if ($null -ne $dst.Cells.Item($r,$cc).Value2 -and [string]$dst.Cells.Item($r,$cc).Value2 -ne "") {$nonempty++; break} }
        if ($nonempty -eq 0) { $sourceRow=19 }
        elseif ($c -eq "Break Down") { $sourceRow=53 }
        elseif ($c -eq "Catalyst Scenarios") { $sourceRow=121 }
        elseif (($b -is [double] -or $b -is [int]) -and $c) { $sourceRow=20 }
        elseif ($d -eq "[%]" -and $a -eq 4) { $sourceRow=11 }
        elseif ($d -eq "[%]") { $sourceRow=22 }
        elseif ($a -eq 4 -and [string]$b -eq "Absolute") { $sourceRow=10 }
        else { $sourceRow=21 }
        Copy-RowFormat $src $sourceRow $dst $r "AE"
    }
    for ($cc=1; $cc -le 31; $cc++) {
        $dst.Columns.Item($cc).ColumnWidth=$src.Columns.Item($cc).ColumnWidth
        $dst.Columns.Item($cc).Hidden=$src.Columns.Item($cc).Hidden
    }

    # Pipeline semantic rows.
    # Pipeline has its own approved authority.  The active MOLN workbook was
    # itself restyled incorrectly, so it must never be allowed to propagate
    # that drift back into MOLN/CMPX or future models.
    $src=$pipelineRefWb.Worksheets.Item("Pipeline"); $dst=$targetWb.Worksheets.Item("Pipeline")
    [void]$dst.Cells.FormatConditions.Delete()
    Copy-Formats $src "A1:AH8" $dst "A1:AH8"
    Copy-Dimensions $src $dst 8 34
    $last=$dst.UsedRange.Row+$dst.UsedRange.Rows.Count-1
    for ($r=9; $r -le $last; $r++) {
        $a=[string]$dst.Cells.Item($r,1).Value2
        $d=[string]$dst.Cells.Item($r,4).Value2
        $formula=[string]$dst.Cells.Item($r,4).Formula
        if (-not $d -and -not $formula) { $sourceRow=15 }
        elseif (($d+$formula) -like "*Market Share*") { $sourceRow=11 }
        elseif (($d+$formula) -like "* TAM*") { $sourceRow=10 }
        elseif (($d+$formula) -like "*List Price*") { $sourceRow=12 }
        elseif (($d+$formula) -like "*Revenue*") { $sourceRow=13 }
        elseif (($d+$formula) -like "*COGS*") { $sourceRow=14 }
        elseif ($a -eq "X") { $sourceRow=9 }
        else { $sourceRow=15 }
        Copy-RowFormat $src $sourceRow $dst $r "AH"
    }
    for ($cc=1; $cc -le 34; $cc++) {
        $dst.Columns.Item($cc).ColumnWidth=$src.Columns.Item($cc).ColumnWidth
        $dst.Columns.Item($cc).Hidden=$src.Columns.Item($cc).Hidden
    }

    # Peer View: company/peer rows retain BIC/T1/AVG colour semantics.
    $src=$refWb.Worksheets.Item("Peer View"); $dst=$targetWb.Worksheets.Item("Peer View")
    Copy-Formats $src "A1:Q7" $dst "A1:Q7"
    Copy-Dimensions $src $dst 7 17
    $companyRows=@{}; $peerRows=@{}
    $srcLast=$src.UsedRange.Row+$src.UsedRange.Rows.Count-1
    for ($r=8; $r -le $srcLast; $r++) {
        $key=Rating-Key ([string]$src.Cells.Item($r,7).Value2)
        $owner=[string]$src.Cells.Item($r,6).Value2
        if ($owner -match "MOLN") { if (-not $companyRows.ContainsKey($key)) {$companyRows[$key]=$r} }
        else { if (-not $peerRows.ContainsKey($key)) {$peerRows[$key]=$r} }
    }
    $last=$dst.UsedRange.Row+$dst.UsedRange.Rows.Count-1
    for ($r=8; $r -le $last; $r++) {
        $key=Rating-Key ([string]$dst.Cells.Item($r,7).Value2)
        $owner=[string]$dst.Cells.Item($r,6).Value2
        $isCompany=($owner -match " US Equity$")
        if ($isCompany -and $companyRows.ContainsKey($key)) {$sourceRow=$companyRows[$key]}
        elseif (-not $isCompany -and $peerRows.ContainsKey($key)) {$sourceRow=$peerRows[$key]}
        elseif ($isCompany) {$sourceRow=8} else {$sourceRow=9}
        Copy-RowFormat $src $sourceRow $dst $r "Q"
    }
    for ($cc=1; $cc -le 17; $cc++) {
        $dst.Columns.Item($cc).ColumnWidth=$src.Columns.Item($cc).ColumnWidth
        $dst.Columns.Item($cc).Hidden=$src.Columns.Item($cc).Hidden
    }

    # Catalyst v3 two-table structure, using MOLN's four original colour
    # families as reusable style primitives.
    $src=$refWb.Worksheets.Item("Catalyst"); $dst=$targetWb.Worksheets.Item("Catalyst")
    Copy-Formats $src "B1:X5" $dst "B1:X5"
    $tableTitleRow=$null
    $dstLast=$dst.UsedRange.Row+$dst.UsedRange.Rows.Count-1
    for ($r=6; $r -le $dstLast; $r++) {
        if ([string]$dst.Cells.Item($r,2).Value2 -eq "Catalyst Input Changes (Table 3)") {$tableTitleRow=$r; break}
    }
    if ($null -eq $tableTitleRow) { throw "Catalyst v3 Table 3 title not found" }
    $scenarioLast=$tableTitleRow-2
    $targetCount=0
    for ($g=7; $g -le 200; $g+=4) {
        if ([string]$dst.Cells.Item(7,$g+2).Value2 -eq "Market Price") {$targetCount++} else {break}
    }
    Copy-Formats $src "B6:F13" $dst "B6:F13"
    for ($r=14; $r -le $scenarioLast; $r++) { Copy-RowFormat $src 13 $dst $r "F" }
    for ($i=0; $i -lt $targetCount; $i++) {
        $g=7+4*$i; $srcG=7+4*($i%4)
        $srcStart=Col-Letter $srcG; $srcEnd=Col-Letter ($srcG+3)
        # G:V are single-letter columns in the MOLN style source.
        Copy-Formats $src ($srcStart+"6:"+$srcEnd+"13") $dst ((Col-Letter $g)+"6:"+(Col-Letter ($g+3))+"13")
        for ($r=14; $r -le $scenarioLast; $r++) {
            [void]$src.Range($srcStart+"13:"+$srcEnd+"13").Copy()
            [void]$dst.Range($dst.Cells.Item($r,$g),$dst.Cells.Item($r,$g+3)).PasteSpecial($xlPasteFormats)
            $dst.Rows.Item($r).RowHeight=$src.Rows.Item(13).RowHeight
        }
        for ($off=0; $off -lt 4; $off++) {
            $dst.Columns.Item($g+$off).ColumnWidth=$src.Columns.Item($srcG+$off).ColumnWidth
        }
    }
    $finalCol=7+4*$targetCount; $upsideCol=$finalCol+1
    for ($r=7; $r -le 13; $r++) {
        [void]$src.Range("W"+$r+":X"+$r).Copy()
        [void]$dst.Range($dst.Cells.Item($r,$finalCol),$dst.Cells.Item($r,$upsideCol)).PasteSpecial($xlPasteFormats)
    }
    for ($r=14; $r -le $scenarioLast; $r++) {
        [void]$src.Range("W13:X13").Copy()
        [void]$dst.Range($dst.Cells.Item($r,$finalCol),$dst.Cells.Item($r,$upsideCol)).PasteSpecial($xlPasteFormats)
    }

    $sourceT3=$null
    for ($col=1; $col -le $src.UsedRange.Columns.Count-3; $col++) {
        if ([string]$src.Cells.Item(8,$col+1).Value2 -eq "Conv." -and
                [string]$src.Cells.Item(8,$col+2).Value2 -like "Market Share*" -and
                [string]$src.Cells.Item(8,$col+3).Value2 -like "LOA*" -and
                $null -eq $src.Cells.Item(9,$col).Value2) {$sourceT3=$col; break}
    }
    if ($null -eq $sourceT3) { throw "MOLN Catalyst Table 3 style source not found" }
    $targetRow=$tableTitleRow+1; $headerRow=$tableTitleRow+2; $inputFirst=$tableTitleRow+3
    [void]$src.Range("B5").Copy(); [void]$dst.Cells.Item($tableTitleRow,2).PasteSpecial($xlPasteFormats)
    for ($j=0; $j -lt 4; $j++) {
        [void]$src.Cells.Item(10+$j,$sourceT3).Copy(); [void]$dst.Cells.Item($inputFirst+$j,2).PasteSpecial($xlPasteFormats)
    }
    for ($i=0; $i -lt $targetCount; $i++) {
        $g=7+4*$i; $ms=$g+2; $loa=$g+3; $conv=$g+4; $srcG=7+4*($i%4)
        for ($cc=$ms; $cc -le $conv; $cc++) {
            [void]$src.Cells.Item(7,$srcG).Copy(); [void]$dst.Cells.Item($targetRow,$cc).PasteSpecial($xlPasteFormats)
        }
        [void]$src.Cells.Item(8,$sourceT3+2).Copy(); [void]$dst.Cells.Item($headerRow,$ms).PasteSpecial($xlPasteFormats)
        [void]$src.Cells.Item(8,$sourceT3+3).Copy(); [void]$dst.Cells.Item($headerRow,$loa).PasteSpecial($xlPasteFormats)
        [void]$src.Cells.Item(8,$sourceT3+1).Copy(); [void]$dst.Cells.Item($headerRow,$conv).PasteSpecial($xlPasteFormats)
        for ($j=0; $j -lt 4; $j++) {
            [void]$src.Cells.Item(10+$j,$sourceT3+2).Copy(); [void]$dst.Cells.Item($inputFirst+$j,$ms).PasteSpecial($xlPasteFormats)
            [void]$src.Cells.Item(10+$j,$sourceT3+3).Copy(); [void]$dst.Cells.Item($inputFirst+$j,$loa).PasteSpecial($xlPasteFormats)
            [void]$src.Cells.Item(10+$j,$sourceT3+1).Copy(); [void]$dst.Cells.Item($inputFirst+$j,$conv).PasteSpecial($xlPasteFormats)
        }
    }

    $targetWb.Save()
    Write-Host "Applied approved reference styles; Pipeline source: $PipelineReferencePath"
}
finally {
    if ($bootstrapWb -ne $null) { $bootstrapWb.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($bootstrapWb) }
    if ($refWb -ne $null) { $refWb.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($refWb) }
    if ($pipelineRefWb -ne $null) { $pipelineRefWb.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pipelineRefWb) }
    if ($targetWb -ne $null) { $targetWb.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($targetWb) }
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
