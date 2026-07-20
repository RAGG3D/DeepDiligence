param(
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [string]$Ticker = ""
)

$ErrorActionPreference = "Stop"
$xlCalculationManual = -4135
$xlPasteFormats = -4122
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $null
$styleSheet = $null

function ColLetter([int]$n) {
    $s = ""
    while ($n -gt 0) {
        $n--
        $s = [char](65 + ($n % 26)) + $s
        $n = [math]::Floor($n / 26)
    }
    return $s
}

function NumValue($value, [double]$fallback) {
    if ($null -eq $value -or [string]$value -eq "") { return $fallback }
    try { return [double]$value } catch { return $fallback }
}

try {
    $wb = $excel.Workbooks.Open($Path, 0, $false)
    $excel.Calculation = $xlCalculationManual
    $c = $wb.Worksheets.Item("Catalyst")
    $s = $wb.Worksheets.Item("Scenarios")
    $v = $wb.Worksheets.Item("VALUATION")

    # One target per Absolute drug x indication, in the same order used by the
    # Breakdown waterfall.  The visible main table now scales to every target.
    $targetRows = @()
    $assetRows = @()
    $absFirst = $null
    $absLast = $null
    $currentAsset = $null
    $usedScenarioRows = $s.UsedRange.Rows.Count
    for ($r = 1; $r -le $usedScenarioRows; $r++) {
        if ($s.Cells.Item($r,1).Value2 -ne 4 -or [string]$s.Cells.Item($r,2).Value2 -ne "Absolute") { continue }
        if ($null -eq $absFirst) { $absFirst = $r }
        $absLast = $r
        $label = [string]$s.Cells.Item($r,3).Formula
        $unit = [string]$s.Cells.Item($r,4).Value2
        if ($unit -eq "[%]" -and $currentAsset) {
            $indication = "All"
            if ($label -match '&"\s*(.*?)\s+Market Share"') { $indication = $Matches[1] }
            elseif ($label -match '\)\s+(.*?)\s+Market Share$') { $indication = $Matches[1] }
            $name = $(if ($indication -eq "All") { $currentAsset } else { $currentAsset + " - " + $indication })
            $targetRows += @{ Name=$name; Short=$indication; MsRow=$r }
        } elseif ($unit -ne "[%]") {
            $raw = [string]$s.Cells.Item($r,3).Value2
            $currentAsset = ($raw -split ' \(')[0].Trim()
            $assetRows += $r
        }
    }
    if ($targetRows.Count -lt 1) { throw "No drug x indication targets found in Scenarios Absolute" }

    $outcomes = @("Increase", "Remain", "Decrease", "Suspension")
    $scenarioCount = 4 * $targetRows.Count
    $scenarioFirst = 10
    $scenarioLast = 9 + $scenarioCount
    $tableTitleRow = $scenarioLast + 2
    $tableTargetRow = $tableTitleRow + 1
    $tableHeaderRow = $tableTitleRow + 2
    $tableInputFirst = $tableTitleRow + 3
    $tableInputLast = $tableInputFirst + 3
    $finalCol = 7 + 4 * $targetRows.Count
    $upsideCol = $finalCol + 1
    $lastUsedCol = $upsideCol

    # Preserve analyst inputs across idempotent rebuilds.  Supports both the old
    # horizontal Table 3 and the new below-main layout.
    $inputMap = @{}
    $baseLoa = @{}
    $oldT3Start = $null
    $oldMaxCol = [math]::Max($c.UsedRange.Columns.Count, 140)
    for ($col = 1; $col -le $oldMaxCol - 3; $col++) {
        $target = [string]$c.Cells.Item(8,$col).Value2
        $h1 = [string]$c.Cells.Item(8,$col+1).Value2
        $h2 = [string]$c.Cells.Item(8,$col+2).Value2
        $h3 = [string]$c.Cells.Item(8,$col+3).Value2
        if ($target -and $h1 -eq "Conv." -and $h2 -like "Market Share Change*" -and $h3 -like "LOA Change*") {
            if ($null -eq $oldT3Start) { $oldT3Start = $col }
            $rows = @{}
            for ($j = 0; $j -lt 4; $j++) {
                $rr = 10 + $j
                $outcome = [string]$c.Cells.Item($rr,$col).Value2
                if (-not $outcome) { $outcome = $outcomes[$j] }
                $rows[$outcome] = @{
                    Conv=(NumValue $c.Cells.Item($rr,$col+1).Value2 $(if ($outcome -eq "Remain") {1} else {0}));
                    MS=(NumValue $c.Cells.Item($rr,$col+2).Value2 0);
                    LOA=(NumValue $c.Cells.Item($rr,$col+3).Value2 0)
                }
            }
            $inputMap[$target] = $rows
        }

        # Old Table 2 baseline LOA row.
        if ($target -and [string]$c.Cells.Item(8,$col+1).Value2 -eq "Conv." -and
                [string]$c.Cells.Item(8,$col+2).Value2 -eq "Market Share" -and
                [string]$c.Cells.Item(8,$col+3).Value2 -eq "LOA") {
            $baseLoa[$target] = NumValue $c.Cells.Item(9,$col+3).Value2 0.10
        }
    }

    $existingNewTitle = $null
    for ($r = 1; $r -le $c.UsedRange.Rows.Count; $r++) {
        if ([string]$c.Cells.Item($r,2).Value2 -eq "Catalyst Input Changes (Table 3)") {
            $existingNewTitle = $r
            break
        }
    }
    if ($null -ne $existingNewTitle) {
        $oldTargetRow = $existingNewTitle + 1
        $oldHeaderRow = $existingNewTitle + 2
        $oldInputFirst = $existingNewTitle + 3
        for ($i = 0; $i -lt $targetRows.Count; $i++) {
            $g = 7 + 4*$i
            $msCol = $g + 2; $loaCol = $g + 3; $convCol = $g + 4
            $target = [string]$c.Cells.Item($oldTargetRow,$msCol).Value2
            if (-not $target) { $target = $targetRows[$i].Name }
            $rows = @{}
            for ($j = 0; $j -lt 4; $j++) {
                $rr = $oldInputFirst + $j
                $outcome = [string]$c.Cells.Item($rr,2).Value2
                if (-not $outcome) { $outcome = $outcomes[$j] }
                $rows[$outcome] = @{
                    MS=(NumValue $c.Cells.Item($rr,$msCol).Value2 0);
                    LOA=(NumValue $c.Cells.Item($rr,$loaCol).Value2 0);
                    Conv=(NumValue $c.Cells.Item($rr,$convCol).Value2 $(if ($outcome -eq "Remain") {1} else {0}))
                }
            }
            $inputMap[$target] = $rows
        }
        for ($i = 0; $i -lt $targetRows.Count; $i++) {
            $g = 7 + 4*$i
            $target = $targetRows[$i].Name
            $baseLoa[$target] = NumValue $c.Cells.Item(9,$g+3).Value2 0.10
        }
    }

    # Default missing target mappings to neutral values.
    foreach ($targetRow in $targetRows) {
        $target = $targetRow.Name
        if (-not $inputMap.ContainsKey($target)) {
            $rows = @{}
            foreach ($outcome in $outcomes) {
                $rows[$outcome] = @{ MS=0.0; LOA=0.0; Conv=$(if ($outcome -eq "Remain") {1.0} else {0.0}) }
            }
            $inputMap[$target] = $rows
        }
        if (-not $baseLoa.ContainsKey($target)) { $baseLoa[$target] = 0.10 }
    }

    # Save neutral source formats in a temporary worksheet before clearing the
    # legacy main/Table 2/Table 3/support areas.
    $styleSheet = $wb.Worksheets.Add()
    $styleSheet.Name = "__CatalystStyleTemp"
    [void]$c.Range("B6:X13").Copy($styleSheet.Range("A1:W8"))
    if ($null -ne $oldT3Start) {
        [void]$c.Range($c.Cells.Item(8,$oldT3Start),$c.Cells.Item(13,$oldT3Start+3)).Copy($styleSheet.Range("A20:D25"))
    } else {
        [void]$c.Range("G7:J12").Copy($styleSheet.Range("A20:D25"))
    }

    # Remove all legacy tables/support content and merges.  Metadata rows 2:5
    # remain untouched.  Clear() removes the old Table 2 visual residue too.
    foreach ($merge in @($c.UsedRange.MergeCells)) { }
    $merged = @($c.UsedRange.MergeArea)
    try {
        for ($idx = $c.UsedRange.MergeCells.Count; $idx -ge 1; $idx--) {
            $area = $c.UsedRange.MergeCells.Item($idx)
            if ($area.Row -ge 6) { $area.UnMerge() }
        }
    } catch {
        # Excel's MergeCells collection is inconsistent; unmerge the full build
        # area as a safe fallback.
        [void]$c.Range("B6:EJ200").UnMerge()
    }
    [void]$c.Range("B6:EJ200").Clear()

    # Descriptor and row formats.
    [void]$styleSheet.Range("A1:E8").Copy()
    [void]$c.Range("B6:F13").PasteSpecial($xlPasteFormats)
    for ($r = 14; $r -le $scenarioLast; $r++) {
        [void]$styleSheet.Range("A8:E8").Copy()
        [void]$c.Range("B${r}:F${r}").PasteSpecial($xlPasteFormats)
    }

    # Main-table target groups.  Cycle the original four colour families while
    # keeping every drug x indication visible.
    for ($i = 0; $i -lt $targetRows.Count; $i++) {
        $g = 7 + 4*$i
        $sourceCol = 6 + 4*($i % 4)  # temp F/I... = original G/J...
        [void]$styleSheet.Range($styleSheet.Cells.Item(1,$sourceCol),$styleSheet.Cells.Item(8,$sourceCol+3)).Copy()
        [void]$c.Range($c.Cells.Item(6,$g),$c.Cells.Item(13,$g+3)).PasteSpecial($xlPasteFormats)
        for ($r = 14; $r -le $scenarioLast; $r++) {
            [void]$styleSheet.Range($styleSheet.Cells.Item(8,$sourceCol),$styleSheet.Cells.Item(8,$sourceCol+3)).Copy()
            [void]$c.Range($c.Cells.Item($r,$g),$c.Cells.Item($r,$g+3)).PasteSpecial($xlPasteFormats)
        }
        $c.Columns.Item($g).ColumnWidth = 13
        $c.Columns.Item($g+1).ColumnWidth = 8
        $c.Columns.Item($g+2).ColumnWidth = 13
        $c.Columns.Item($g+3).ColumnWidth = 8
    }

    # Final price/upside styles.
    [void]$styleSheet.Range("V2:W8").Copy()
    [void]$c.Range($c.Cells.Item(7,$finalCol),$c.Cells.Item(13,$upsideCol)).PasteSpecial($xlPasteFormats)
    for ($r = 14; $r -le $scenarioLast; $r++) {
        [void]$styleSheet.Range("V8:W8").Copy()
        [void]$c.Range($c.Cells.Item($r,$finalCol),$c.Cells.Item($r,$upsideCol)).PasteSpecial($xlPasteFormats)
    }

    # Main table headers.
    $c.Range("B7").Value2 = "Scenario"
    $c.Range("C7").Value2 = "Catalyst Target"
    $c.Range("D7").Value2 = "Outcome"
    $c.Range("E7").Value2 = "Scenario Summary"
    $c.Range("F7").Value2 = "Base Case (USD/Share)"
    $c.Range("F8").Value2 = "USD/Share"
    for ($i = 0; $i -lt $targetRows.Count; $i++) {
        $g = 7 + 4*$i
        $c.Cells.Item(7,$g).Value2 = $targetRows[$i].Name
        $c.Cells.Item(7,$g+2).Value2 = "Market Price"
        $c.Cells.Item(8,$g).Value2 = "USD/Share"
        $c.Cells.Item(8,$g+1).Value2 = "LOA"
        $c.Cells.Item(8,$g+2).Value2 = "USD/Share"
        $c.Cells.Item(8,$g+3).Value2 = "LOA"
    }
    $c.Cells.Item(7,$finalCol).Value2 = "Final Market Price"
    $c.Cells.Item(8,$finalCol).Value2 = "USD/Share"
    $c.Cells.Item(7,$upsideCol).Value2 = "Upside"
    $c.Cells.Item(8,$upsideCol).Value2 = "%"

    $cashPerShare = "MAX(0,INDEX(RCFS!`$G`$38:`$W`$38,1,MATCH(YEAR(`$B`$9),RCFS!`$G`$4:`$W`$4,0)))/VALUATION!`$C`$47"
    $oldCashPerShare = "MAX(0,VALUATION!`$C`$45/VALUATION!`$C`$47)"
    $c.Range("B9").Formula = "=TODAY()"
    $c.Range("C9").Value2 = "Base Case"
    $c.Range("D9").Value2 = "Base"
    $c.Range("E9").Value2 = "Scenario 1"
    $c.Range("F9").Formula = "=IFERROR(VALUATION!`$C`$48-"+$oldCashPerShare+"+"+$cashPerShare+",0)"

    $marketPriceCols = @()
    for ($i = 0; $i -lt $targetRows.Count; $i++) {
        $g = 7 + 4*$i
        $gL = ColLetter $g; $marketL = ColLetter ($g+2); $marketLoaL = ColLetter ($g+3)
        $c.Cells.Item(9,$g).Formula = "=IFERROR(VALUATION!`$G`$"+(26+$i)+",0)"
        $c.Cells.Item(9,$g+1).Value2 = 1.0
        $c.Cells.Item(9,$g+3).Value2 = [double]$baseLoa[$targetRows[$i].Name]
        $c.Cells.Item(9,$g+2).Formula = "=IFERROR("+$gL+"9*"+$marketLoaL+"9,0)"
        $c.Cells.Item(6,$g).Formula = "=IFERROR("+$gL+"9/`$F`$9,0)"
        $c.Cells.Item(6,$g).NumberFormat = "0.0%"
        $marketPriceCols += $marketL
    }
    $finalL = ColLetter $finalCol; $upsideL = ColLetter $upsideCol
    $c.Cells.Item(9,$finalCol).Formula = "=IFERROR(SUM("+(($marketPriceCols | ForEach-Object {$_+"9"}) -join ",")+"),0)"
    $c.Cells.Item(9,$upsideCol).Formula = "=IFERROR("+$finalL+"9/VALUATION!`$C`$52-1,0)"

    # Table 3 is the only input table.  Market Share aligns with the same
    # target's Market Price-USD/Share column; LOA aligns with Market Price-LOA;
    # Conv. is immediately to the right of LOA.
    $c.Cells.Item($tableTitleRow,2).Value2 = "Catalyst Input Changes (Table 3)"
    $c.Cells.Item($tableHeaderRow,2).Value2 = "Outcome"
    for ($j = 0; $j -lt 4; $j++) { $c.Cells.Item($tableInputFirst+$j,2).Value2 = $outcomes[$j] }
    for ($i = 0; $i -lt $targetRows.Count; $i++) {
        $g = 7 + 4*$i
        $msCol = $g+2; $loaCol = $g+3; $convCol = $g+4
        $sourceCol = 6 + 4*($i % 4)
        # Header fill follows the matching main-table target colour.
        for ($cc = $msCol; $cc -le $convCol; $cc++) {
            [void]$styleSheet.Cells.Item(2,$sourceCol).Copy()
            [void]$c.Cells.Item($tableTargetRow,$cc).PasteSpecial($xlPasteFormats)
        }
        $c.Cells.Item($tableTargetRow,$msCol).Value2 = $targetRows[$i].Name
        # Reordered legacy Table 3 styles: MS, LOA, then Conv.
        [void]$styleSheet.Cells.Item(20,3).Copy(); [void]$c.Cells.Item($tableHeaderRow,$msCol).PasteSpecial($xlPasteFormats)
        [void]$styleSheet.Cells.Item(20,4).Copy(); [void]$c.Cells.Item($tableHeaderRow,$loaCol).PasteSpecial($xlPasteFormats)
        [void]$styleSheet.Cells.Item(20,2).Copy(); [void]$c.Cells.Item($tableHeaderRow,$convCol).PasteSpecial($xlPasteFormats)
        $c.Cells.Item($tableHeaderRow,$msCol).Value2 = "Market Share Change"
        $c.Cells.Item($tableHeaderRow,$loaCol).Value2 = "LOA Change"
        $c.Cells.Item($tableHeaderRow,$convCol).Value2 = "Conv."
        for ($j = 0; $j -lt 4; $j++) {
            $rr = $tableInputFirst+$j; $outcome=$outcomes[$j]
            [void]$styleSheet.Cells.Item(22+$j,3).Copy(); [void]$c.Cells.Item($rr,$msCol).PasteSpecial($xlPasteFormats)
            [void]$styleSheet.Cells.Item(22+$j,4).Copy(); [void]$c.Cells.Item($rr,$loaCol).PasteSpecial($xlPasteFormats)
            [void]$styleSheet.Cells.Item(22+$j,2).Copy(); [void]$c.Cells.Item($rr,$convCol).PasteSpecial($xlPasteFormats)
            $c.Cells.Item($rr,$msCol).Value2 = [double]$inputMap[$targetRows[$i].Name][$outcome].MS
            $c.Cells.Item($rr,$loaCol).Value2 = [double]$inputMap[$targetRows[$i].Name][$outcome].LOA
            $c.Cells.Item($rr,$convCol).Value2 = [double]$inputMap[$targetRows[$i].Name][$outcome].Conv
            $c.Cells.Item($rr,$msCol).NumberFormat = "0.0%"
            $c.Cells.Item($rr,$loaCol).NumberFormat = "0.0%"
            $c.Cells.Item($rr,$convCol).NumberFormat = "0.0%"
        }
    }

    # Scenario IDs start immediately after the last Breakdown ID.
    $dividerRow = $null; $maxScenarioId = 4; $inBreakdown = $false
    for ($r=1; $r -le $usedScenarioRows; $r++) {
        $label = [string]$s.Cells.Item($r,3).Value2
        if ($label -eq "Break Down") { $inBreakdown=$true; continue }
        if ($label -eq "Catalyst Scenarios") { $dividerRow=$r; break }
        if ($inBreakdown) {
            $candidate = $s.Cells.Item($r,2).Value2
            if ($candidate -is [double] -or $candidate -is [int]) { $maxScenarioId=[math]::Max($maxScenarioId,[int]$candidate) }
        }
    }
    if ($null -eq $dividerRow) { throw "Catalyst Scenarios divider not found" }
    $scenarioStart = $maxScenarioId + 1

    # Main scenario formulas use the Scenario-1 breakdown proportions locked in
    # row 6.  LOA is base row 9 plus the matching Table-3 outcome change.
    for ($i = 0; $i -lt $targetRows.Count; $i++) {
        for ($j = 0; $j -lt 4; $j++) {
            $idx = 4*$i+$j; $row=$scenarioFirst+$idx; $sid=$scenarioStart+$idx
            $changedTarget = $targetRows[$i].Name; $outcome=$outcomes[$j]
            $c.Cells.Item($row,2).Value2 = [double]$sid
            $c.Cells.Item($row,3).Value2 = $changedTarget
            $c.Cells.Item($row,4).Value2 = $outcome
            $c.Cells.Item($row,5).Value2 = $targetRows[$i].Short+": "+$outcome
            $c.Cells.Item($row,6).Formula = "=IFERROR(VALUATION!P"+(5+$idx)+"-"+$oldCashPerShare+"+"+$cashPerShare+",0)"
            for ($k = 0; $k -lt $targetRows.Count; $k++) {
                $g=7+4*$k; $gL=ColLetter $g; $marketL=ColLetter ($g+2); $marketLoaL=ColLetter ($g+3)
                $msInputL=ColLetter ($g+2); $loaInputL=ColLetter ($g+3)
                $c.Cells.Item($row,$g).Formula = "=IFERROR(`$F"+$row+"*`$"+$gL+"`$6,0)"
                $c.Cells.Item($row,$g+1).Value2 = 1.0
                $criteria = "IF(`$C"+$row+"="+$gL+"`$7,`$D"+$row+",`"Remain`")"
                if ($k -eq $i -and $outcome -eq "Suspension") {
                    # A suspended target has zero post-catalyst LOA regardless
                    # of any stale/manual Table-3 delta retained from an older
                    # workbook.  The prior base+0 formula left Suspension at
                    # the base LOA and therefore valued it like Remain.
                    $c.Cells.Item($row,$g+3).Value2 = 0.0
                } else {
                    $c.Cells.Item($row,$g+3).Formula = "=MAX(0,"+$marketLoaL+"`$9+SUMIF(`$B`$"+$tableInputFirst+":`$B`$"+$tableInputLast+","+$criteria+",`$"+$loaInputL+"`$"+$tableInputFirst+":`$"+$loaInputL+"`$"+$tableInputLast+"))"
                }
                $c.Cells.Item($row,$g+2).Formula = "=IFERROR("+$gL+$row+"*"+$marketLoaL+$row+",0)"
            }
            $c.Cells.Item($row,$finalCol).Formula = "=IFERROR(SUM("+(($marketPriceCols | ForEach-Object {$_+$row}) -join ",")+"),0)"
            $c.Cells.Item($row,$upsideCol).Formula = "=IFERROR("+$finalL+$row+"/VALUATION!`$C`$52-1,0)"
        }
    }

    # Rebuild all Scenarios > Catalyst Scenarios blocks.  Only the changed
    # target receives Absolute MS + the aligned Table-3 MS change.
    $firstHeader = $dividerRow + 1
    $blockHeight = ($absLast - $absFirst + 2)
    # Legacy builds contain two headers: the visible "# + catalyst title" row
    # followed by a numeric-only row.  Delete the latter so the asset rows link
    # directly to the single visible header.
    $legacyHeaderRow = $firstHeader + 1
    $legacyNumber = $s.Cells.Item($legacyHeaderRow,2).Value2
    $legacyTitle = [string]$s.Cells.Item($legacyHeaderRow,3).Value2
    if (($legacyNumber -is [double] -or $legacyNumber -is [int]) -and -not $legacyTitle) {
        [void]$s.Rows.Item($legacyHeaderRow).Delete()
        $usedScenarioRows--
    }
    # The first rebuilt block occupies firstHeader..firstHeader+blockHeight-1.
    # Delete from the immediately following row; legacy workbooks may already
    # have the next title there.  A one-row spacer is recreated by the +1
    # destination stride below.
    $nextBlock = $firstHeader + $blockHeight
    if ($usedScenarioRows -ge $nextBlock) { [void]$s.Rows.Item($nextBlock.ToString()+":"+$usedScenarioRows).Delete() }
    for ($block = 1; $block -lt $scenarioCount; $block++) {
        $dest = $firstHeader + ($blockHeight+1)*$block
        [void]$s.Range("A"+$firstHeader+":AE"+($firstHeader+$blockHeight-1)).Copy($s.Range("A"+$dest+":AE"+($dest+$blockHeight-1)))
    }
    for ($i = 0; $i -lt $targetRows.Count; $i++) {
        for ($j = 0; $j -lt 4; $j++) {
            $idx=4*$i+$j; $h=$firstHeader+($blockHeight+1)*$idx; $sid=$scenarioStart+$idx
            $outcome=$outcomes[$j]
            $s.Range("B"+$h).Value2 = [double]$sid
            $s.Range("C"+$h).Value2 = $targetRows[$i].Name+" - "+$outcome
            # Deleting the legacy numeric-only header invalidates the copied
            # asset-row references.  Rebuild them explicitly: A = scenario ID,
            # B = catalyst title, matching the standard Scenario blocks.
            foreach ($assetRow in $assetRows) {
                $destAsset = $h + ($assetRow-$absFirst+1)
                $s.Range("A"+$destAsset).Formula = "=B"+$h
                $s.Range("B"+$destAsset).Formula = "=C"+$h
            }
            for ($k=0; $k -lt $targetRows.Count; $k++) {
                $msRow = $h + ($targetRows[$k].MsRow-$absFirst+1)
                if ($k -eq $i) {
                    $msInputL=ColLetter (7+4*$k+2)
                    if ($outcome -eq "Suspension") {
                        # Suspension is a hard zero, not a zero change to the
                        # current market-share assumption.
                        $s.Range("Y"+$msRow).Value2 = 0.0
                    } else {
                        $s.Range("Y"+$msRow).Formula = "=MAX(0,`$Y`$"+$targetRows[$k].MsRow+"+Catalyst!`$"+$msInputL+"`$"+($tableInputFirst+$j)+")"
                    }
                } else {
                    $s.Range("Y"+$msRow).Formula = "=`$Y`$"+$targetRows[$k].MsRow
                }
            }
        }
    }

    # The visible Absolute Value table and the Catalyst scenario table must both
    # return DCF price/share (C48).  The legacy L4 anchor used RIS!X30, which is
    # an operating-model value and produced the reported 1,804 "price" at 3%.
    [void]$v.Range("L4:M5").ClearContents()
    $v.Range("L4").Formula = "=IFERROR(VALUATION!`$C`$48,0)"
    $v.Range("M4").Formula = "=VALUATION!`$C`$5"
    $v.Range("L5").Value2 = 4.0
    [void]$v.Range("L4:M5").Table($v.Range("C5"),$v.Range("C3"))
    $v.Range("L4").NumberFormat = '"Abs. Value"'
    $v.Range("M5").NumberFormat = "0.00"

    # Same Excel What-If Table function for every Catalyst scenario ID.
    [void]$v.Range("O4:P200").ClearContents()
    $v.Range("O4").Formula = "=IFERROR(VALUATION!`$C`$48,0)"
    $v.Range("P4").Formula = "=VALUATION!`$C`$5"
    for ($i=0; $i -lt $scenarioCount; $i++) { $v.Cells.Item(5+$i,15).Value2 = [double]($scenarioStart+$i) }
    [void]$v.Range("O4:P"+(4+$scenarioCount)).Table($v.Range("C5"),$v.Range("C3"))
    $v.Range("P5:P"+(4+$scenarioCount)).NumberFormat = "0.00"

    $c.Range("B5").Value2 = "Full-pipeline two-table catalyst framework"
    $c.Range("B2").Value2 = "Upcoming Catalyst"
    $c.Range("B3").Value2 = "Expected Disclosure"
    $c.Range("B4").Value2 = "Source / Cutoff"
    $c.Range("B2:E4").Font.Name = $c.Range("B5").Font.Name
    [void]$c.Activate()
    [void]$c.Range("A1").Select()

    $styleSheet.Delete()
    $styleSheet = $null
    $c.Calculate()
    $s.Calculate()
    $v.Calculate()
    $wb.Save()
    Write-Host "Built two-table Catalyst for $($targetRows.Count) targets x 4 outcomes (IDs $scenarioStart-$($scenarioStart+$scenarioCount-1)); Table 3 rows $tableTargetRow-${tableInputLast}: $Path"
}
finally {
    if ($styleSheet -ne $null) {
        try { $styleSheet.Delete() } catch { }
    }
    if ($wb -ne $null) {
        try { $wb.Close($false) } catch { Write-Host "Workbook close warning: $($_.Exception.Message)" }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
    }
    try { $excel.Quit() } catch { }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
