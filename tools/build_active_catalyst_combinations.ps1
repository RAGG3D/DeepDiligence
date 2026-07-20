param(
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [string]$Ticker = "",
    [Parameter(Mandatory = $true)]
    [string]$RelevantTargetsJson,
    [double]$ConvictionThreshold = 0.10
)

$ErrorActionPreference = "Stop"
$xlCalculationManual = -4135
$xlCalculationAutomatic = -4105
$xlDone = 0
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

function HasTarget([object[]]$targets, [string]$name) {
    foreach ($target in $targets) {
        if ([string]$target.Name -eq $name) { return $true }
    }
    return $false
}

try {
    $parsedRelevantTargets = $RelevantTargetsJson | ConvertFrom-Json
    $relevantNames = @()
    foreach ($item in $parsedRelevantTargets) { $relevantNames += [string]$item }
    if ($relevantNames.Count -lt 1) { throw "At least one relevant Catalyst target is required" }

    $wb = $excel.Workbooks.Open($Path, 0, $false)
    $excel.Calculation = $xlCalculationManual
    $c = $wb.Worksheets.Item("Catalyst")
    $s = $wb.Worksheets.Item("Scenarios")
    $v = $wb.Worksheets.Item("VALUATION")
    $scenarioSelector = NumValue $v.Range("C3").Value2 1.0
    $terminalGrowth = NumValue $v.Range("C5").Value2 0.03

    # Read the canonical target universe and each target's Absolute market-share
    # row before touching either worksheet.  ValuationRow retains the original
    # waterfall order even though active targets are displayed first.
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
            $targetRows += @{
                Name=$name
                Short=$indication
                MsRow=$r
                OriginalIndex=$targetRows.Count
                ValuationRow=(26+$targetRows.Count)
            }
        } elseif ($unit -ne "[%]") {
            $raw = [string]$s.Cells.Item($r,3).Value2
            $currentAsset = ($raw -split ' \(')[0].Trim()
            $assetRows += $r
        }
    }
    if ($targetRows.Count -lt 1) { throw "No drug x indication targets found in Scenarios Absolute" }

    $activeTargets = @()
    foreach ($name in $relevantNames) {
        $match = @($targetRows | Where-Object { [string]$_.Name -eq [string]$name })
        if ($match.Count -ne 1) { throw "Relevant target not found exactly once in Scenarios Absolute: $name" }
        $activeTargets += $match[0]
    }
    $inactiveTargets = @($targetRows | Where-Object { -not (HasTarget $activeTargets ([string]$_.Name)) })
    $displayTargets = @($activeTargets) + @($inactiveTargets)

    $outcomes = @("Increase", "Remain", "Decrease", "Suspension")
    $inputMap = @{}
    $baseLoa = @{}

    # Preserve row-9 base LOA by target name rather than by physical column.
    $oldMaxCol = [math]::Max($c.UsedRange.Columns.Count, 140)
    for ($col = 1; $col -le $oldMaxCol - 3; $col++) {
        $target = [string]$c.Cells.Item(7,$col).Value2
        if ((HasTarget $targetRows $target) -and [string]$c.Cells.Item(8,$col+3).Value2 -eq "LOA") {
            $baseLoa[$target] = NumValue $c.Cells.Item(9,$col+3).Value2 0.10
        }
    }

    # Preserve Table-3 inputs by scanning the actual target/header row.  This
    # supports both framework-v3 alignment (MS at g+2) and the v4+ independent
    # four-column groups introduced here (MS at g).
    $oldTableTitle = $null
    for ($r = 1; $r -le $c.UsedRange.Rows.Count; $r++) {
        if ([string]$c.Cells.Item($r,2).Value2 -eq "Catalyst Input Changes (Table 3)") {
            $oldTableTitle = $r
            break
        }
    }
    $oldStyleInputCol = $null
    if ($null -ne $oldTableTitle) {
        $oldTargetRow = $oldTableTitle + 1
        $oldHeaderRow = $oldTableTitle + 2
        $oldInputFirst = $oldTableTitle + 3
        for ($col = 1; $col -le $oldMaxCol - 2; $col++) {
            $target = [string]$c.Cells.Item($oldTargetRow,$col).Value2
            if (-not (HasTarget $targetRows $target)) { continue }
            if ([string]$c.Cells.Item($oldHeaderRow,$col).Value2 -ne "Market Share Change" -or
                    [string]$c.Cells.Item($oldHeaderRow,$col+1).Value2 -ne "LOA Change" -or
                    [string]$c.Cells.Item($oldHeaderRow,$col+2).Value2 -ne "Conv.") { continue }
            if ($null -eq $oldStyleInputCol) { $oldStyleInputCol = $col }
            $rows = @{}
            for ($j = 0; $j -lt 4; $j++) {
                $rr = $oldInputFirst + $j
                $outcome = [string]$c.Cells.Item($rr,2).Value2
                if (-not $outcome) { $outcome = $outcomes[$j] }
                $rows[$outcome] = @{
                    MS=(NumValue $c.Cells.Item($rr,$col).Value2 0)
                    LOA=(NumValue $c.Cells.Item($rr,$col+1).Value2 0)
                    Conv=(NumValue $c.Cells.Item($rr,$col+2).Value2 $(if ($outcome -eq "Remain") {1} else {0}))
                }
            }
            $inputMap[$target] = $rows
        }
    }

    foreach ($targetRow in $targetRows) {
        $target = [string]$targetRow.Name
        if (-not $inputMap.ContainsKey($target)) {
            $rows = @{}
            foreach ($outcome in $outcomes) {
                $rows[$outcome] = @{ MS=0.0; LOA=0.0; Conv=$(if ($outcome -eq "Remain") {1.0} else {0.0}) }
            }
            $inputMap[$target] = $rows
        }
        if (-not $baseLoa.ContainsKey($target)) { $baseLoa[$target] = 0.10 }
    }

    # Conviction is an inclusion filter, not a weighting factor.  Every active
    # target must retain at least one >= threshold outcome; the scenario set is
    # the Cartesian product of the surviving outcome lists.
    $allowedByTarget = @{}
    foreach ($targetRow in $activeTargets) {
        $target = [string]$targetRow.Name
        $allowed = @()
        foreach ($outcome in $outcomes) {
            if ([double]$inputMap[$target][$outcome].Conv -ge $ConvictionThreshold) { $allowed += $outcome }
        }
        if ($allowed.Count -lt 1) { throw "No outcome has conviction >= $ConvictionThreshold for $target" }
        $allowedByTarget[$target] = $allowed
    }
    $combinations = New-Object System.Collections.ArrayList
    function Add-Combinations([int]$index, [object[]]$prefix) {
        if ($index -ge $activeTargets.Count) {
            [void]$combinations.Add(@($prefix))
            return
        }
        $target = [string]$activeTargets[$index].Name
        foreach ($outcome in @($allowedByTarget[$target])) {
            Add-Combinations ($index + 1) (@($prefix) + @($outcome))
        }
    }
    Add-Combinations 0 @()
    # Rank the complete surviving Cartesian set by RJConv.  RJConv. is the
    # raw product of the selected per-target convictions; a
    # deterministic outcome key resolves exact ties.
    $rankedCombinations = @()
    foreach ($rawCombo in @($combinations)) {
        $combo = @($rawCombo)
        [double]$probability = 1.0
        for ($a = 0; $a -lt $activeTargets.Count; $a++) {
            $target = [string]$activeTargets[$a].Name
            $probability *= [double]$inputMap[$target][[string]$combo[$a]].Conv
        }
        $rankedCombinations += [pscustomobject]@{
            Combo=$combo
            Probability=$probability
            Key=($combo -join "|")
        }
    }
    $rankedCombinations = @($rankedCombinations | Sort-Object `
        @{Expression={$_.Probability};Descending=$true}, `
        @{Expression={$_.Key};Descending=$false})
    $combinations = New-Object System.Collections.ArrayList
    foreach ($item in $rankedCombinations) { [void]$combinations.Add(@($item.Combo)) }
    $scenarioCount = $combinations.Count
    if ($scenarioCount -lt 1) { throw "No active Catalyst combinations were generated" }
    if ($scenarioCount -gt 512) { throw "Catalyst scenario count $scenarioCount exceeds safety cap 512" }

    $activeCount = $activeTargets.Count
    $scenarioFirst = 10
    $scenarioLast = 9 + $scenarioCount
    $tableTitleRow = $scenarioLast + 2
    $tableTargetRow = $tableTitleRow + 1
    $tableHeaderRow = $tableTitleRow + 2
    $tableInputFirst = $tableTitleRow + 3
    $tableInputLast = $tableInputFirst + 3
    # The Catalyst price table is embedded in B:C.  Outcomes follow the three
    # visible valuation outputs so Catalyst and clinical-only Test tabs remain
    # cleanly separated:
    # Scenario | Base Case | Final Market | Upside | RJConv. | outcomes | targets.
    $baseCol = 3
    $finalCol = 4
    $upsideCol = 5
    $rjconvCol = 6
    $outcomeFirstCol = 7
    $groupFirstCol = $outcomeFirstCol + $activeCount
    $lastUsedCol = $groupFirstCol + 4*$displayTargets.Count - 1

    # Save source formats and the first existing Scenarios catalyst block before
    # clearing/rebuilding.  Values/formulas are rebuilt explicitly below.
    $styleSheet = $wb.Worksheets.Add()
    $styleSheet.Name = "__CatalystStyleTemp"
    [void]$c.Range("B6:X13").Copy($styleSheet.Range("A1:W8"))
    if ($null -ne $oldTableTitle -and $null -ne $oldStyleInputCol) {
        [void]$c.Range(
            $c.Cells.Item($oldTargetRow,$oldStyleInputCol),
            $c.Cells.Item($oldInputFirst+3,$oldStyleInputCol+2)
        ).Copy($styleSheet.Range("A20:C25"))
    } else {
        [void]$c.Range("I68:K73").Copy($styleSheet.Range("A20:C25"))
    }
    $oldFinalCol = $null
    for ($col = 1; $col -le $oldMaxCol - 1; $col++) {
        if ([string]$c.Cells.Item(7,$col).Value2 -eq "Final Market Price") { $oldFinalCol = $col; break }
    }
    if ($null -ne $oldFinalCol) {
        [void]$c.Range($c.Cells.Item(7,$oldFinalCol),$c.Cells.Item(13,$oldFinalCol+1)).Copy($styleSheet.Range("A40:B46"))
    } else {
        [void]$c.Range("W7:X13").Copy($styleSheet.Range("A40:B46"))
    }

    # Clear the full legacy Catalyst build area and unhide/reset its columns.
    $clearLastRow = [math]::Max([int]$c.UsedRange.Rows.Count, [math]::Max(200,$tableInputLast))
    try { [void]$c.Range("B6:EJ"+$clearLastRow).UnMerge() } catch { }
    # Clear an existing embedded What-If array as a whole before rebuilding.
    # Excel rejects edits to only part of a Data Table.
    try {
        if ($c.Range("C9").HasArray) { [void]$c.Range("C9").CurrentArray.Clear() }
    } catch { }
    [void]$c.Range("B6:EJ"+$clearLastRow).Clear()
    for ($col = 2; $col -le [math]::Max($lastUsedCol,140); $col++) {
        $c.Columns.Item($col).Hidden = $false
    }

    # Descriptor/main-output formats.  The three deleted legacy descriptor
    # columns are now one active-target outcome column per relevant target.
    for ($r = 6; $r -le $scenarioLast; $r++) {
        $srcRow = $(if ($r -le 13) { $r - 5 } else { 8 })
        [void]$styleSheet.Cells.Item($srcRow,1).Copy(); [void]$c.Cells.Item($r,2).PasteSpecial($xlPasteFormats)
        [void]$styleSheet.Cells.Item($srcRow,5).Copy(); [void]$c.Cells.Item($r,$baseCol).PasteSpecial($xlPasteFormats)
        $finalSrcRow = $(if ($r -le 13) { 40 + [math]::Max(0,$r-7) } else { 46 })
        [void]$styleSheet.Cells.Item($finalSrcRow,1).Copy(); [void]$c.Cells.Item($r,$finalCol).PasteSpecial($xlPasteFormats)
        [void]$styleSheet.Cells.Item($finalSrcRow,2).Copy(); [void]$c.Cells.Item($r,$upsideCol).PasteSpecial($xlPasteFormats)
        [void]$styleSheet.Cells.Item($finalSrcRow,2).Copy(); [void]$c.Cells.Item($r,$rjconvCol).PasteSpecial($xlPasteFormats)
        for ($i = 0; $i -lt $activeCount; $i++) {
            [void]$styleSheet.Cells.Item($srcRow,2).Copy(); [void]$c.Cells.Item($r,$outcomeFirstCol+$i).PasteSpecial($xlPasteFormats)
        }
    }
    $c.Columns.Item(2).ColumnWidth = 12
    $c.Columns.Item($baseCol).ColumnWidth = 23
    $c.Columns.Item($finalCol).ColumnWidth = 15
    $c.Columns.Item($upsideCol).ColumnWidth = 11
    $c.Columns.Item($rjconvCol).ColumnWidth = 11
    for ($i = 0; $i -lt $activeCount; $i++) { $c.Columns.Item($outcomeFirstCol+$i).ColumnWidth = 20 }

    # Target block formats and widths.  OriginalIndex keeps the established
    # four-colour cycle stable after active targets move to the front.
    for ($i = 0; $i -lt $displayTargets.Count; $i++) {
        $g = $groupFirstCol + 4*$i
        $sourceCol = 6 + 4*([int]$displayTargets[$i].OriginalIndex % 4)
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

    # Normalize the two semantic LOA columns after the legacy colour-cycle
    # formats have been copied. Rebuilding an already-active v4+ sheet can feed
    # a descriptor/final-output format back into the last active target's
    # Market Price-LOA column (for example CMPX T). Every active final LOA must
    # use the same input style as the first active final LOA. For inactive
    # groups the source LOA column must use the group's correctly grey-masked
    # final-LOA style (V like X, Z like AB, and so on). Values and formulas are
    # deliberately untouched.
    $activeLoaStyleCol = $groupFirstCol + 3
    for ($i = 0; $i -lt $displayTargets.Count; $i++) {
        $g = $groupFirstCol + 4*$i
        $isActive = HasTarget $activeTargets ([string]$displayTargets[$i].Name)
        foreach ($r in @(6) + @(8..$scenarioLast)) {
            if ($isActive) {
                [void]$c.Cells.Item($r,$activeLoaStyleCol).Copy()
                [void]$c.Cells.Item($r,$g+3).PasteSpecial($xlPasteFormats)
            } else {
                [void]$c.Cells.Item($r,$g+3).Copy()
                [void]$c.Cells.Item($r,$g+1).PasteSpecial($xlPasteFormats)
            }
        }
    }

    # Main headers and row-9 base case.
    $c.Cells.Item(7,2).Value2 = "Scenario"
    $c.Cells.Item(8,2).Value2 = "ID"
    $c.Cells.Item(7,$baseCol).Value2 = "Base Case (USD/Share)"
    $c.Cells.Item(8,$baseCol).Value2 = [double]$terminalGrowth
    $c.Cells.Item(8,$baseCol).NumberFormat = "0.0%"
    $c.Cells.Item(7,$finalCol).Value2 = "Final Market Price"
    $c.Cells.Item(8,$finalCol).Value2 = "USD/Share"
    $c.Cells.Item(7,$upsideCol).Value2 = "Upside"
    $c.Cells.Item(8,$upsideCol).Value2 = "%"
    $c.Cells.Item(7,$rjconvCol).Value2 = "RJConv."
    $c.Cells.Item(8,$rjconvCol).Value2 = "%"
    for ($i = 0; $i -lt $activeCount; $i++) {
        $c.Cells.Item(7,$outcomeFirstCol+$i).Value2 = [string]$activeTargets[$i].Name
        $c.Cells.Item(8,$outcomeFirstCol+$i).Value2 = "Outcome"
    }
    for ($i = 0; $i -lt $displayTargets.Count; $i++) {
        $g = $groupFirstCol + 4*$i
        $c.Cells.Item(7,$g).Value2 = [string]$displayTargets[$i].Name
        $c.Cells.Item(7,$g+2).Value2 = "Market Price"
        $c.Cells.Item(8,$g).Value2 = "USD/Share"
        $c.Cells.Item(8,$g+1).Value2 = "LOA"
        $c.Cells.Item(8,$g+2).Value2 = "USD/Share"
        $c.Cells.Item(8,$g+3).Value2 = "LOA"
        [void]$c.Range($c.Cells.Item(7,$g),$c.Cells.Item(7,$g+1)).Merge()
        [void]$c.Range($c.Cells.Item(7,$g+2),$c.Cells.Item(7,$g+3)).Merge()
    }

    $cashPerShare = "MAX(0,INDEX(RCFS!`$G`$38:`$W`$38,1,MATCH(YEAR(VALUATION!`$C`$4),RCFS!`$G`$4:`$W`$4,0)))/VALUATION!`$C`$47"
    $oldCashPerShare = "MAX(0,VALUATION!`$C`$45/VALUATION!`$C`$47)"
    # B8 is the Data Table corner formula but displays as the ID label.  B9 is
    # scenario 1 but displays as Base.  C9 and subsequent rows are native Excel
    # What-If results at the fixed 3% terminal-growth input in C8.
    $c.Cells.Item(8,2).Formula = "=IFERROR(VALUATION!`$C`$48-"+$oldCashPerShare+"+"+$cashPerShare+",0)"
    $c.Cells.Item(8,2).NumberFormat = '"ID"'
    $c.Cells.Item(9,2).Value2 = 1.0
    $c.Cells.Item(9,2).NumberFormat = '"Base"'
    for ($i = 0; $i -lt $activeCount; $i++) { $c.Cells.Item(9,$outcomeFirstCol+$i).Value2 = "Base" }
    # Native Excel Data Tables require row/column input cells on the same sheet.
    # B6/C6 are invisible local bridges; VALUATION C3/C5 remain the canonical
    # model inputs and point to them. C6 follows the visible 3% cell in C8.
    $c.Cells.Item(6,2).Value2 = [double]$scenarioSelector
    $c.Cells.Item(6,3).Formula = "=`$C`$8"
    $c.Range("B6:C6").NumberFormat = ";;;"
    $v.Range("C3").Formula = "=Catalyst!`$B`$6"
    $v.Range("C5").Formula = "=Catalyst!`$C`$6"
    $baseL = ColLetter $baseCol
    $finalL = ColLetter $finalCol
    $upsideL = ColLetter $upsideCol

    $groupByTarget = @{}
    $marketPriceCols = @()
    for ($i = 0; $i -lt $displayTargets.Count; $i++) {
        $targetRow = $displayTargets[$i]
        $target = [string]$targetRow.Name
        $g = $groupFirstCol + 4*$i
        $groupByTarget[$target] = $g
        $gL = ColLetter $g
        $marketL = ColLetter ($g+2)
        $marketLoaL = ColLetter ($g+3)
        $c.Cells.Item(9,$g).Formula = "=IFERROR(VALUATION!`$G`$"+[int]$targetRow.ValuationRow+",0)"
        $c.Cells.Item(9,$g+1).Value2 = 1.0
        $c.Cells.Item(9,$g+3).Value2 = [double]$baseLoa[$target]
        $c.Cells.Item(9,$g+2).Formula = "=IFERROR(`$"+$gL+"9*`$"+$marketLoaL+"9,0)"
        $c.Cells.Item(6,$g).Formula = "=IFERROR(`$"+$gL+"`$9/`$"+$baseL+"`$9,0)"
        $c.Cells.Item(6,$g).NumberFormat = "0.0%"
        $marketPriceCols += $marketL
    }
    $c.Cells.Item(9,$finalCol).Formula = "=IFERROR(SUM("+(($marketPriceCols | ForEach-Object {"`$"+$_+"9"}) -join ",")+"),0)"
    $c.Cells.Item(9,$upsideCol).Formula = "=IFERROR(`$"+$finalL+"9/VALUATION!`$C`$52-1,0)"
    $c.Cells.Item(9,$rjconvCol).Value2 = 1.0
    $c.Cells.Item(9,$rjconvCol).NumberFormat = "0.0%"

    # Table 3 uses one independent four-column group per target.  The first
    # three columns hold MS / LOA / Conv. and the fourth is a spacer.  Every
    # group remains visible; the lifecycle layer applies a grey visual mask to
    # non-catalyst groups without removing their data or formulas.
    $c.Cells.Item($tableTitleRow,2).Value2 = "Catalyst Input Changes (Table 3)"
    $c.Cells.Item($tableHeaderRow,2).Value2 = "Outcome"
    for ($j = 0; $j -lt 4; $j++) { $c.Cells.Item($tableInputFirst+$j,2).Value2 = $outcomes[$j] }
    for ($i = 0; $i -lt $displayTargets.Count; $i++) {
        $target = [string]$displayTargets[$i].Name
        $g = $groupFirstCol + 4*$i
        for ($cc = 0; $cc -lt 3; $cc++) {
            [void]$styleSheet.Cells.Item(20,1+$cc).Copy(); [void]$c.Cells.Item($tableTargetRow,$g+$cc).PasteSpecial($xlPasteFormats)
            [void]$styleSheet.Cells.Item(21,1+$cc).Copy(); [void]$c.Cells.Item($tableHeaderRow,$g+$cc).PasteSpecial($xlPasteFormats)
        }
        $c.Cells.Item($tableTargetRow,$g).Value2 = $target
        $c.Cells.Item($tableHeaderRow,$g).Value2 = "Market Share Change"
        $c.Cells.Item($tableHeaderRow,$g+1).Value2 = "LOA Change"
        $c.Cells.Item($tableHeaderRow,$g+2).Value2 = "Conv."
        for ($j = 0; $j -lt 4; $j++) {
            $rr = $tableInputFirst+$j
            $outcome = $outcomes[$j]
            for ($cc = 0; $cc -lt 3; $cc++) {
                [void]$styleSheet.Cells.Item(22+$j,1+$cc).Copy(); [void]$c.Cells.Item($rr,$g+$cc).PasteSpecial($xlPasteFormats)
            }
            $c.Cells.Item($rr,$g).Value2 = [double]$inputMap[$target][$outcome].MS
            $c.Cells.Item($rr,$g+1).Value2 = [double]$inputMap[$target][$outcome].LOA
            $c.Cells.Item($rr,$g+2).Value2 = [double]$inputMap[$target][$outcome].Conv
            $c.Cells.Item($rr,$g).NumberFormat = "0.0%"
            $c.Cells.Item($rr,$g+1).NumberFormat = "0.0%"
            $c.Cells.Item($rr,$g+2).NumberFormat = "0.0%"
        }
    }

    # Scenario IDs continue immediately after the Breakdown module.
    $dividerRow = $null
    $maxScenarioId = 4
    $inBreakdown = $false
    for ($r = 1; $r -le $usedScenarioRows; $r++) {
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

    # Main table: one row per Cartesian-product combination.  Formula columns
    # are fully anchored so a future physical column move cannot alter logic.
    for ($idx = 0; $idx -lt $scenarioCount; $idx++) {
        $row = $scenarioFirst + $idx
        $sid = $scenarioStart + $idx
        $combo = @($combinations[$idx])
        $c.Cells.Item($row,2).Value2 = [double]$sid
        for ($a = 0; $a -lt $activeCount; $a++) { $c.Cells.Item($row,$outcomeFirstCol+$a).Value2 = [string]$combo[$a] }
        for ($i = 0; $i -lt $displayTargets.Count; $i++) {
            $target = [string]$displayTargets[$i].Name
            $g = $groupFirstCol + 4*$i
            $gL = ColLetter $g
            $marketL = ColLetter ($g+2)
            $marketLoaL = ColLetter ($g+3)
            $c.Cells.Item($row,$g).Formula = "=IFERROR(`$"+$baseL+$row+"*`$"+$gL+"`$6,0)"
            $c.Cells.Item($row,$g+1).Value2 = 1.0
            $activeIndex = -1
            for ($a = 0; $a -lt $activeCount; $a++) {
                if ([string]$activeTargets[$a].Name -eq $target) { $activeIndex = $a; break }
            }
            if ($activeIndex -ge 0) {
                $outcomeL = ColLetter ($outcomeFirstCol+$activeIndex)
                $tableLoaL = ColLetter ($g+1)
                $outcomeRef = "`$"+$outcomeL+$row
                $c.Cells.Item($row,$g+3).Formula = "=IF("+$outcomeRef+"=`"Suspension`",0,MAX(0,`$"+$marketLoaL+"`$9+IFERROR(INDEX(`$"+$tableLoaL+"`$"+$tableInputFirst+":`$"+$tableLoaL+"`$"+$tableInputLast+",MATCH("+$outcomeRef+",`$B`$"+$tableInputFirst+":`$B`$"+$tableInputLast+",0)),0)))"
            } else {
                $c.Cells.Item($row,$g+3).Formula = "=`$"+$marketLoaL+"`$9"
            }
            $c.Cells.Item($row,$g+2).Formula = "=IFERROR(`$"+$gL+$row+"*`$"+$marketLoaL+$row+",0)"
        }
        $c.Cells.Item($row,$finalCol).Formula = "=IFERROR(SUM("+(($marketPriceCols | ForEach-Object {"`$"+$_+$row}) -join ",")+"),0)"
        $c.Cells.Item($row,$upsideCol).Formula = "=IFERROR(`$"+$finalL+$row+"/VALUATION!`$C`$52-1,0)"
        $convTerms = @()
        for ($a = 0; $a -lt $activeCount; $a++) {
            $target = [string]$activeTargets[$a].Name
            $outcomeL = ColLetter ($outcomeFirstCol+$a)
            $targetGroup = [int]$groupByTarget[$target]
            $tableConvL = ColLetter ($targetGroup+2)
            $convTerms += "IFERROR(INDEX(`$"+$tableConvL+"`$"+$tableInputFirst+":`$"+$tableConvL+"`$"+$tableInputLast+",MATCH(`$"+$outcomeL+$row+",`$B`$"+$tableInputFirst+":`$B`$"+$tableInputLast+",0)),0)"
        }
        $c.Cells.Item($row,$rjconvCol).Formula = "=IFERROR(PRODUCT("+($convTerms -join ",")+"),0)"
        $c.Cells.Item($row,$rjconvCol).NumberFormat = "0.0%"
    }

    # Every target must stay visible because the complete drug breakdown drives
    # Final Market Price.  The lifecycle layer masks non-catalyst groups with a
    # grey background and grey font after this structural rebuild.
    for ($i = 0; $i -lt $displayTargets.Count; $i++) {
        $g = $groupFirstCol + 4*$i
        for ($cc = 0; $cc -lt 4; $cc++) { $c.Columns.Item($g+$cc).Hidden = $false }
    }

    # Rebuild Scenarios from the first normalized catalyst block.  Asset title
    # rows deliberately have a blank Y cell (fixes Y373/Y376-type residue).
    # Older workbooks may keep one or more blank rows between the divider and
    # the first scenario header.  Normalize that gap before measuring/copying
    # the block; otherwise the blank row is mistaken for the header and the
    # final indication row is truncated (and the old scenario ID survives as a
    # spurious second header).
    $firstHeader = $null
    for ($r = $dividerRow + 1; $r -le $usedScenarioRows; $r++) {
        if ([string]$s.Cells.Item($r,3).Value2 -like "Test Scenarios - *") { break }
        $candidateId = $s.Cells.Item($r,2).Value2
        if (($candidateId -is [double] -or $candidateId -is [int]) -and
                [string]$s.Cells.Item($r,3).Value2 -ne "") {
            $firstHeader = $r
            break
        }
    }
    if ($null -eq $firstHeader) { throw "Catalyst scenario template header not found" }
    if ($firstHeader -gt $dividerRow + 1) {
        $gapFirst = $dividerRow + 1
        $gapLast = $firstHeader - 1
        [void]$s.Rows.Item($gapFirst.ToString()+":"+$gapLast).Delete()
        $usedScenarioRows -= ($gapLast - $gapFirst + 1)
        $firstHeader = $dividerRow + 1
    }
    $blockHeight = ($absLast - $absFirst + 2)
    $firstBlockLast = $firstHeader + $blockHeight - 1
    if ($usedScenarioRows -lt $firstBlockLast) { throw "Existing Catalyst scenario template block is incomplete" }
    # Preserve historical Test Scenarios modules.  Collapse only the current
    # Catalyst section to its first template block, then insert the exact space
    # needed for the rebuilt actual-Catalyst combinations. Excel shifts every
    # Test formula and style together with its rows.
    $testDividerRow = $null
    for ($r = $firstBlockLast + 1; $r -le $usedScenarioRows; $r++) {
        if ([string]$s.Cells.Item($r,3).Value2 -like "Test Scenarios - *") {
            $testDividerRow = $r
            break
        }
    }
    $deleteFirst = $firstBlockLast + 1
    $deleteLast = $(if ($null -ne $testDividerRow) { $testDividerRow - 2 } else { $usedScenarioRows })
    if ($deleteLast -ge $deleteFirst) {
        [void]$s.Rows.Item($deleteFirst.ToString()+":"+$deleteLast).Delete()
    }
    if ($null -ne $testDividerRow -and $scenarioCount -gt 1) {
        $insertAt = $firstBlockLast + 2
        $insertCount = ($scenarioCount - 1) * ($blockHeight + 1)
        [void]$s.Rows.Item($insertAt.ToString()+":"+($insertAt+$insertCount-1)).Insert()
    }

    $assetRowSet = @{}
    foreach ($assetRow in $assetRows) { $assetRowSet[[int]$assetRow] = $true }
    for ($idx = 0; $idx -lt $scenarioCount; $idx++) {
        $h = $firstHeader + ($blockHeight+1)*$idx
        if ($idx -gt 0) {
            [void]$s.Range("A"+$firstHeader+":AE"+$firstBlockLast).Copy($s.Range("A"+$h+":AE"+($h+$blockHeight-1)))
        }
        $sid = $scenarioStart + $idx
        $combo = @($combinations[$idx])
        $parts = @()
        for ($a = 0; $a -lt $activeCount; $a++) { $parts += ([string]$activeTargets[$a].Name+": "+[string]$combo[$a]) }
        $s.Cells.Item($h,2).Value2 = [double]$sid
        $s.Cells.Item($h,3).Value2 = ($parts -join " | ")
        for ($sourceRow = $absFirst; $sourceRow -le $absLast; $sourceRow++) {
            $destRow = $h + ($sourceRow-$absFirst+1)
            if ($assetRowSet.ContainsKey([int]$sourceRow)) {
                $s.Cells.Item($destRow,25).ClearContents()
            }
        }
        for ($i = 0; $i -lt $targetRows.Count; $i++) {
            $target = [string]$targetRows[$i].Name
            $msRow = $h + ([int]$targetRows[$i].MsRow-$absFirst+1)
            $activeIndex = -1
            for ($a = 0; $a -lt $activeCount; $a++) {
                if ([string]$activeTargets[$a].Name -eq $target) { $activeIndex = $a; break }
            }
            if ($activeIndex -ge 0) {
                $outcomeL = ColLetter ($outcomeFirstCol+$activeIndex)
                $tableMsL = ColLetter ([int]$groupByTarget[$target])
                $mainRow = $scenarioFirst + $idx
                $outcomeRef = "Catalyst!`$"+$outcomeL+"`$"+$mainRow
                $baseMsRow = [int]$targetRows[$i].MsRow
                $s.Cells.Item($msRow,25).Formula = "=IF("+$outcomeRef+"=`"Suspension`",0,MAX(0,`$Y`$"+$baseMsRow+"+IFERROR(INDEX(Catalyst!`$"+$tableMsL+"`$"+$tableInputFirst+":`$"+$tableMsL+"`$"+$tableInputLast+",MATCH("+$outcomeRef+",Catalyst!`$B`$"+$tableInputFirst+":`$B`$"+$tableInputLast+",0)),0)))"
            } else {
                $s.Cells.Item($msRow,25).Formula = "=`$Y`$"+[int]$targetRows[$i].MsRow
            }
        }
    }

    # Renumber every preserved Test module after the rebuilt actual scenarios,
    # and mirror those IDs back to its Test-EVENT tab. This prevents collisions
    # when the actual Cartesian scenario count changes on refresh.
    $nextTestId = $scenarioStart + $scenarioCount
    $scenarioUsedAfterBuild = $s.UsedRange.Rows.Count
    for ($r = $firstHeader; $r -le $scenarioUsedAfterBuild; $r++) {
        $label = [string]$s.Cells.Item($r,3).Value2
        if ($label -notlike "Test Scenarios - *") { continue }
        $eventName = $label.Substring("Test Scenarios - ".Length)
        $testWs = $null
        foreach ($candidateWs in @($wb.Worksheets)) {
            if ([string]$candidateWs.Name -like "Test-*" -and
                    ([string]$candidateWs.Range("C2").Value2 -eq $eventName -or
                     [string]$candidateWs.Name -eq ("Test-"+$eventName))) {
                $testWs = $candidateWs
                break
            }
        }
        $testIndex = 0
        for ($rr = $r + 1; $rr -le $scenarioUsedAfterBuild; $rr++) {
            if ([string]$s.Cells.Item($rr,3).Value2 -like "Test Scenarios - *") { break }
            $candidateId = $s.Cells.Item($rr,2).Value2
            if (($candidateId -is [double] -or $candidateId -is [int]) -and
                    [string]$s.Cells.Item($rr,3).Value2 -ne "") {
                $s.Cells.Item($rr,2).Value2 = [double]$nextTestId
                if ($null -ne $testWs) { $testWs.Cells.Item(10+$testIndex,2).Value2 = [double]$nextTestId }
                $nextTestId++
                $testIndex++
            }
        }
    }

    # Remove the legacy VALUATION O:P Catalyst table and create the new native
    # two-variable What-If table directly in Catalyst B8:C(last scenario).
    try {
        if ($v.Range("P5").HasArray) { [void]$v.Range("P5").CurrentArray.Clear() }
    } catch { }
    [void]$v.Range("O4:P200").Clear()
    [void]$c.Range("B8:C"+$scenarioLast).Table($c.Range("C6"),$c.Range("B6"))
    $c.Range("C9:C"+$scenarioLast).NumberFormat = "0.00"

    $c.Range("B5").Value2 = "Active-catalyst Cartesian scenario framework"
    $c.Range("B2").Value2 = "Upcoming Catalyst"
    $c.Range("B3").Value2 = "Expected Disclosure"
    $c.Range("B4").Value2 = "Source / Cutoff"
    $c.Range("B2:E4").Font.Name = $c.Range("B5").Font.Name
    [void]$c.Activate()
    [void]$c.Range("A1").Select()

    $styleSheet.Delete()
    $styleSheet = $null
    # Calculate only the rebuilt dependency chain.  Some legacy models contain
    # slow external/whole-book dependencies for which CalculateFullRebuild can
    # remain busy indefinitely even though the Catalyst ranges are complete.
    # The Python caller normalizes calcPr to completed Automatic state after the
    # save, so keeping COM in Manual mode here preserves deterministic Data
    # Table caches without reintroducing Excel's stale-value strikethrough flag.
    $s.Calculate()
    $c.Range("C9:C"+$scenarioLast).Calculate()
    $c.Calculate()
    $excel.CalculateBeforeSave = $false
    $wb.ForceFullCalculation = $false
    try { $wb.FullCalculationOnLoad = $false } catch { }
    $wb.Save()
    Write-Host "Built active Catalyst combinations: $scenarioCount scenarios; active targets=$($activeTargets.Count); IDs $scenarioStart-$($scenarioStart+$scenarioCount-1); Table 3 rows $tableTargetRow-$tableInputLast; $Path"
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
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
