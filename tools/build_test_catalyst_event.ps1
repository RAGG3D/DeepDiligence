param(
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [Parameter(Mandatory = $true)]
    [string]$ResearchPath,
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
$test = $null

function ColLetter([int]$n) {
    $s = ""
    while ($n -gt 0) {
        $n--
        $s = [char](65 + ($n % 26)) + $s
        $n = [math]::Floor($n / 26)
    }
    return $s
}

function Rgb([int]$r, [int]$g, [int]$b) {
    return [int]($r + 256*$g + 65536*$b)
}

function HasName([object[]]$items, [string]$name) {
    foreach ($item in $items) {
        if ([string]$item -eq $name) { return $true }
    }
    return $false
}

function NumValue($value, [double]$fallback) {
    if ($null -eq $value -or [string]$value -eq "") { return $fallback }
    try { return [double]$value } catch { return $fallback }
}

function FindOutcome($assessment, [string]$outcome) {
    foreach ($row in @($assessment.outcomes)) {
        if ([string]$row.outcome -eq $outcome) { return $row }
    }
    throw "Outcome $outcome missing from assessment $($assessment.target)"
}

function PasteFormats($source, $destination) {
    [void]$source.Copy()
    [void]$destination.PasteSpecial($xlPasteFormats)
}

function SetGreyMask($range) {
    $range.Interior.Pattern = 1
    $range.Interior.Color = Rgb 231 230 230
    $range.Font.Color = Rgb 127 127 127
}

try {
    $research = Get-Content -Raw -LiteralPath $ResearchPath | ConvertFrom-Json
    if (-not [bool]$research.clinical_only -or
            -not [bool]$research.clinical_interpretation_price_blind) {
        throw "Clinical interpretation is not marked price-blind"
    }
    if (-not [bool]$research.pre_event_price_calibration_used -or
            [bool]$research.same_day_or_post_event_price_data_used) {
        throw "Pre-disclosure price calibration boundary is invalid"
    }
    $priceCalibration = $research.pre_event_price_calibration
    [double]$averageClose = NumValue $priceCalibration.average_close 0.0
    if ($averageClose -le 0) { throw "Pre-disclosure average close is missing" }
    $sheetName = [string]$research.sheet_name
    if (-not $sheetName -or $sheetName.Length -gt 31) { throw "Invalid Test sheet name" }
    $activeNames = @($research.relevant_targets | ForEach-Object { [string]$_ })
    $targetUniverse = @($research.target_universe | ForEach-Object { [string]$_ })
    if ($activeNames.Count -lt 1 -or $targetUniverse.Count -lt $activeNames.Count) {
        throw "Invalid target universe in clinical research artifact"
    }
    $inactiveNames = @($targetUniverse | Where-Object { -not (HasName $activeNames ([string]$_)) })
    $displayTargets = @($activeNames) + @($inactiveNames)
    $assessmentByTarget = @{}
    foreach ($assessment in @($research.target_assessments)) {
        $assessmentByTarget[[string]$assessment.target] = $assessment
    }

    # Conviction is a filter. Build the complete Cartesian product of only the
    # outcomes whose clinical conviction meets the threshold.
    $allowedByTarget = @{}
    foreach ($target in $activeNames) {
        if (-not $assessmentByTarget.ContainsKey($target)) { throw "Assessment missing for $target" }
        $allowed = @()
        foreach ($row in @($assessmentByTarget[$target].outcomes)) {
            if ([double]$row.conviction -ge $ConvictionThreshold) {
                $allowed += [string]$row.outcome
            }
        }
        if ($allowed.Count -lt 1) { throw "No outcome survives threshold for $target" }
        $allowedByTarget[$target] = $allowed
    }
    $combinations = New-Object System.Collections.ArrayList
    function Add-Combinations([int]$index, [object[]]$prefix) {
        if ($index -ge $activeNames.Count) {
            [void]$combinations.Add(@($prefix))
            return
        }
        $target = [string]$activeNames[$index]
        foreach ($outcome in @($allowedByTarget[$target])) {
            Add-Combinations ($index + 1) (@($prefix) + @($outcome))
        }
    }
    Add-Combinations 0 @()
    # Display every surviving Cartesian outcome in descending joint
    # probability. RJConv. is the raw product of the selected target-level
    # convictions; the outcome string resolves exact ties deterministically.
    $rankedCombinations = @()
    foreach ($rawCombo in @($combinations)) {
        $combo = @($rawCombo)
        [double]$probability = 1.0
        for ($a = 0; $a -lt $activeNames.Count; $a++) {
            $target = [string]$activeNames[$a]
            $item = FindOutcome $assessmentByTarget[$target] ([string]$combo[$a])
            $probability *= [double]$item.conviction
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
    if ($scenarioCount -lt 1) { throw "No clinical test scenarios generated" }
    if ($scenarioCount -gt 512) { throw "Clinical scenario count $scenarioCount exceeds safety cap 512" }

    $scenarioFirst = 10
    $scenarioLast = 9 + $scenarioCount
    # Match Catalyst exactly:
    # Scenario | Base Case | Final Market | Upside | RJConv. | outcomes | targets.
    $baseCol = 3
    $finalCol = 4
    $upsideCol = 5
    $rjconvCol = 6
    $outcomeFirstCol = 7
    $groupFirstCol = $outcomeFirstCol + $activeNames.Count
    $lastGroupCol = $groupFirstCol + 4*$displayTargets.Count - 1
    $tableTitleRow = $scenarioLast + 2
    $tableTargetRow = $tableTitleRow + 1
    $tableHeaderRow = $tableTitleRow + 2
    $tableInputFirst = $tableTitleRow + 3
    $tableInputLast = $tableInputFirst + 3
    $evidenceRow = $tableInputLast + 3
    $outcomes = @("Increase", "Remain", "Decrease", "Suspension")

    $wb = $excel.Workbooks.Open($Path, 0, $false)
    $excel.Calculation = $xlCalculationManual
    $source = $wb.Worksheets.Item("Catalyst")
    $s = $wb.Worksheets.Item("Scenarios")
    $v = $wb.Worksheets.Item("VALUATION")

    # Replace only this event's prior Scenarios module. Other historical Test
    # modules remain in place and keep their clinical audit trail.
    $moduleLabel = "Test Scenarios - "+[string]$research.event_name
    $scenarioUsed = $s.UsedRange.Rows.Count
    $oldModuleStart = $null
    $oldModuleEnd = $null
    for ($r = 1; $r -le $scenarioUsed; $r++) {
        if ([string]$s.Cells.Item($r,3).Value2 -eq $moduleLabel) {
            $oldModuleStart = $r
            $oldModuleEnd = $scenarioUsed
            for ($rr = $r + 1; $rr -le $scenarioUsed; $rr++) {
                if ([string]$s.Cells.Item($rr,3).Value2 -like "Test Scenarios - *") {
                    $oldModuleEnd = $rr - 1
                    break
                }
            }
            break
        }
    }
    if ($null -ne $oldModuleStart) {
        [void]$s.Rows.Item($oldModuleStart.ToString()+":"+$oldModuleEnd).Delete()
    }

    # Scenario IDs are workbook-global. Allocate after every remaining actual
    # or Test scenario header so a rerun cannot collide with another event.
    $maxScenarioId = 4
    $scenarioUsed = $s.UsedRange.Rows.Count
    for ($r = 1; $r -le $scenarioUsed; $r++) {
        $candidate = $s.Cells.Item($r,2).Value2
        if (($candidate -is [double] -or $candidate -is [int]) -and
                [string]$s.Cells.Item($r,3).Value2 -ne "") {
            $maxScenarioId = [math]::Max($maxScenarioId,[int]$candidate)
        }
    }
    $scenarioStartId = $maxScenarioId + 1

    for ($i = $wb.Worksheets.Count; $i -ge 1; $i--) {
        if ([string]$wb.Worksheets.Item($i).Name -eq $sheetName) {
            $wb.Worksheets.Item($i).Delete()
        }
    }
    $missing = [Type]::Missing
    $test = $wb.Worksheets.Add($missing, $wb.Worksheets.Item($wb.Worksheets.Count))
    $test.Name = $sheetName

    # Discover reusable colored target-block styles from the live Catalyst tab.
    $sourceGroupCols = @()
    $sourceMaxCol = [math]::Max(20, [int]$source.UsedRange.Columns.Count)
    for ($col = 3; $col -le $sourceMaxCol - 3; $col++) {
        if ([string]$source.Cells.Item(8,$col).Value2 -eq "USD/Share" -and
                [string]$source.Cells.Item(8,$col+1).Value2 -eq "LOA" -and
                [string]$source.Cells.Item(8,$col+2).Value2 -eq "USD/Share" -and
                [string]$source.Cells.Item(8,$col+3).Value2 -eq "LOA") {
            $sourceGroupCols += $col
            $col += 3
        }
    }
    if ($sourceGroupCols.Count -lt 1) { throw "No Catalyst target-block style source found" }
    $grayColor = Rgb 231 230 230
    $coloredSourceCols = @($sourceGroupCols | Where-Object {
        [int]$source.Cells.Item(8,[int]$_).Interior.Color -ne $grayColor
    })
    if ($coloredSourceCols.Count -lt 1) { $coloredSourceCols = @($sourceGroupCols) }

    $sourceByTarget = @{}
    $baseLoaByTarget = @{}
    foreach ($srcCol in $sourceGroupCols) {
        $target = [string]$source.Cells.Item(7,[int]$srcCol).Value2
        if ($target) {
            $sourceByTarget[$target] = [int]$srcCol
            $targetCalibration = $priceCalibration.targets.PSObject.Properties[$target].Value
            if ($null -eq $targetCalibration) { throw "Price calibration missing target: $target" }
            $baseLoaByTarget[$target] = (NumValue $targetCalibration.estimated_loa -1.0)
            if ($baseLoaByTarget[$target] -lt 0 -or $baseLoaByTarget[$target] -gt 1) {
                throw "Invalid calibrated LOA for ${target}: $($baseLoaByTarget[$target])"
            }
        }
    }
    foreach ($target in $targetUniverse) {
        if (-not $sourceByTarget.ContainsKey($target)) {
            throw "Test target missing from Catalyst valuation breakdown: $target"
        }
    }

    $sourceBaseCol = 3
    $sourceFinalCol = 4
    $sourceUpsideCol = 5
    for ($col = 2; $col -le $sourceMaxCol; $col++) {
        $header = [string]$source.Cells.Item(7,$col).Value2
        if ($header -like "Base Case*") { $sourceBaseCol = $col }
        elseif ($header -eq "Final Market Price") { $sourceFinalCol = $col }
        elseif ($header -eq "Upside") { $sourceUpsideCol = $col }
    }

    # Basic title and Catalyst-identical main-table formats.
    PasteFormats $source.Range("B2:E5") $test.Range("B2:E5")
    $mainStylePairs = @(
        @{ Source=2; Destination=2 },
        @{ Source=$sourceBaseCol; Destination=$baseCol },
        @{ Source=$sourceFinalCol; Destination=$finalCol },
        @{ Source=$sourceUpsideCol; Destination=$upsideCol },
        @{ Source=$sourceUpsideCol; Destination=$rjconvCol }
    )
    foreach ($pair in $mainStylePairs) {
        $srcCol = [int]$pair.Source
        $dstCol = [int]$pair.Destination
        PasteFormats $source.Range((ColLetter $srcCol)+"6:"+(ColLetter $srcCol)+"13") `
            $test.Range((ColLetter $dstCol)+"6:"+(ColLetter $dstCol)+"13")
        if ($scenarioLast -gt 13) {
            PasteFormats $source.Cells.Item(13,$srcCol) `
                $test.Range((ColLetter $dstCol)+"14:"+(ColLetter $dstCol)+$scenarioLast)
        }
    }
    $sourceOutcomeCols = @()
    $firstSourceGroup = [int]($sourceGroupCols | Measure-Object -Minimum).Minimum
    for ($col = 6; $col -lt $firstSourceGroup; $col++) {
        if ([string]$source.Cells.Item(8,$col).Value2 -eq "Outcome") { $sourceOutcomeCols += $col }
    }
    if ($sourceOutcomeCols.Count -lt 1) { $sourceOutcomeCols = @(6) }
    for ($i = 0; $i -lt $activeNames.Count; $i++) {
        $dstCol = $outcomeFirstCol + $i
        $srcCol = [int]$sourceOutcomeCols[$i % $sourceOutcomeCols.Count]
        PasteFormats $source.Range((ColLetter $srcCol)+"6:"+(ColLetter $srcCol)+"13") `
            $test.Range((ColLetter $dstCol)+"6:"+(ColLetter $dstCol)+"13")
        if ($scenarioLast -gt 13) {
            PasteFormats $source.Cells.Item(13,$srcCol) `
                $test.Range((ColLetter $dstCol)+"14:"+(ColLetter $dstCol)+$scenarioLast)
        }
        $test.Columns.Item($dstCol).ColumnWidth = 20
    }
    $test.Columns.Item(2).ColumnWidth = 12
    $test.Columns.Item($baseCol).ColumnWidth = 23
    $test.Columns.Item($finalCol).ColumnWidth = 15
    $test.Columns.Item($upsideCol).ColumnWidth = 11
    $test.Columns.Item($rjconvCol).ColumnWidth = 11

    # Apply stable colored Catalyst formats to event targets and preserve a
    # complete visible four-column breakdown for every non-event target.
    for ($i = 0; $i -lt $displayTargets.Count; $i++) {
        $target = [string]$displayTargets[$i]
        $g = $groupFirstCol + 4*$i
        $src = $(if (HasName $activeNames $target) {
            [int]$coloredSourceCols[$i % $coloredSourceCols.Count]
        } else {
            [int]$sourceByTarget[$target]
        })
        PasteFormats $source.Range((ColLetter $src)+"6:"+(ColLetter ($src+3))+"13") `
            $test.Range((ColLetter $g)+"6:"+(ColLetter ($g+3))+"13")
        if ($scenarioLast -gt 13) {
            PasteFormats $source.Range((ColLetter $src)+"13:"+(ColLetter ($src+3))+"13") `
                $test.Range((ColLetter $g)+"14:"+(ColLetter ($g+3))+$scenarioLast)
        }
        $test.Columns.Item($g).ColumnWidth = 13
        $test.Columns.Item($g+1).ColumnWidth = 8
        $test.Columns.Item($g+2).ColumnWidth = 13
        $test.Columns.Item($g+3).ColumnWidth = 8
        for ($cc = 0; $cc -lt 4; $cc++) { $test.Columns.Item($g+$cc).Hidden = $false }
    }

    $test.Range("B2").Value2 = "Clinical Catalyst Test"
    $test.Range("C2").Value2 = [string]$research.event_name
    $test.Range("B3").Value2 = "Completed Disclosure"
    $test.Range("C3").Value2 = ([string]$research.event_date+" | "+[string]$research.venue+" | "+[string]$research.trial+" | "+[string]$research.phase)
    $test.Range("B4").Value2 = "Pre-Disclosure Avg Close"
    $test.Range("C4").Value2 = [double]$averageClose
    $test.Range("C4").NumberFormat = "0.0000"
    $test.Range("D4").Value2 = ([string]$priceCalibration.window_start+" to "+[string]$priceCalibration.window_end_exclusive+" (exclusive); "+[string]$priceCalibration.session_count+" sessions; "+[string]$priceCalibration.currency)
    $test.Range("E4").Value2 = "Source"
    [void]$test.Hyperlinks.Add($test.Range("E4"),[string]$priceCalibration.source_url)
    $test.Range("B5").Value2 = "CLINICAL INTERPRETATION | PRE-DISCLOSURE PRICE-CALIBRATED"

    $test.Cells.Item(7,2).Value2 = "Scenario"
    $test.Cells.Item(8,2).Value2 = "ID"
    $test.Cells.Item(7,$baseCol).Value2 = "Base Case (USD/Share)"
    [double]$testTerminalGrowth = NumValue ($source.Cells.Item(8,$sourceBaseCol).Value2) 0.03
    $test.Cells.Item(8,$baseCol).Value2 = [double]$testTerminalGrowth
    $test.Cells.Item(8,$baseCol).NumberFormat = "0.0%"
    $test.Cells.Item(7,$finalCol).Value2 = "Final Market Price"
    $test.Cells.Item(8,$finalCol).Value2 = "USD/Share"
    $test.Cells.Item(7,$upsideCol).Value2 = "Upside"
    $test.Cells.Item(8,$upsideCol).Value2 = "%"
    $test.Cells.Item(7,$rjconvCol).Value2 = "RJConv."
    $test.Cells.Item(8,$rjconvCol).Value2 = "%"
    for ($i = 0; $i -lt $activeNames.Count; $i++) {
        $test.Cells.Item(7,$outcomeFirstCol+$i).Value2 = [string]$activeNames[$i]
        $test.Cells.Item(8,$outcomeFirstCol+$i).Value2 = "Outcome"
    }
    for ($i = 0; $i -lt $displayTargets.Count; $i++) {
        $g = $groupFirstCol + 4*$i
        $test.Cells.Item(7,$g).Value2 = [string]$displayTargets[$i]
        $test.Cells.Item(7,$g+2).Value2 = "Market Price"
        $test.Cells.Item(8,$g).Value2 = "USD/Share"
        $test.Cells.Item(8,$g+1).Value2 = "LOA"
        $test.Cells.Item(8,$g+2).Value2 = "USD/Share"
        $test.Cells.Item(8,$g+3).Value2 = "LOA"
        [void]$test.Range($test.Cells.Item(7,$g),$test.Cells.Item(7,$g+1)).Merge()
        [void]$test.Range($test.Cells.Item(7,$g+2),$test.Cells.Item(7,$g+3)).Merge()
    }

    # Native two-variable What-If inputs mirror Catalyst. Row 9 is Base; rows
    # below are the conviction-filtered Cartesian scenarios. The actual
    # observed outcome remains explicitly recorded in the evidence section.
    $test.Cells.Item(6,2).Value2 = 1.0
    $test.Cells.Item(6,3).Formula = "=`$C`$8"
    $test.Range("B6:C6").NumberFormat = ";;;"
    $test.Cells.Item(8,2).Formula = [string]$source.Cells.Item(8,2).Formula
    $test.Cells.Item(8,2).NumberFormat = '"ID"'
    $test.Cells.Item(9,2).Value2 = 1.0
    $test.Cells.Item(9,2).NumberFormat = '"Base"'
    for ($i = 0; $i -lt $activeNames.Count; $i++) {
        $test.Cells.Item(9,$outcomeFirstCol+$i).Value2 = "Base"
    }

    $baseL = ColLetter $baseCol
    $finalL = ColLetter $finalCol
    $marketPriceCols = @()
    for ($i = 0; $i -lt $displayTargets.Count; $i++) {
        $target = [string]$displayTargets[$i]
        $g = $groupFirstCol + 4*$i
        $gL = ColLetter $g
        $marketL = ColLetter ($g+2)
        $marketLoaL = ColLetter ($g+3)
        $sourceFormula = [string]$source.Cells.Item(9,[int]$sourceByTarget[$target]).Formula
        if (-not $sourceFormula) { throw "Catalyst valuation formula missing for $target" }
        $test.Cells.Item(9,$g).Formula = $sourceFormula
        $test.Cells.Item(9,$g+1).Value2 = 1.0
        $test.Cells.Item(9,$g+3).Value2 = [double]$baseLoaByTarget[$target]
        $test.Cells.Item(9,$g+2).Formula = "=IFERROR(`$"+$gL+"9*`$"+$marketLoaL+"9,0)"
        $test.Cells.Item(6,$g).Formula = "=IFERROR(`$"+$gL+"`$9/`$"+$baseL+"`$9,0)"
        $test.Cells.Item(6,$g).NumberFormat = "0.0%"
        $marketPriceCols += $marketL
    }
    $test.Cells.Item(9,$finalCol).Formula = "=IFERROR(SUM("+(($marketPriceCols | ForEach-Object {"`$"+$_+"9"}) -join ",")+"),0)"
    $test.Cells.Item(9,$upsideCol).Formula = "=IFERROR(`$"+$finalL+"9/`$C`$4-1,0)"
    $test.Cells.Item(9,$rjconvCol).Value2 = 1.0
    $test.Cells.Item(9,$rjconvCol).NumberFormat = "0.0%"

    for ($index = 0; $index -lt $scenarioCount; $index++) {
        $rowNum = $scenarioFirst + $index
        $combo = @($combinations[$index])
        $test.Cells.Item($rowNum,2).Value2 = [double]($scenarioStartId+$index)
        for ($a = 0; $a -lt $activeNames.Count; $a++) {
            $outcome = [string]$combo[$a]
            $test.Cells.Item($rowNum,$outcomeFirstCol+$a).Value2 = $outcome
        }
        for ($i = 0; $i -lt $displayTargets.Count; $i++) {
            $target = [string]$displayTargets[$i]
            $g = $groupFirstCol + 4*$i
            $gL = ColLetter $g
            $marketL = ColLetter ($g+2)
            $marketLoaL = ColLetter ($g+3)
            $test.Cells.Item($rowNum,$g).Formula = "=IFERROR(`$"+$baseL+$rowNum+"*`$"+$gL+"`$6,0)"
            $test.Cells.Item($rowNum,$g+1).Value2 = 1.0
            if (HasName $activeNames $target) {
                $activeIndex = [array]::IndexOf($activeNames,$target)
                $outcomeL = ColLetter ($outcomeFirstCol+$activeIndex)
                $tableLoaL = ColLetter ($g+1)
                $outcomeRef = "`$"+$outcomeL+$rowNum
                $test.Cells.Item($rowNum,$g+3).Formula = "=IF("+$outcomeRef+"=`"Suspension`",0,MAX(0,`$"+$marketLoaL+"`$9+IFERROR(INDEX(`$"+$tableLoaL+"`$"+$tableInputFirst+":`$"+$tableLoaL+"`$"+$tableInputLast+",MATCH("+$outcomeRef+",`$B`$"+$tableInputFirst+":`$B`$"+$tableInputLast+",0)),0)))"
            } else {
                $test.Cells.Item($rowNum,$g+3).Formula = "=`$"+$marketLoaL+"`$9"
            }
            $test.Cells.Item($rowNum,$g+2).Formula = "=IFERROR(`$"+$gL+$rowNum+"*`$"+$marketLoaL+$rowNum+",0)"
        }
        $test.Cells.Item($rowNum,$finalCol).Formula = "=IFERROR(SUM("+(($marketPriceCols | ForEach-Object {"`$"+$_+$rowNum}) -join ",")+"),0)"
        $test.Cells.Item($rowNum,$upsideCol).Formula = "=IFERROR(`$"+$finalL+$rowNum+"/`$C`$4-1,0)"
        $convTerms = @()
        for ($a = 0; $a -lt $activeNames.Count; $a++) {
            $target = [string]$activeNames[$a]
            $displayIndex = [array]::IndexOf($displayTargets,$target)
            $targetGroup = $groupFirstCol + 4*$displayIndex
            $tableConvL = ColLetter ($targetGroup+2)
            $outcomeL = ColLetter ($outcomeFirstCol+$a)
            $convTerms += "IFERROR(INDEX(`$"+$tableConvL+"`$"+$tableInputFirst+":`$"+$tableConvL+"`$"+$tableInputLast+",MATCH(`$"+$outcomeL+$rowNum+",`$B`$"+$tableInputFirst+":`$B`$"+$tableInputLast+",0)),0)"
        }
        $test.Cells.Item($rowNum,$rjconvCol).Formula = "=IFERROR(PRODUCT("+($convTerms -join ",")+"),0)"
        $test.Cells.Item($rowNum,$rjconvCol).NumberFormat = "0.0%"
    }

    # Clinical assumptions table. Active values are model-generated from the
    # completed disclosure and matched competitors; inactive values are neutral.
    $test.Cells.Item($tableTitleRow,2).Value2 = "Clinical Outcome Inputs (Table 3; Conviction Filter >= "+($ConvictionThreshold.ToString("0%"))+")"
    $test.Cells.Item($tableHeaderRow,2).Value2 = "Outcome"
    for ($j = 0; $j -lt 4; $j++) { $test.Cells.Item($tableInputFirst+$j,2).Value2 = $outcomes[$j] }
    for ($i = 0; $i -lt $displayTargets.Count; $i++) {
        $target = [string]$displayTargets[$i]
        $g = $groupFirstCol + 4*$i
        $src = [int]$coloredSourceCols[$i % $coloredSourceCols.Count]
        PasteFormats $source.Range((ColLetter $src)+"7:"+(ColLetter ($src+2))+"12") `
            $test.Range((ColLetter $g)+$tableTargetRow+":"+(ColLetter ($g+2))+$tableInputLast)
        $test.Cells.Item($tableTargetRow,$g).Value2 = $target
        $test.Cells.Item($tableHeaderRow,$g).Value2 = "Market Share Change"
        $test.Cells.Item($tableHeaderRow,$g+1).Value2 = "LOA Change"
        $test.Cells.Item($tableHeaderRow,$g+2).Value2 = "Conv."
        for ($j = 0; $j -lt 4; $j++) {
            $rr = $tableInputFirst+$j
            $outcome = $outcomes[$j]
            if (HasName $activeNames $target) {
                $item = FindOutcome $assessmentByTarget[$target] $outcome
                $test.Cells.Item($rr,$g).Value2 = [double]$item.market_share_change
                $test.Cells.Item($rr,$g+1).Value2 = [double]$item.loa_change
                $test.Cells.Item($rr,$g+2).Value2 = [double]$item.conviction
            } else {
                $test.Cells.Item($rr,$g).Value2 = 0.0
                $test.Cells.Item($rr,$g+1).Value2 = 0.0
                $test.Cells.Item($rr,$g+2).Value2 = [double]$(if ($outcome -eq "Remain") {1.0} else {0.0})
            }
            $test.Range($test.Cells.Item($rr,$g),$test.Cells.Item($rr,$g+2)).NumberFormat = "0.0%"
        }
    }

    # Keep the semantic LOA formatting stable across repeated rebuilds. Active
    # final LOA columns share one input style; in inactive groups the first LOA
    # column follows the correctly formatted final LOA column before the visual
    # mask is applied.
    $activeLoaStyleCol = $groupFirstCol + 3
    for ($i = 0; $i -lt $displayTargets.Count; $i++) {
        $target = [string]$displayTargets[$i]
        $g = $groupFirstCol + 4*$i
        $isActive = HasName $activeNames $target
        foreach ($rr in @(6) + @(8..$scenarioLast)) {
            if ($isActive) {
                PasteFormats $test.Cells.Item($rr,$activeLoaStyleCol) $test.Cells.Item($rr,$g+3)
            } else {
                PasteFormats $test.Cells.Item($rr,$g+3) $test.Cells.Item($rr,$g+1)
            }
        }
        if (-not $isActive) {
            SetGreyMask $test.Range($test.Cells.Item(6,$g),$test.Cells.Item($scenarioLast,$g+3))
            SetGreyMask $test.Range($test.Cells.Item($tableTargetRow,$g),$test.Cells.Item($tableInputLast,$g+3))
        }
    }

    # Add the matching clinical Test module to Scenarios.  It has the same
    # target-by-target market-share logic as actual Catalyst scenarios, but the
    # selected outcome and MS-change lookup come only from the Test-EVENT tab.
    $absFirst = $null
    $absLast = $null
    $currentAsset = $null
    $scenarioTargetRows = @()
    $scenarioAssetRows = @()
    $scenarioUsed = $s.UsedRange.Rows.Count
    for ($r = 1; $r -le $scenarioUsed; $r++) {
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
            $scenarioTargetRows += @{ Name=$name; MsRow=$r }
        } elseif ($unit -ne "[%]") {
            $raw = [string]$s.Cells.Item($r,3).Value2
            $currentAsset = ($raw -split ' \(')[0].Trim()
            $scenarioAssetRows += $r
        }
    }
    if ($null -eq $absFirst -or $null -eq $absLast) { throw "Scenarios Absolute template missing" }
    foreach ($target in $targetUniverse) {
        if (-not (HasName @($scenarioTargetRows | ForEach-Object { [string]$_.Name }) $target)) {
            throw "Test target absent from Scenarios Absolute: $target"
        }
    }

    $actualDivider = $null
    for ($r = 1; $r -le $scenarioUsed; $r++) {
        if ([string]$s.Cells.Item($r,3).Value2 -eq "Catalyst Scenarios") { $actualDivider = $r; break }
    }
    if ($null -eq $actualDivider) { throw "Catalyst Scenarios divider missing" }
    $templateHeader = $actualDivider + 1
    $blockHeight = $absLast - $absFirst + 2
    $templateLast = $templateHeader + $blockHeight - 1
    $moduleDivider = $s.UsedRange.Rows.Count + 2
    [void]$s.Range("A"+$actualDivider+":AE"+$actualDivider).Copy($s.Range("A"+$moduleDivider+":AE"+$moduleDivider))
    $s.Cells.Item($moduleDivider,2).ClearContents()
    $s.Cells.Item($moduleDivider,3).Value2 = $moduleLabel
    $quotedTestSheet = "'"+$sheetName.Replace("'","''")+"'!"

    for ($index = 0; $index -lt $scenarioCount; $index++) {
        $h = $moduleDivider + 1 + ($blockHeight+1)*$index
        [void]$s.Range("A"+$templateHeader+":AE"+$templateLast).Copy(
            $s.Range("A"+$h+":AE"+($h+$blockHeight-1))
        )
        $sid = $scenarioStartId + $index
        $combo = @($combinations[$index])
        $parts = @()
        for ($a = 0; $a -lt $activeNames.Count; $a++) {
            $parts += ([string]$activeNames[$a]+": "+[string]$combo[$a])
        }
        $s.Cells.Item($h,2).Value2 = [double]$sid
        $s.Cells.Item($h,3).Value2 = ($parts -join " | ")
        foreach ($assetRow in $scenarioAssetRows) {
            $s.Cells.Item($h + ([int]$assetRow-$absFirst+1),25).ClearContents()
        }
        foreach ($targetRow in $scenarioTargetRows) {
            $target = [string]$targetRow.Name
            $destMsRow = $h + ([int]$targetRow.MsRow-$absFirst+1)
            if (HasName $activeNames $target) {
                $activeIndex = [array]::IndexOf($activeNames,$target)
                $displayIndex = [array]::IndexOf($displayTargets,$target)
                $outcomeL = ColLetter ($outcomeFirstCol+$activeIndex)
                $msInputL = ColLetter ($groupFirstCol + 4*$displayIndex)
                $testMainRow = $scenarioFirst + $index
                $outcomeRef = $quotedTestSheet+"`$"+$outcomeL+"`$"+$testMainRow
                $msRange = $quotedTestSheet+"`$"+$msInputL+"`$"+$tableInputFirst+":`$"+$msInputL+"`$"+$tableInputLast
                $outcomeRange = $quotedTestSheet+"`$B`$"+$tableInputFirst+":`$B`$"+$tableInputLast
                $s.Cells.Item($destMsRow,25).Formula = "=IF("+$outcomeRef+"=`"Suspension`",0,MAX(0,`$Y`$"+[int]$targetRow.MsRow+"+IFERROR(INDEX("+$msRange+",MATCH("+$outcomeRef+","+$outcomeRange+",0)),0)))"
            } else {
                $s.Cells.Item($destMsRow,25).Formula = "=`$Y`$"+[int]$targetRow.MsRow
            }
        }
    }

    # Route the model's two canonical What-If inputs through every Catalyst/Test
    # local bridge. All sheets rest at the neutral values (scenario 1 and 3%);
    # while one native Data Table substitutes its own bridge, the other sheets
    # remain neutral. This lets actual Catalyst and every historical Test tab
    # retain independent, refreshable price tables without importing observed
    # security data.
    $scenarioTerms = @("Catalyst!`$B`$6")
    $growthTerms = @("Catalyst!`$C`$6")
    foreach ($ws in @($wb.Worksheets)) {
        if ([string]$ws.Name -notlike "Test-*") { continue }
        $banner = [string]$ws.Range("B5").Value2
        if ($banner -ne "CLINICAL INPUTS ONLY | MODEL-IMPLIED VALUATION" -and
                $banner -ne "CLINICAL INTERPRETATION | PRE-DISCLOSURE PRICE-CALIBRATED") { continue }
        $quoted = "'"+([string]$ws.Name).Replace("'","''")+"'!"
        $scenarioTerms += $quoted+"`$B`$6"
        $growthTerms += $quoted+"`$C`$6"
    }
    $neutralTests = $scenarioTerms.Count - 1
    $v.Range("C3").Formula = "="+($scenarioTerms -join "+")+$(if ($neutralTests -gt 0) { "-"+$neutralTests } else { "" })
    $v.Range("C5").Formula = "="+($growthTerms -join "+")+$(if ($neutralTests -gt 0) { "-"+$neutralTests+"*3%" } else { "" })

    # Embed the Catalyst-identical native two-variable What-If table in B:C.
    # Column B supplies globally unique Test scenario IDs; row C8 supplies the
    # fixed terminal-growth observation. Results in C are model-implied; the
    # row-9 risk-adjusted breakdown is calibrated to the pre-disclosure close.
    try {
        if ($test.Range("C9").HasArray) { [void]$test.Range("C9").CurrentArray.Clear() }
    } catch { }
    [void]$test.Range("B8:C"+$scenarioLast).Table($test.Range("C6"),$test.Range("B6"))
    $test.Range("C9:C"+$scenarioLast).NumberFormat = "0.00"

    # Keep the bounded pre-disclosure closes and implied target LOAs separate
    # from the clinical evidence so the calibration cannot contaminate the
    # medical judgement.
    $r = $evidenceRow
    $test.Cells.Item($r,2).Value2 = "Pre-Disclosure Price Calibration"; $r++
    $headers = @("Session","Raw Close","Currency","Provider","Window End (Exclusive)")
    for ($i = 0; $i -lt $headers.Count; $i++) { $test.Cells.Item($r,2+$i).Value2 = $headers[$i] }
    $r++
    foreach ($session in @($priceCalibration.sessions)) {
        $test.Cells.Item($r,2).Value2 = [string]$session.date
        $test.Cells.Item($r,3).Value2 = [double]$session.close
        $test.Cells.Item($r,3).NumberFormat = "0.0000"
        $test.Cells.Item($r,4).Value2 = [string]$priceCalibration.currency
        $test.Cells.Item($r,5).Value2 = [string]$priceCalibration.provider
        $test.Cells.Item($r,6).Value2 = [string]$priceCalibration.window_end_exclusive
        $r++
    }
    $test.Cells.Item($r,2).Value2 = "Average"
    $test.Cells.Item($r,3).Value2 = [double]$averageClose
    $test.Cells.Item($r,3).NumberFormat = "0.0000"
    $r += 2
    $test.Cells.Item($r,2).Value2 = "Implied Baseline LOA Breakdown"; $r++
    $headers = @("Target","Unrisked Value","Prior LOA","Estimated LOA","Calibrated Market Price")
    for ($i = 0; $i -lt $headers.Count; $i++) { $test.Cells.Item($r,2+$i).Value2 = $headers[$i] }
    $r++
    foreach ($target in $targetUniverse) {
        $item = $priceCalibration.targets.PSObject.Properties[[string]$target].Value
        $test.Cells.Item($r,2).Value2 = [string]$target
        $test.Cells.Item($r,3).Value2 = [double]$item.unrisked_value
        $test.Cells.Item($r,4).Value2 = [double]$item.prior_loa
        $test.Cells.Item($r,5).Value2 = [double]$item.estimated_loa
        $test.Cells.Item($r,6).Value2 = [double]$item.calibrated_market_price
        $test.Range($test.Cells.Item($r,4),$test.Cells.Item($r,5)).NumberFormat = "0.0%"
        $test.Range($test.Cells.Item($r,3),$test.Cells.Item($r,3)).NumberFormat = "0.0000"
        $test.Range($test.Cells.Item($r,6),$test.Cells.Item($r,6)).NumberFormat = "0.0000"
        $r++
    }
    $r++
    $test.Cells.Item($r,2).Value2 = "Observed Clinical Readout"; $r++
    foreach ($header in @("Target","Observed Outcome","Data Quality","Clinical Interpretation")) {
        $test.Cells.Item($r,1+[array]::IndexOf(@("Target","Observed Outcome","Data Quality","Clinical Interpretation"),$header)+1).Value2 = $header
    }
    $r++
    foreach ($target in $activeNames) {
        $a = $assessmentByTarget[$target]
        $test.Cells.Item($r,2).Value2 = $target
        $test.Cells.Item($r,3).Value2 = [string]$a.observed_outcome
        $test.Cells.Item($r,4).Value2 = [string]$a.data_quality
        $test.Cells.Item($r,5).Value2 = [string]$a.clinical_interpretation
        $r++
    }
    $r++
    $test.Cells.Item($r,2).Value2 = "Reported Data"; $r++
    $headers = @("Target","Population","N Enrolled","N Evaluable","Efficacy","Safety","Follow-up","Source IDs")
    for ($i = 0; $i -lt $headers.Count; $i++) { $test.Cells.Item($r,2+$i).Value2 = $headers[$i] }
    $r++
    foreach ($item in @($research.reported_data)) {
        $values = @([string]$item.target,[string]$item.population,[double]$item.n_enrolled,[double]$item.n_evaluable,[string]$item.efficacy,[string]$item.safety,[string]$item.follow_up,(@($item.source_ids) -join ", "))
        for ($i = 0; $i -lt $values.Count; $i++) {
            if ($i -eq 2 -or $i -eq 3) {
                $test.Cells.Item($r,2+$i).Value2 = [double]$values[$i]
            } else {
                $test.Cells.Item($r,2+$i).Value2 = [string]$values[$i]
            }
        }
        $r++
    }
    $r++
    $test.Cells.Item($r,2).Value2 = "Matched Competitor Clinical Data"; $r++
    $headers = @("Target","Competitor","Matched Setting","Clinical Comparison","Limitations","Source IDs")
    for ($i = 0; $i -lt $headers.Count; $i++) { $test.Cells.Item($r,2+$i).Value2 = $headers[$i] }
    $r++
    foreach ($item in @($research.competitor_comparisons)) {
        $values = @([string]$item.target,[string]$item.competitor,[string]$item.matched_setting,[string]$item.clinical_comparison,[string]$item.limitations,(@($item.source_ids) -join ", "))
        for ($i = 0; $i -lt $values.Count; $i++) { $test.Cells.Item($r,2+$i).Value2 = $values[$i] }
        $r++
    }
    $r++
    $test.Cells.Item($r,2).Value2 = "Clinical Sources"; $r++
    $headers = @("ID","Kind","Title","URL","Accessed")
    for ($i = 0; $i -lt $headers.Count; $i++) { $test.Cells.Item($r,2+$i).Value2 = $headers[$i] }
    $r++
    foreach ($item in @($research.sources)) {
        $test.Cells.Item($r,2).Value2 = [string]$item.id
        $test.Cells.Item($r,3).Value2 = [string]$item.source_kind
        $test.Cells.Item($r,4).Value2 = [string]$item.title
        $test.Cells.Item($r,5).Value2 = [string]$item.url
        [void]$test.Hyperlinks.Add($test.Cells.Item($r,5),[string]$item.url)
        $test.Cells.Item($r,6).Value2 = [string]$item.accessed
        $r++
    }
    $test.Range("B"+$evidenceRow+":"+(ColLetter ([math]::Min($lastGroupCol,20)))+$r).WrapText = $true
    $test.Columns.Item(5).ColumnWidth = 55
    $test.Range("B2:"+(ColLetter $lastGroupCol)+$r).Font.Name = $source.Range("B5").Font.Name
    [void]$test.Activate()
    [void]$test.Range("A1").Select()
    # Calculate only the Test sheet and its native What-If table.  Whole-book
    # rebuilds can hang on legacy external dependencies unrelated to this tab.
    # The Python caller immediately normalizes calcPr to completed Automatic
    # state, while this scoped calculation preserves deterministic Test caches.
    $test.Range("C9:C"+$scenarioLast).Calculate()
    $test.Calculate()
    $excel.CalculateBeforeSave = $false
    $wb.ForceFullCalculation = $false
    try { $wb.FullCalculationOnLoad = $false } catch { }
    $wb.Save()
    Write-Host "Built pre-disclosure-price-calibrated clinical test sheet ${sheetName}: $scenarioCount filtered scenarios, $($activeNames.Count) event targets, $($displayTargets.Count) visible targets, base=$averageClose"
}
catch {
    Write-Host $_.InvocationInfo.PositionMessage
    Write-Host $_.ScriptStackTrace
    throw
}
finally {
    if ($wb -ne $null) {
        try { $wb.Close($false) } catch { Write-Host "Workbook close warning: $($_.Exception.Message)" }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
    }
    try { $excel.Quit() } catch { }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [gc]::Collect(); [gc]::WaitForPendingFinalizers()
}
