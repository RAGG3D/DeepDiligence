param(
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [Parameter(Mandatory = $true)]
    [string]$ScorePath
)

$ErrorActionPreference = "Stop"
$xlCalculationManual = -4135
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $null

function PasteFormats($source, $destination) {
    [void]$source.Copy()
    [void]$destination.PasteSpecial(-4122)
}

try {
    $score = Get-Content -Raw -Encoding UTF8 -LiteralPath $ScorePath | ConvertFrom-Json
    $sheetName = "Test-"+[string]$score.event
    if ($sheetName.Length -gt 31) { $sheetName = $sheetName.Substring(0,31) }
    if (@($score.post_release.sessions).Count -ne 3) {
        throw "Exactly three post-release sessions are required"
    }

    $wb = $excel.Workbooks.Open($Path, 0, $false)
    $excel.Calculation = $xlCalculationManual
    $test = $wb.Worksheets.Item($sheetName)
    if ([string]$test.Range("B5").Value2 -ne "CLINICAL INTERPRETATION | PRE-DISCLOSURE PRICE-CALIBRATED") {
        throw "Frozen blind Test banner is missing"
    }

    # The only top-table post-release cells are D2:F2 and the final-LOA result
    # column in row 2 for each active target.  No prediction row is touched.
    for ($i = 0; $i -lt 3; $i++) {
        $session = @($score.post_release.sessions)[$i]
        $cell = $test.Cells.Item(2,4+$i)
        PasteFormats $test.Range("C2") $cell
        $shortDate = ([datetime]::ParseExact([string]$session.date,"yyyy-MM-dd",$null)).ToString("MM/dd")
        [double]$close = [double]$session.close
        $cell.Value2 = ($shortDate+" `$"+$close.ToString("0.0000"))
        $cell.HorizontalAlignment = -4108
        $cell.Font.Bold = $true
        try { $cell.ClearComments() } catch { }
        [void]$cell.AddComment(
            "Post-release raw Close: "+[string]$session.date+" | "+$close.ToString("0.0000")+
            " | "+[string]$score.post_release.source_url
        )
    }

    $maxCol = [int]$test.UsedRange.Columns.Count
    foreach ($item in @($score.target_scores)) {
        $target = [string]$item.target
        $groupCol = $null
        for ($col = 6; $col -le $maxCol-3; $col++) {
            if ([string]$test.Cells.Item(7,$col).Value2 -eq $target -and
                    [string]$test.Cells.Item(8,$col).Value2 -eq "USD/Share" -and
                    [string]$test.Cells.Item(8,$col+3).Value2 -eq "LOA") {
                $groupCol = $col
                break
            }
        }
        if ($null -eq $groupCol -or [string]$test.Cells.Item(8,$groupCol+3).Value2 -ne "LOA") {
            throw "Cannot resolve final blind LOA result column for $target"
        }
        $cell = $test.Cells.Item(2,$groupCol+3)
        PasteFormats $test.Range("C2") $cell
        $cell.Interior.Pattern = 1
        $cell.Interior.Color = $test.Cells.Item(7,$groupCol).Interior.Color
        $cell.Font.Color = $test.Cells.Item(7,$groupCol).Font.Color
        $cell.Font.Bold = $true
        $cell.HorizontalAlignment = -4108
        $cell.NumberFormat = "@"
        $cell.Value2 = ([string]$item.score+"/10")
        try { $cell.ClearComments() } catch { }
        [double]$baseLoa = [double]$test.Cells.Item(9,$groupCol+3).Value2
        [void]$cell.AddComment(
            "Independent post-release score | "+$target+" | Base LOA "+$baseLoa.ToString("0.0%")+
            " | "+[string]$item.rationale
        )
    }

    # Append a compact audit trail.  It is intentionally separate from all
    # blind prediction and scenario cells.
    $r = [int]$test.UsedRange.Rows.Count + 3
    $test.Cells.Item($r,2).Value2 = "Independent Post-Release Backtest"
    $test.Cells.Item($r,2).Font.Bold = $true
    $r++
    $test.Cells.Item($r,2).Value2 = "Release Date"
    $test.Cells.Item($r,3).Value2 = [string]$score.post_release.release_date
    $test.Cells.Item($r,4).Value2 = "Window Rule"
    $test.Cells.Item($r,5).Value2 = [string]$score.post_release.release_timing_basis
    $test.Cells.Item($r,5).WrapText = $true
    $r++
    foreach ($header in @("Session","Raw Close","Source")) {
        $test.Cells.Item($r,1+[array]::IndexOf(@("Session","Raw Close","Source"),$header)+1).Value2 = $header
    }
    $r++
    foreach ($session in @($score.post_release.sessions)) {
        $test.Cells.Item($r,2).Value2 = [string]$session.date
        $test.Cells.Item($r,3).Value2 = [double]$session.close
        $test.Cells.Item($r,3).NumberFormat = "0.0000"
        $test.Cells.Item($r,4).Value2 = [string]$score.post_release.source_url
        [void]$test.Hyperlinks.Add($test.Cells.Item($r,4),[string]$score.post_release.source_url)
        $r++
    }
    $r++
    $test.Cells.Item($r,2).Value2 = "Target"
    $test.Cells.Item($r,3).Value2 = "Score"
    $test.Cells.Item($r,4).Value2 = "Brief Review"
    $r++
    foreach ($item in @($score.target_scores)) {
        $test.Cells.Item($r,2).Value2 = [string]$item.target
        $test.Cells.Item($r,3).Value2 = [double]$item.score
        $test.Cells.Item($r,3).NumberFormat = "0`/10"
        $test.Cells.Item($r,4).Value2 = [string]$item.rationale
        $r++
    }
    $r++
    $test.Cells.Item($r,2).Value2 = "Three-Day Intraday High Test"
    $test.Cells.Item($r,2).Font.Bold = $true
    $r++
    $test.Cells.Item($r,2).Value2 = "Session"
    $test.Cells.Item($r,3).Value2 = "60m Raw High"
    $test.Cells.Item($r,4).Value2 = "Peak Bar Start"
    $r++
    foreach ($item in @($score.intraday.daily_highs)) {
        $test.Cells.Item($r,2).Value2 = [string]$item.date
        $test.Cells.Item($r,3).Value2 = [double]$item.raw_high
        $test.Cells.Item($r,3).NumberFormat = "0.0000"
        $test.Cells.Item($r,4).Value2 = [string]$item.bar_start
        $r++
    }
    $r++
    $test.Cells.Item($r,2).Value2 = "Highest-RJConv. Scenario"
    $test.Cells.Item($r,3).Value2 = "RJConv."
    $test.Cells.Item($r,4).Value2 = "Blind Price"
    $test.Cells.Item($r,5).Value2 = "3-Day Peak"
    $test.Cells.Item($r,6).Value2 = "Result"
    $r++
    foreach ($item in @($score.highest_rjconv_scenario_test.scenarios)) {
        [double]$difference = [double]$item.difference_usd_per_share
        [double]$differencePct = [double]$item.difference_as_pct_of_real_peak
        $resultText = $(if ([string]$item.result -eq "MISS") {
            "MISS | blind exceeds peak by `$"+$difference.ToString("0.0000")+" | "+$differencePct.ToString("0.0%")+" above real peak"
        } else {
            "REACHED | peak exceeds blind by `$"+$difference.ToString("0.0000")+" | "+$differencePct.ToString("0.0%")+" profit below real peak"
        })
        $test.Cells.Item($r,2).Value2 = ("Scenario "+[string]$item.scenario_id+" | "+(($item.outcomes | ConvertTo-Json -Compress)))
        $test.Cells.Item($r,3).Value2 = [double]$item.rjconv
        $test.Cells.Item($r,3).NumberFormat = "0.0000%"
        $test.Cells.Item($r,4).Value2 = [double]$item.blind_final_market_price
        $test.Cells.Item($r,4).NumberFormat = "0.0000"
        $test.Cells.Item($r,5).Value2 = [double]$score.highest_rjconv_scenario_test.three_day_real_peak
        $test.Cells.Item($r,5).NumberFormat = "0.0000"
        $test.Cells.Item($r,6).Value2 = $resultText
        $r++
    }
    $test.Range("B"+($r-@($score.target_scores).Count-1)+":D"+$r).WrapText = $true
    $test.Columns.Item(4).ColumnWidth = [math]::Max([double]$test.Columns.Item(4).ColumnWidth,55)
    $test.Columns.Item(6).ColumnWidth = [math]::Max([double]$test.Columns.Item(6).ColumnWidth,42)

    $excel.CalculateBeforeSave = $false
    $wb.ForceFullCalculation = $false
    try { $wb.FullCalculationOnLoad = $false } catch { }
    $wb.Save()
    Write-Host "Wrote independent post-release overlay to $sheetName"
}
catch {
    Write-Host $_.InvocationInfo.PositionMessage
    Write-Host $_.ScriptStackTrace
    throw
}
finally {
    if ($wb -ne $null) {
        try { $wb.Close($false) } catch { }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
    }
    try { $excel.Quit() } catch { }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [gc]::Collect(); [gc]::WaitForPendingFinalizers()
}
