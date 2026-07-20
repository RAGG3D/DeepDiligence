param(
    [Parameter(Mandatory = $true)]
    [string]$Path
)

$ErrorActionPreference = "Stop"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $null

try {
    $wb = $excel.Workbooks.Open($Path, 0, $false)

    try { $ws = $wb.Worksheets.Item("RCFS") } catch { $ws = $null }
    if ($ws -ne $null) {
        $ws.Range("W67").Formula = '=IF(W$66<>"",IF(YEAR($F$67)=YEAR(W$66),MAX(W$66-$F$67,0)/365.25,V67+1),"")'
    }
    try { $ws = $wb.Worksheets.Item("Schedules") } catch { $ws = $null }
    if ($ws -ne $null) {
        $ws.Range("F34").Formula = '=IFERROR(SUMIFS(''FY DATA''!F:F,''FY DATA''!$D:$D,"*Deprec*",''FY DATA''!$C:$C,"ISN"),0)'
    }
    try { $ws = $wb.Worksheets.Item("TAM Blood") } catch { $ws = $null }
    if ($ws -ne $null) {
        for ($r = 1; $r -le $ws.UsedRange.Rows.Count; $r++) {
            $cell = $ws.Cells.Item($r,2)
            if ([string]$cell.Formula -like '*#REF!*') {
                $cell.Formula = '=IF(AND($C'+$r+'<>"",COUNTIFS($D:$D,$D'+$r+',$C:$C,$C'+$r+')>1),COUNTIFS($D$1:D'+$r+',$D'+$r+',$C$1:C'+$r+',$C'+$r+'),"")'
            }
        }
    }

    $unresolved = @()
    foreach ($ws in $wb.Worksheets) {
        $formulas = $null
        try { $formulas = $ws.UsedRange.SpecialCells(-4123) } catch { continue }
        foreach ($cell in $formulas.Cells) {
            if ([string]$cell.Formula -like '*#REF!*') { $unresolved += $ws.Name + '!' + $cell.Address($false,$false) }
        }
    }
    if ($unresolved.Count -gt 0) { throw "Unresolved #REF! formulas: " + ($unresolved -join ', ') }

    $wb.Save()
    Write-Host "Known broken references repaired; zero unresolved #REF!: $Path"
}
finally {
    if ($wb -ne $null) { $wb.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
    $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
