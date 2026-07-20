param(
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [Parameter(Mandatory = $true)]
    [string]$BackupPath,
    [string]$DividerLabel = "Catalyst Scenarios"
)

$ErrorActionPreference = "Stop"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$targetBook = $null
$backupBook = $null

try {
    $targetBook = $excel.Workbooks.Open($Path, 0, $false)
    $backupBook = $excel.Workbooks.Open($BackupPath, 0, $true)
    $target = $targetBook.Worksheets.Item("Scenarios")
    $source = $backupBook.Worksheets.Item("Scenarios")

    function Find-Divider($sheet, [string]$label) {
        $used = [int]$sheet.UsedRange.Rows.Count
        for ($row = 1; $row -le $used; $row++) {
            if ([string]$sheet.Cells.Item($row,3).Value2 -eq $label) { return $row }
        }
        throw "Divider '$label' not found in $([string]$sheet.Parent.Name)"
    }

    $targetDivider = Find-Divider $target $DividerLabel
    $sourceDivider = Find-Divider $source $DividerLabel
    $targetUsed = [int]$target.UsedRange.Rows.Count
    $sourceUsed = [int]$source.UsedRange.Rows.Count

    # This recovery is intentionally narrow: only the scenario section from
    # the named divider to the end of the Scenarios sheet is replaced.  All
    # assumptions, financial statements and every other sheet remain intact.
    if ($targetUsed -ge $targetDivider) {
        [void]$target.Rows.Item($targetDivider.ToString()+":"+$targetUsed).Delete()
    }
    $sourceRange = $source.Range("A"+$sourceDivider+":AE"+$sourceUsed)
    [void]$sourceRange.Copy($target.Range("A"+$targetDivider))
    for ($offset = 0; $offset -le ($sourceUsed-$sourceDivider); $offset++) {
        $target.Rows.Item($targetDivider+$offset).RowHeight = $source.Rows.Item($sourceDivider+$offset).RowHeight
    }
    $targetBook.Save()
    Write-Host "Restored Scenarios section from rows $sourceDivider-$sourceUsed into row ${targetDivider}: $Path"
}
finally {
    if ($backupBook -ne $null) {
        try { $backupBook.Close($false) } catch { }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($backupBook)
    }
    if ($targetBook -ne $null) {
        try { $targetBook.Close($false) } catch { }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($targetBook)
    }
    try { $excel.Quit() } catch { }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
