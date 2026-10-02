# Recalculate a workbook with desktop Excel (COM), scan every sheet for error values, print chosen cells, and save
# so cached values are stored. Nothing is hard-coded: pass the workbook and the cells you want echoed.
#
#   powershell -ExecutionPolicy Bypass -File recalc_excel.ps1 -Path C:\out\book.xlsx `
#       -Cells "Time Spent!D7|Time Spent!E7|Summary!B4:B13"      (cells separated by |)
#   Add -NoSave to leave the file untouched (read-only check).
#
# Notes: if the workbook is already open in Excel with unsaved changes, do NOT run this on it; build a new file name
# instead. When another Excel session is running, this script never quits it.
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [string]$Cells = '',
    [switch]$NoSave
)
$Path = (Resolve-Path $Path).Path
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$shared = ($xl.Workbooks.Count -gt 0)
try {
    $wb = $xl.Workbooks.Open($Path)
    $xl.CalculateFull()
    $errs = New-Object System.Collections.Generic.List[string]
    $total = 0
    foreach ($ws in $wb.Worksheets) {
        $used = $ws.UsedRange
        $n = 0
        try { $n = $used.SpecialCells(-4123).Count } catch {}
        $total += $n
        $vals = $used.Value2
        if ($vals -is [array]) {
            $rows = $used.Rows.Count; $cols = $used.Columns.Count
            for ($r = 1; $r -le $rows; $r++) {
                for ($c = 1; $c -le $cols; $c++) {
                    $v = $vals[$r, $c]
                    if ($v -is [int] -and $v -lt -2146826200 -and $v -gt -2146826300) { $errs.Add("$($ws.Name)!R${r}C${c}") }
                }
            }
        }
        "{0}: {1} formulas" -f $ws.Name, $n
    }
    "total formulas: $total; error cells: $($errs.Count)"
    if ($errs.Count) { $errs | Select-Object -First 25 }
    foreach ($spec in ($Cells -split '\|' | Where-Object { $_ })) {
        $parts = $spec -split '!', 2
        $rng = $wb.Worksheets.Item($parts[0]).Range($parts[1])
        $vals = @(); foreach ($c in $rng.Cells) { $vals += $c.Text }
        "{0} = {1}" -f $spec, ($vals -join ', ')
    }
    if (-not $NoSave) { $wb.Save() }
    $wb.Close($false)
}
finally {
    if (-not $shared) { $xl.Quit() }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
