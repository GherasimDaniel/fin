param(
  [Parameter(Mandatory=$true)][string]$Path,
  [int]$MaxRows = 120,
  [int]$MaxCols = 40,
  [string]$OutFile
)
# Dumps every sheet of a workbook: values + formulas (read-only, links/macros disabled)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AskToUpdateLinks = $false
$xl.AutomationSecurity = 3   # msoAutomationSecurityForceDisable
$sb = New-Object System.Text.StringBuilder
try {
  $wb = $xl.Workbooks.Open($Path, 0, $true)
  [void]$sb.AppendLine("WORKBOOK: $Path")
  foreach ($ws in $wb.Worksheets) {
    $ur = $ws.UsedRange
    [void]$sb.AppendLine("=== SHEET: " + $ws.Name + "  used: " + $ur.Address($false,$false) + " ===")
    $c1 = $ws.Cells(1,1)
    [void]$sb.AppendLine("  font: " + $c1.Font.Name + " " + $c1.Font.Size + " | gridlines n/a | zoom n/a")
    $rows = [Math]::Min($ur.Row + $ur.Rows.Count - 1, $MaxRows)
    $cols = [Math]::Min($ur.Column + $ur.Columns.Count - 1, $MaxCols)
    for ($r = 1; $r -le $rows; $r++) {
      $parts = @()
      for ($c = 1; $c -le $cols; $c++) {
        $cell = $ws.Cells($r, $c)
        $f = [string]$cell.Formula
        if ($f -ne "") {
          $addr = $cell.Address($false,$false)
          $t = [string]$cell.Text
          if ($f.StartsWith("=")) { $parts += "$addr{$t|$f}" } else { $parts += "$addr{$t}" }
        }
      }
      if ($parts.Count -gt 0) { [void]$sb.AppendLine("R${r}: " + ($parts -join "  ")) }
    }
  }
  $wb.Close($false)
} finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
if ($OutFile) { [System.IO.File]::WriteAllText($OutFile, $sb.ToString(), [System.Text.Encoding]::UTF8); "wrote $OutFile (" + $sb.Length + " chars)" } else { $sb.ToString() }
