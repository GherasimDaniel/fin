param([Parameter(Mandatory=$true)][string]$XlsxPath, [Parameter(Mandatory=$true)][string]$JsonPath)
# Verifies a workbook from build_model.ps1: error scan, history ties, blank inputs, valuation, sensitivity tie-out, scenario table.
[System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]'en-US'
$RW = Get-Content -Raw ($JsonPath -replace '\.json$', '.rowmap.json') | ConvertFrom-Json
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.AutomationSecurity = 3
$fail = 0
try {
  $wb = $xl.Workbooks.Open($XlsxPath, 0, $true); $xl.CalculateFull()
  $nf = 0; $errs = @()
  foreach ($ws in $wb.Worksheets) {
    $ur = $ws.UsedRange; $vals = $ur.Value2
    for ($r = 1; $r -le $ur.Rows.Count; $r++) { for ($c = 1; $c -le $ur.Columns.Count; $c++) {
      if ($ur.Cells($r, $c).HasFormula) { $nf++ }
      $v = $vals[$r, $c]; if ($v -is [int] -and $v -le -2146826200 -and $v -ge -2146826300) { $errs += "$($ws.Name)!$($ur.Cells($r,$c).Address($false,$false)) $($ur.Cells($r,$c).Text)" } } }
  }
  "Formulas: $nf ; error cells: $($errs.Count)"; $errs | Select-Object -First 30; if ($errs.Count) { $fail++ }
  $m = $wb.Worksheets.Item('Model')
  $lastCol = $m.UsedRange.Columns.Count
  $cols = @(); for ($c = 3; $c -le $lastCol; $c++) { $t = $m.Cells(3, $c).Text; if ($t -and $t -ne 'Notes / sources') { $cols += $c } }
  function Show($label, $row, $fmt = '{0,9:N0}') { $line = '{0,-26}' -f $label; foreach ($c in $cols) { $v = $m.Cells($row, $c).Value2; if ($null -eq $v -or $v -is [string]) { $line += '{0,9}' -f '' } else { $line += ($fmt -f $v) } }; $line }
  '{0,-26}' -f '' + (($cols | ForEach-Object { '{0,9}' -f $m.Cells(3, $_).Text }) -join '')
  Show 'Sales' $RW.sales; Show 'Net sales (GAAP)' $RW.gaap; Show 'Gross margin' $RW.gm '{0,9:P1}'; Show 'Operating income' $RW.op; Show 'Net income' $RW.ni; Show 'EPS' $RW.eps '{0,9:N2}'; Show 'EPS reported' $RW.epsrep '{0,9:N2}'
  Show 'Sales y/y' $RW.gSales '{0,9:P1}'; Show 'Op margin' $RW.opm '{0,9:P1}'; Show 'UFCF' $RW.ufcf; Show 'Discounted CF' $RW.dcf
  # EPS check
  $bad = 0; foreach ($c in $cols) { $a = $m.Cells($RW.eps, $c).Value2; $b = $m.Cells($RW.epsrep, $c).Value2; if ($a -is [double] -and $b -is [double] -and [Math]::Abs($a - $b) -gt 0.015) { $bad++; "EPS mismatch col $c : model $a vs reported $b" } }
  if ($bad) { $fail++ } else { 'EPS check: computed EPS matches reported within $0.015' }
  # sales row populated
  $salesVals = @($cols | ForEach-Object { $m.Cells($RW.sales, $_).Value2 } | Where-Object { $_ -is [double] -and $_ -ne 0 })
  if ($salesVals.Count -lt $cols.Count - 2) { "Sales row is empty/zero in $($cols.Count - $salesVals.Count) columns"; $fail++ }
  # every projection column needs its driver inputs (blank inputs silently zero the margins)
  $projCols = @($cols | Where-Object { $m.Cells(3, $_).Text.EndsWith('E') })
  # estimate quarters (labels starting with Q) carry P&L drivers only; DCF drivers live on the annual columns
  foreach ($c in $projCols) { $isQ = $m.Cells(3, $c).Text.StartsWith('Q'); foreach ($k in 'gm','taxR','sgaP','rdP','da','cx','frac') { if ($isQ -and @('da','cx','frac') -contains $k) { continue }; $x = $m.Cells($RW.$k, $c).Value2; if (-not ($x -is [double])) { "Blank projection input: row $k, column $($m.Cells(3, $c).Text)"; $fail++ } } }
  if (-not ($m.Range("C$($RW.vg)").Value2 -is [double])) { 'No valuation computed (value per share is blank)'; $fail++ }
  "`nValuation:"; foreach ($k in 'ev','eqv','vg','vm','nvps','price','tvpct','impm') { '{0,-50} {1}' -f $m.Range("B$($RW.$k)").Text, $m.Range("C$($RW.$k)").Text }
  "`nScenario table:"; for ($r = $RW.tHdr; $r -le $RW.t1 + 2; $r++) { (2..9 | ForEach-Object { $m.Cells($r, $_).Text }) -join ' | ' }
  $w = $wb.Worksheets.Item('WACC')
  $d1 = [Math]::Abs($w.Range('E39').Value2 - $m.Range("C$($RW.vg)").Value2); $d2 = [Math]::Abs($w.Range('E49').Value2 - $m.Range("C$($RW.vm)").Value2)
  "WACC " + $w.Range('C27').Text + " ; sensitivity tie-out diffs: $d1 / $d2"; if ($d1 -gt 0.005 -or $d2 -gt 0.005) { $fail++ }
  $wb.Close($false)
} finally { $xl.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null }
if ($fail) { "VERIFY: FAILED ($fail checks)"; exit 1 } else { 'VERIFY: PASSED' }
