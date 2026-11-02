param([Parameter(Mandatory=$true)][string]$JsonPath, [string]$OutPath)
# Generic Shkreli-style stock model builder: <company>.json -> <TICKER>_Model.xlsx (Excel COM).
# Schema: see ../SKILL.md and examples/glw.json.
$ErrorActionPreference = 'Stop'
[System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]'en-US'
$CFG = Get-Content -Raw -Encoding UTF8 $JsonPath | ConvertFrom-Json
if (-not $OutPath) { $OutPath = $CFG.outFile }

# ---------------- helpers ----------------
function Inv([double]$x) { return $x.ToString([System.Globalization.CultureInfo]::InvariantCulture) }
function SetP($o, [string]$n, $v) { $bv = $null; if ($null -ne $v) { $bv = $v.PSObject.BaseObject }; [void][System.__ComObject].InvokeMember($n, [System.Reflection.BindingFlags]::SetProperty, $null, $o, [object[]]@($bv)) }
function BGR([string]$hex) { $h = $hex.TrimStart('#'); return [Convert]::ToInt32($h.Substring(0,2),16) + 256*[Convert]::ToInt32($h.Substring(2,2),16) + 65536*[Convert]::ToInt32($h.Substring(4,2),16) }
function ColL([int]$n) { $s = ''; while ($n -gt 0) { $m = ($n - 1) % 26; $s = [char](65 + $m) + $s; $n = [int][Math]::Floor(($n - 1) / 26) }; return $s }
function ToDate([string]$s) { return ([datetime]::ParseExact($s, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)).ToOADate() }
$BLUE = BGR '0000FF'; $GREEN = BGR '008000'; $YELLOW = BGR 'FFFF00'; $GRAYF = BGR 'F2F2F2'; $SECT = BGR 'D9E1F2'; $DCFF = BGR 'DDEBF7'; $WHITE = BGR 'FFFFFF'; $NOTE = BGR '595959'
$fUSD = '$#,##0;($#,##0);"-"'; $fEPS = '$0.00;($0.00);"-"'; $fPCT = '0.0%;(0.0%);"-"'; $fNUM = '#,##0;(#,##0);"-"'; $fMULT = '0.0"x"'
function Put($ws, [string]$addr, $v, [string]$kind = '') {
  if ($null -eq $v) { return }
  $c = $ws.Range($addr)
  if ($v -is [string] -and $v.StartsWith('=')) { SetP $c 'Formula' $v } else { SetP $c 'Value2' $v }
  if ($kind -eq 'in') { $c.Font.Color = $BLUE } elseif ($kind -eq 'lk') { $c.Font.Color = $GREEN }
}
function FR($ws, [string]$rng, [string]$f, [string]$kind = '') { $r = $ws.Range($rng); SetP $r 'Formula' $f; if ($kind -eq 'in') { $r.Font.Color = $BLUE } elseif ($kind -eq 'lk') { $r.Font.Color = $GREEN } }
function Txt($ws, [string]$addr, [string]$t) { $c = $ws.Range($addr); $c.NumberFormat = '@'; SetP $c 'Value2' $t }
function Lbl($ws, [string]$addr, [string]$t) { SetP ($ws.Range($addr)) 'Value2' $t }
function Sect($ws, [string]$rng, [string]$title) { $r = $ws.Range($rng); $r.Interior.Color = $SECT; $r.Font.Bold = $true; Lbl $ws ($rng.Split(':')[0]) $title }
function Note($ws, [string]$addr, [string]$t) { if (-not $t) { return }; $c = $ws.Range($addr); SetP $c 'Value2' $t; $c.Font.Italic = $true; $c.Font.Color = $NOTE }
function CW($ws, [string]$col, [double]$w) { SetP ($ws.Columns.Item($col)) 'ColumnWidth' $w }
function Arr($a) { if ($null -eq $a) { return ,@() } ; return ,@($a) }   # comma: keep 1-item arrays as arrays
function At($a, [int]$i) { $x = Arr $a; if ($i -lt $x.Count) { return $x[$i] } ; return $null }

# ---------------- periods / columns ----------------
$Qs = Arr $CFG.periods.quarters; $Ys = Arr $CFG.periods.years; $Ps = Arr $CFG.periods.proj
$nQ = $Qs.Count; $nY = $Ys.Count; $nP = $Ps.Count
$qc = @(); for ($i = 0; $i -lt $nQ; $i++) { $qc += ColL (3 + $i) }
$yStart = 4 + $nQ
$hc = @(); for ($i = 0; $i -lt $nY; $i++) { $hc += ColL ($yStart + $i) }
$pc = @(); for ($i = 0; $i -lt $nP; $i++) { $pc += ColL ($yStart + $nY + $i) }
$ac = $hc + $pc
$noteCol = ColL ($yStart + $nY + $nP + 1)
$firstCol = 'C'; $lastCol = $ac[-1]
$qAct = @(); $qEst = @(); for ($i = 0; $i -lt $nQ; $i++) { if ($Qs[$i].est) { $qEst += $i } else { $qAct += $i } }
# which projection years are built from quarters
$qbIdx = @{}
for ($k = 0; $k -lt $nP; $k++) {
  # quarters belonging to this fiscal year: after the previous FY end, up to this FY end (+10 days for 52/53-week years)
  $pe = if ($k -eq 0) { ToDate $Ys[-1].end } else { ToDate $Ps[$k-1].end }; $ce = ToDate $Ps[$k].end
  $ix = @(); for ($i = 0; $i -lt $nQ; $i++) { $qe = ToDate $Qs[$i].end; if ($qe -gt $pe + 10 -and $qe -le $ce + 10) { $ix += $i } }
  if ($ix.Count -eq 4) { $qbIdx[$k] = $ix } }
$scenYears = @(); for ($k = 0; $k -lt $nP; $k++) { if (-not $qbIdx.ContainsKey($k)) { $scenYears += $k } }
$segs = Arr $CFG.segments; $nS = $segs.Count
$lastH = $hc[-1]; $prevH = $hc[-2]

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
try {
  $wb = $xl.Workbooks.Add(); $xl.Calculation = -4105
  $wb.Styles.Item('Normal').Font.Name = 'Arial'; $wb.Styles.Item('Normal').Font.Size = 10
  while ($wb.Worksheets.Count -lt 6) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
  while ($wb.Worksheets.Count -gt 6) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
  $names = 'Main','Model','Scenarios','WACC','Contracts','Sources'
  for ($i = 1; $i -le 6; $i++) { $wb.Worksheets.Item($i).Name = $names[$i-1] }
  $wsMain = $wb.Worksheets.Item('Main'); $ws = $wb.Worksheets.Item('Model'); $wsS = $wb.Worksheets.Item('Scenarios'); $wsW = $wb.Worksheets.Item('WACC'); $wsC = $wb.Worksheets.Item('Contracts'); $wsSrc = $wb.Worksheets.Item('Sources')
  foreach ($w in @($ws,$wsS,$wsW,$wsC,$wsSrc)) { Lbl $w 'A1' 'Main'; [void]$w.Hyperlinks.Add($w.Range('A1'), '', 'Main!A1', 'Back to Main', 'Main') }
  $basis = $CFG.basis; if (-not $basis) { $basis = 'adjusted' }

  # ---------------- row map ----------------
  $RW = @{}; $r = 4
  $RW.secSeg = $r; $r++
  $RW.seg = @(); for ($s = 0; $s -lt $nS; $s++) { $RW.seg += $r; $r++ }
  $RW.sales = $r; $RW.adj = $r+1; $RW.gaap = $r+2; $r += 4
  $RW.secPL = $r; $r++
  foreach ($k in 'cogs','gp','sga','rd','op','other','pretax','tax','nci','ni','eps','sh','epsrep','gop','gni','geps') { $RW[$k] = $r; $r++ }
  $r++; $RW.secG = $r; $r++
  $RW.gSales = $r; $r++
  $RW.segG = @(); for ($s = 0; $s -lt $nS; $s++) { $RW.segG += $r; $r++ }
  foreach ($k in 'gm','sgaP','rdP','opm','taxR','nciP','shG','epsG','adjP') { $RW[$k] = $r; $r++ }
  $r++; $RW.secBS = $r; $r++
  foreach ($k in 'cash','ar','inv','ppe','assets','ap','debt','bsnci','eq','netcash','nwc') { $RW[$k] = $r; $r++ }
  $r++; $RW.secCF = $r; $r++
  foreach ($k in 'ocf','capex','hedge','afcf','dep','amort','sbc','div','buyback') { $RW[$k] = $r; $r++ }
  $r++; $RW.secDrv = $r; $r++
  foreach ($k in 'da','cx','nwcP','nonc','frac') { $RW[$k] = $r; $r++ }
  $r++; $RW.secDCF = $r; $r++
  foreach ($k in 'ebit','ncc','aebit','etax','nopat','pda','pcx','pnwc','ufcf','ufcfc','per','df','dcf') { $RW[$k] = $r; $r++ }
  $r++; $RW.secV = $r; $r++
  foreach ($k in 'sumpv','g','mult','ebitda','tvg','tvm','tvmeth','tvused','yrs','pvtv','ev','ldebt','pcash','lnci','eqv','shs','vps','price','upside','tvpct','impm','impg','vg','vm') { $RW[$k] = $r; $r++ }
  $r++; $RW.secSk = $r; $r++
  foreach ($k in 'disc','mat','npv','nvps','nup') { $RW[$k] = $r; $r++ }
  $r++; $RW.secT = $r; $RW.tHdr = $r+1; $RW.tCur = $r+2; $RW.t1 = $r+3

  $selCol = $lastH; $selCell = "`$$selCol`$1"; $sel = "$selCol" + '1'
  $nameCol = $pc[0]
  # ---------------- header ----------------
  Lbl $ws 'B1' ("{0} ({1}) - Model ($ in millions, except per share). '{2}' = company non-GAAP basis." -f $CFG.company, $CFG.ticker, $basis); $ws.Range('B1').Font.Bold = $true
  Lbl $ws "$prevH`1" 'Scenario:'; $ws.Range("$prevH`1").HorizontalAlignment = -4152
  $scen = $CFG.scenario; if (-not $scen) { $scen = 2 }
  Put $ws $sel ([double]$scen) 'in'; $ws.Range($sel).Interior.Color = $YELLOW; $ws.Range($sel).HorizontalAlignment = -4108
  Put $ws "$nameCol`1" "=CHOOSE($sel,""Bear"",""Base"",""Bull"")"; $ws.Range("$nameCol`1").Font.Bold = $true
  [void]$ws.Range($sel).AddComment('Scenario selector for annual projections (Scenarios tab): 1 = Bear, 2 = Base, 3 = Bull')
  for ($i = 0; $i -lt $nQ; $i++) { Put $ws "$($qc[$i])2" (ToDate $Qs[$i].end); Txt $ws "$($qc[$i])3" $Qs[$i].label }
  for ($i = 0; $i -lt $nY; $i++) { Put $ws "$($hc[$i])2" (ToDate $Ys[$i].end); Txt $ws "$($hc[$i])3" $Ys[$i].label }
  for ($i = 0; $i -lt $nP; $i++) { Put $ws "$($pc[$i])2" (ToDate $Ps[$i].end); Txt $ws "$($pc[$i])3" $Ps[$i].label }
  $ws.Range("C2:$lastCol`2").NumberFormat = 'm/d/yy'; $ws.Range("C2:$lastCol`3").HorizontalAlignment = -4152; $ws.Range("C3:$lastCol`3").Font.Bold = $true
  foreach ($c in (@($qEst | ForEach-Object { $qc[$_] }) + $pc)) { $ws.Range("${c}2:${c}3").Interior.Color = $GRAYF; $ws.Range("${c}3").Font.Italic = $true }
  Lbl $ws "$noteCol`3" 'Notes / sources'; $ws.Range("$noteCol`3").Font.Bold = $true

  # convenience: write a data row (quarters actual + history years)
  function DataRow([int]$row, $qArr, $yArr) {
    foreach ($i in $qAct) { Put $ws "$($qc[$i])$row" (At $qArr $i) 'in' }
    for ($i = 0; $i -lt $nY; $i++) { Put $ws "$($hc[$i])$row" (At $yArr $i) 'in' }
  }
  $QD = $CFG.q; $YD = $CFG.y; $QE = $CFG.qEst; $PD = $CFG.proj

  # ---------------- segments ----------------
  Sect $ws "B$($RW.secSeg):$lastCol$($RW.secSeg)" "Sales by segment - $basis basis (`$mm)"
  for ($s = 0; $s -lt $nS; $s++) {
    $sg = $segs[$s]; $row = $RW.seg[$s]; $gr = $RW.segG[$s]
    Lbl $ws "B$row" $sg.name; $ws.Range("B$row").IndentLevel = 1
    DataRow $row $sg.q $sg.y
    foreach ($i in $qEst) { Put $ws "$($qc[$i])$row" "=$($qc[$i-4])$row*(1+$($qc[$i])$gr)" }
    for ($k = 0; $k -lt $nP; $k++) {
      $c = $pc[$k]; $prev = $ac[$nY + $k - 1]
      if ($qbIdx.ContainsKey($k)) { $ix = $qbIdx[$k]; Put $ws "$c$row" "=SUM($($qc[$ix[0]])${row}:$($qc[$ix[3]])$row)" } else { Put $ws "$c$row" "=$prev$row*(1+$c$gr)" }
    }
    Note $ws "$noteCol$row" $sg.note
  }
  $s0 = $RW.seg[0]; $sN = $RW.seg[-1]
  foreach ($c in ($qc + $ac)) { Put $ws "$c$($RW.sales)" "=SUM($c${s0}:$c$sN)" }
  Lbl $ws "B$($RW.sales)" ("{0} sales" -f (Get-Culture).TextInfo.ToTitleCase($basis)); Lbl $ws "B$($RW.adj)" $(if ($CFG.adjLabel) { $CFG.adjLabel } else { 'Adjustment to GAAP sales' }); Lbl $ws "B$($RW.gaap)" 'Net sales (GAAP)'
  DataRow $RW.gaap $QD.gaapSales $YD.gaapSales
  foreach ($c in (@($qAct | ForEach-Object { $qc[$_] }) + $hc)) { Put $ws "$c$($RW.adj)" "=IF($c$($RW.gaap)="""","""",$c$($RW.gaap)-$c$($RW.sales))" }
  foreach ($i in $qEst) { $c = $qc[$i]; Put $ws "$c$($RW.adj)" "=$c$($RW.sales)*$c$($RW.adjP)"; Put $ws "$c$($RW.gaap)" "=$c$($RW.sales)+$c$($RW.adj)" }

  # ---------------- P&L ----------------
  Sect $ws "B$($RW.secPL):$lastCol$($RW.secPL)" "$((Get-Culture).TextInfo.ToTitleCase($basis)) P&L (`$mm)"
  $lab = @{cogs='Cost of sales';gp='Gross profit';sga='SG&A';rd='R&D';op='Operating income';other='Other income (expense), net';pretax='Pretax income';tax='Taxes';nci='Non-controlling interest';ni='Net income';eps="EPS ($basis, diluted)";sh='Diluted shares (mm)';epsrep='EPS - company reported (check)';gop='GAAP operating income';gni='GAAP net income';geps='GAAP EPS (diluted)'}
  foreach ($k in $lab.Keys) { Lbl $ws "B$($RW[$k])" $lab[$k] }
  DataRow $RW.gp $QD.gp $YD.gp; DataRow $RW.sga $QD.sga $YD.sga; DataRow $RW.rd $QD.rd $YD.rd; DataRow $RW.pretax $QD.pretax $YD.pretax
  DataRow $RW.ni $QD.ni $YD.ni; DataRow $RW.sh $QD.shares $YD.shares; DataRow $RW.epsrep $QD.epsRep $YD.epsRep
  DataRow $RW.gop $QD.gaapOpInc $YD.gaapOpInc; DataRow $RW.gni $QD.gaapNi $YD.gaapNi; DataRow $RW.geps $QD.gaapEps $YD.gaapEps
  DataRow $RW.taxR $QD.taxRate $YD.taxRate
  $histCols = @($qAct | ForEach-Object { $qc[$_] }) + $hc
  foreach ($c in $histCols) {
    Put $ws "$c$($RW.other)" "=$c$($RW.pretax)-$c$($RW.op)"
    Put $ws "$c$($RW.tax)" "=$c$($RW.pretax)*$c$($RW.taxR)"
    Put $ws "$c$($RW.nci)" "=$c$($RW.pretax)-$c$($RW.tax)-$c$($RW.ni)"
  }
  foreach ($c in ($qc + $ac)) {
    Put $ws "$c$($RW.cogs)" "=$c$($RW.sales)-$c$($RW.gp)"
    Put $ws "$c$($RW.op)" "=$c$($RW.gp)-$c$($RW.sga)-$c$($RW.rd)"
    Put $ws "$c$($RW.eps)" "=IFERROR($c$($RW.ni)/$c$($RW.sh),"""")"
  }
  # estimated quarters
  foreach ($i in $qEst) {
    $c = $qc[$i]; $j = [array]::IndexOf($qEst, $i)
    Put $ws "$c$($RW.gp)" "=$c$($RW.sales)*$c$($RW.gm)"; Put $ws "$c$($RW.sga)" "=$c$($RW.sales)*$c$($RW.sgaP)"; Put $ws "$c$($RW.rd)" "=$c$($RW.sales)*$c$($RW.rdP)"
    Put $ws "$c$($RW.other)" (At $QE.other $j) 'in'
    Put $ws "$c$($RW.pretax)" "=$c$($RW.op)+$c$($RW.other)"; Put $ws "$c$($RW.tax)" "=$c$($RW.pretax)*$c$($RW.taxR)"; Put $ws "$c$($RW.nci)" "=$c$($RW.pretax)*$c$($RW.nciP)"
    Put $ws "$c$($RW.ni)" "=$c$($RW.pretax)-$c$($RW.tax)-$c$($RW.nci)"; Put $ws "$c$($RW.sh)" "=$($qc[$i-1])$($RW.sh)*(1+$c$($RW.shG))"
    for ($s = 0; $s -lt $nS; $s++) { Put $ws "$c$($RW.segG[$s])" (At $segs[$s].qEstGrowth $j) 'in' }
    Put $ws "$c$($RW.gm)" (At $QE.gm $j) 'in'; Put $ws "$c$($RW.sgaP)" (At $QE.sga $j) 'in'; Put $ws "$c$($RW.rdP)" (At $QE.rd $j) 'in'
    Put $ws "$c$($RW.taxR)" (At $QE.tax $j) 'in'; Put $ws "$c$($RW.nciP)" (At $QE.nci $j) 'in'; Put $ws "$c$($RW.shG)" (At $QE.shareGrowth $j) 'in'; Put $ws "$c$($RW.adjP)" (At $QE.adjPct $j) 'in'
  }
  # projection years
  for ($k = 0; $k -lt $nP; $k++) {
    $c = $pc[$k]; $prev = $ac[$nY + $k - 1]
    if ($qbIdx.ContainsKey($k)) {
      $a = $qc[$qbIdx[$k][0]]; $b = $qc[$qbIdx[$k][3]]
      foreach ($key in 'gaap','adj','gp','sga','rd','other','pretax','tax','nci','ni') { $rr = $RW[$key]; Put $ws "$c$rr" "=SUM($a${rr}:$b$rr)" }
      Put $ws "$c$($RW.sh)" "=AVERAGE($a$($RW.sh):$b$($RW.sh))"
    } else {
      Put $ws "$c$($RW.adj)" "=$c$($RW.sales)*$c$($RW.adjP)"; Put $ws "$c$($RW.gaap)" "=$c$($RW.sales)+$c$($RW.adj)"
      Put $ws "$c$($RW.gp)" "=$c$($RW.sales)*$c$($RW.gm)"; Put $ws "$c$($RW.sga)" "=$c$($RW.sales)*$c$($RW.sgaP)"; Put $ws "$c$($RW.rd)" "=$c$($RW.sales)*$c$($RW.rdP)"
      Put $ws "$c$($RW.other)" (At $PD.other $k) 'in'
      Put $ws "$c$($RW.pretax)" "=$c$($RW.op)+$c$($RW.other)"; Put $ws "$c$($RW.tax)" "=$c$($RW.pretax)*$c$($RW.taxR)"; Put $ws "$c$($RW.nci)" "=$c$($RW.pretax)*$c$($RW.nciP)"
      Put $ws "$c$($RW.ni)" "=$c$($RW.pretax)-$c$($RW.tax)-$c$($RW.nci)"; Put $ws "$c$($RW.sh)" "=$prev$($RW.sh)*(1+$c$($RW.shG))"
      Put $ws "$c$($RW.sgaP)" (At $PD.sga $k) 'in'; Put $ws "$c$($RW.rdP)" (At $PD.rd $k) 'in'; Put $ws "$c$($RW.taxR)" (At $PD.tax $k) 'in'
      Put $ws "$c$($RW.nciP)" (At $PD.nci $k) 'in'; Put $ws "$c$($RW.shG)" (At $PD.shareGrowth $k) 'in'; Put $ws "$c$($RW.adjP)" (At $PD.adjPct $k) 'in'
    }
  }

  # ---------------- growth & margins ----------------
  Sect $ws "B$($RW.secG):$lastCol$($RW.secG)" 'Growth & margins (blue = input; green = from Scenarios tab)'
  $lab = @{gSales='Sales y/y';gm='Gross margin';sgaP='SG&A % of sales';rdP='R&D % of sales';opm='Operating margin';taxR="Tax rate ($basis)";nciP='NCI % of pretax';shG='Diluted share growth';epsG='EPS y/y';adjP='GAAP adjustment % of sales'}
  foreach ($k in $lab.Keys) { Lbl $ws "B$($RW[$k])" $lab[$k] }
  for ($s = 0; $s -lt $nS; $s++) { Lbl $ws "B$($RW.segG[$s])" ("{0} y/y" -f $segs[$s].name); $ws.Range("B$($RW.segG[$s])").IndentLevel = 1 }
  $yoy = @(@{row=$RW.gSales; src=$RW.sales}, @{row=$RW.epsG; src=$RW.eps})
  for ($s = 0; $s -lt $nS; $s++) { $yoy += @{row=$RW.segG[$s]; src=$RW.seg[$s]; seg=$segs[$s]} }
  foreach ($y in $yoy) {
    foreach ($i in $qAct) { if ($i -ge 4) { Put $ws "$($qc[$i])$($y.row)" "=IFERROR($($qc[$i])$($y.src)/$($qc[$i-4])$($y.src)-1,"""")" } }
    if (-not $y.seg) { foreach ($i in $qEst) { Put $ws "$($qc[$i])$($y.row)" "=IFERROR($($qc[$i])$($y.src)/$($qc[$i-4])$($y.src)-1,"""")" } }
    for ($i = 1; $i -lt $nY + $nP; $i++) {
      $c = $ac[$i]; $k = $i - $nY
      if ($y.seg -and $k -ge 0 -and -not $qbIdx.ContainsKey($k)) { continue }
      if ($y.seg -and $k -lt 0 -and (Arr $y.seg.noYoY) -contains $Ys[$i].label) { continue }
      Put $ws "$c$($y.row)" "=IFERROR($c$($y.src)/$($ac[$i-1])$($y.src)-1,"""")"
    }
  }
  foreach ($c in ($histCols + @($pc | Where-Object { $qbIdx.ContainsKey([array]::IndexOf($pc, $_)) }))) {
    Put $ws "$c$($RW.gm)" "=IFERROR($c$($RW.gp)/$c$($RW.sales),"""")"; Put $ws "$c$($RW.sgaP)" "=IFERROR($c$($RW.sga)/$c$($RW.sales),"""")"; Put $ws "$c$($RW.rdP)" "=IFERROR($c$($RW.rd)/$c$($RW.sales),"""")"
    Put $ws "$c$($RW.nciP)" "=IFERROR($c$($RW.nci)/$c$($RW.pretax),"""")"; Put $ws "$c$($RW.adjP)" "=IFERROR($c$($RW.adj)/$c$($RW.sales),"""")"
  }
  foreach ($k in $qbIdx.Keys) { $c = $pc[$k]; Put $ws "$c$($RW.taxR)" "=IFERROR($c$($RW.tax)/$c$($RW.pretax),"""")"; Put $ws "$c$($RW.shG)" "=IFERROR($c$($RW.sh)/$($ac[$nY+$k-1])$($RW.sh)-1,"""")" }
  foreach ($i in $qAct) { if ($i -ge 1) { Put $ws "$($qc[$i])$($RW.shG)" "=IFERROR($($qc[$i])$($RW.sh)/$($qc[$i-1])$($RW.sh)-1,"""")" } }
  for ($i = 1; $i -lt $nY; $i++) { Put $ws "$($hc[$i])$($RW.shG)" "=IFERROR($($hc[$i])$($RW.sh)/$($hc[$i-1])$($RW.sh)-1,"""")" }
  foreach ($c in ($qc + $ac)) { Put $ws "$c$($RW.opm)" "=IFERROR($c$($RW.op)/$c$($RW.sales),"""")" }
  # scenario links
  $scRow = @{}
  $sr = 5
  for ($s = 0; $s -lt $nS; $s++) { $scRow["seg$s"] = $sr; $sr += 4 }
  $scRow['gm'] = $sr
  for ($j = 0; $j -lt $scenYears.Count; $j++) {
    $c = $pc[$scenYears[$j]]; $scCol = ColL (3 + $j)
    for ($s = 0; $s -lt $nS; $s++) { $b = $scRow["seg$s"]; Put $ws "$c$($RW.segG[$s])" ("=CHOOSE($selCell,Scenarios!$scCol{0},Scenarios!$scCol{1},Scenarios!$scCol{2})" -f ($b+1),($b+2),($b+3)) 'lk' }
    $b = $scRow['gm']; Put $ws "$c$($RW.gm)" ("=CHOOSE($selCell,Scenarios!$scCol{0},Scenarios!$scCol{1},Scenarios!$scCol{2})" -f ($b+1),($b+2),($b+3)) 'lk'
  }

  # ---------------- balance sheet & cash flow ----------------
  Sect $ws "B$($RW.secBS):$lastCol$($RW.secBS)" 'Balance sheet ($mm, period end)'
  $lab = @{cash='Cash & equivalents';ar='Accounts receivable';inv='Inventories';ppe='PP&E, net';assets='Total assets';ap='Accounts payable';debt='Total debt';bsnci='Non-controlling interest';eq="Shareholders' equity";netcash='Net cash / (net debt)';nwc='Net working capital (AR + inv. - AP)'}
  foreach ($k in $lab.Keys) { Lbl $ws "B$($RW[$k])" $lab[$k] }
  foreach ($k in 'cash','ar','inv','ppe','assets','ap','debt') { DataRow $RW[$k] $QD.$k $YD.$k }
  DataRow $RW.bsnci $QD.nciBS $YD.nciBS; DataRow $RW.eq $QD.equity $YD.equity
  foreach ($c in $histCols) { Put $ws "$c$($RW.netcash)" "=IF($c$($RW.cash)="""","""",$c$($RW.cash)-$c$($RW.debt))"; Put $ws "$c$($RW.nwc)" "=IF($c$($RW.ar)="""","""",$c$($RW.ar)+$c$($RW.inv)-$c$($RW.ap))" }
  Sect $ws "B$($RW.secCF):$lastCol$($RW.secCF)" 'Cash flow ($mm)'
  $lab = @{ocf='Operating cash flow (GAAP)';capex='Capital expenditures';hedge='Other FCF items (e.g. realized hedge gains)';afcf='Free cash flow (OCF + other - capex)';dep='Depreciation';amort='Amortization of intangibles';sbc='Stock-based compensation';div='Dividends paid';buyback='Share repurchases'}
  foreach ($k in $lab.Keys) { Lbl $ws "B$($RW[$k])" $lab[$k] }
  foreach ($k in 'ocf','capex','hedge','dep','amort','sbc','div','buyback') { DataRow $RW[$k] $QD.$k $YD.$k }
  foreach ($c in $histCols) { Put $ws "$c$($RW.afcf)" "=IF($c$($RW.ocf)="""","""",$c$($RW.ocf)+$c$($RW.capex)+N($c$($RW.hedge)))" }

  # ---------------- DCF ----------------
  Sect $ws "B$($RW.secDrv):$lastCol$($RW.secDrv)" 'DCF drivers (history = actual ratios; projections = inputs)'
  $lab = @{da='D&A (depreciation) % of sales';cx='Capex % of sales';nwcP='NWC % of incremental sales';nonc='Non-core cash costs % of sales';frac='Fraction of year in valuation'}
  foreach ($k in $lab.Keys) { Lbl $ws "B$($RW[$k])" $lab[$k] }
  for ($i = 0; $i -lt $nY; $i++) {
    $c = $hc[$i]
    Put $ws "$c$($RW.da)" "=IFERROR($c$($RW.dep)/$c$($RW.sales),"""")"; Put $ws "$c$($RW.cx)" "=IFERROR(-$c$($RW.capex)/$c$($RW.sales),"""")"
    if ($i -ge 1) { Put $ws "$c$($RW.nwcP)" "=IFERROR(($c$($RW.nwc)-$($hc[$i-1])$($RW.nwc))/($c$($RW.sales)-$($hc[$i-1])$($RW.sales)),"""")" }
    Put $ws "$c$($RW.nonc)" "=IFERROR(($c$($RW.op)-$c$($RW.gop)-N($c$($RW.amort))-N($c$($RW.hedge)))/$c$($RW.sales),"""")"
  }
  for ($k = 0; $k -lt $nP; $k++) { $c = $pc[$k]; Put $ws "$c$($RW.da)" (At $PD.da $k) 'in'; Put $ws "$c$($RW.cx)" (At $PD.capex $k) 'in'; Put $ws "$c$($RW.nwcP)" (At $PD.nwc $k) 'in'; Put $ws "$c$($RW.nonc)" (At $PD.noncore $k) 'in'; Put $ws "$c$($RW.frac)" (At $PD.fraction $k) 'in' }
  Sect $ws "B$($RW.secDCF):$lastCol$($RW.secDCF)" 'Discounted cash flow ($mm) - unlevered, valuation date on WACC tab'
  $lab = @{ebit="$basis EBIT (operating income)";ncc='Less: non-core cash costs';aebit='Adjusted EBIT (cash basis)';etax='Less: taxes on EBIT';nopat='NOPAT';pda='Plus: D&A';pcx='Less: capex';pnwc='Less: increase in NWC';ufcf='Unlevered free cash flow';ufcfc='UFCF counted in valuation';per='Discount period (years, mid-period)';df='Discount factor';dcf='Discounted Cash Flow (PV of UFCF)'}
  foreach ($k in $lab.Keys) { Lbl $ws "B$($RW[$k])" $lab[$k] }
  for ($k = 0; $k -lt $nP; $k++) {
    $c = $pc[$k]; $prev = $ac[$nY + $k - 1]
    Put $ws "$c$($RW.ebit)" "=$c$($RW.op)"; Put $ws "$c$($RW.ncc)" "=-$c$($RW.sales)*$c$($RW.nonc)"; Put $ws "$c$($RW.aebit)" "=$c$($RW.ebit)+$c$($RW.ncc)"
    Put $ws "$c$($RW.etax)" "=-$c$($RW.aebit)*$c$($RW.taxR)"; Put $ws "$c$($RW.nopat)" "=$c$($RW.aebit)+$c$($RW.etax)"
    Put $ws "$c$($RW.pda)" "=$c$($RW.sales)*$c$($RW.da)"; Put $ws "$c$($RW.pcx)" "=-$c$($RW.sales)*$c$($RW.cx)"; Put $ws "$c$($RW.pnwc)" "=-($c$($RW.sales)-$prev$($RW.sales))*$c$($RW.nwcP)"
    Put $ws "$c$($RW.ufcf)" "=$c$($RW.nopat)+$c$($RW.pda)+$c$($RW.pcx)+$c$($RW.pnwc)"; Put $ws "$c$($RW.ufcfc)" "=$c$($RW.ufcf)*$c$($RW.frac)"
    Put $ws "$c$($RW.per)" "=MAX(0,($c`$2-$c$($RW.frac)*365.25/2-WACC!`$C`$32)/365.25)"; Put $ws "$c$($RW.df)" "=1/(1+WACC!`$C`$27)^$c$($RW.per)"; Put $ws "$c$($RW.dcf)" "=$c$($RW.ufcfc)*$c$($RW.df)"
  }
  $ws.Range("B$($RW.dcf):$lastCol$($RW.dcf)").Interior.Color = $DCFF; $ws.Range("B$($RW.dcf):$lastCol$($RW.dcf)").Font.Bold = $true
  $p0 = $pc[0]; $pN = $pc[-1]
  # ---------------- valuation ----------------
  Sect $ws "B$($RW.secV):$lastCol$($RW.secV)" 'Valuation'
  $V = [ordered]@{
    sumpv=@("Sum of PV of FCF ($($Ps[0].label) to $($Ps[-1].label))","=SUM($p0$($RW.dcf):$pN$($RW.dcf))",$fUSD,'')
    g=@('Terminal growth (g)','=WACC!C30',$fPCT,'lk'); mult=@('Exit EV/EBITDA multiple','=WACC!C31',$fMULT,'lk')
    ebitda=@("$($Ps[-1].label) EBITDA (EBIT + D&A)","=$pN$($RW.ebit)+$pN$($RW.pda)",$fUSD,'')
    tvg=@('Terminal value - Gordon growth',"=$pN$($RW.ufcf)*(1+C$($RW.g))/(WACC!C27-C$($RW.g))",$fUSD,''); tvm=@('Terminal value - exit multiple',"=C$($RW.ebitda)*C$($RW.mult)",$fUSD,'')
    tvmeth=@('TV method used (1 = Gordon, 2 = exit multiple)',[double]$(if ($CFG.tvMethod) { $CFG.tvMethod } else { 1 }),'0','in'); tvused=@('Terminal value used',"=CHOOSE(C$($RW.tvmeth),C$($RW.tvg),C$($RW.tvm))",$fUSD,'')
    yrs=@('Years to terminal',"=($pN`2-WACC!C32)/365.25",'0.00',''); pvtv=@('PV of terminal value',"=C$($RW.tvused)/(1+WACC!C27)^C$($RW.yrs)",$fUSD,'')
    ev=@('Enterprise value',"=C$($RW.sumpv)+C$($RW.pvtv)",$fUSD,''); ldebt=@('Less: total debt','=-Main!I6',$fUSD,'lk'); pcash=@('Plus: cash','=Main!I5',$fUSD,'lk'); lnci=@('Less: non-controlling interest (book)','=-Main!I8',$fUSD,'lk')
    eqv=@('Equity value',"=SUM(C$($RW.ev):C$($RW.lnci))",$fUSD,''); shs=@('Diluted shares (mm)','=Main!I3',$fNUM,'lk')
    vps=@('Value per share (selected TV method)',"=C$($RW.eqv)/C$($RW.shs)",$fEPS,''); price=@('Current share price','=Main!I2',$fEPS,'lk'); upside=@('Upside / (downside)',"=C$($RW.vps)/C$($RW.price)-1",$fPCT,'')
    tvpct=@('Terminal value as % of EV',"=C$($RW.pvtv)/C$($RW.ev)",$fPCT,''); impm=@('Implied exit multiple of Gordon TV',"=C$($RW.tvg)/C$($RW.ebitda)",$fMULT,'')
    impg=@('Implied perpetual growth in exit-multiple TV',"=(C$($RW.tvm)*WACC!C27-$pN$($RW.ufcf))/(C$($RW.tvm)+$pN$($RW.ufcf))",$fPCT,'')
    vg=@('Value per share - Gordon growth',"=(C$($RW.sumpv)+C$($RW.tvg)/(1+WACC!C27)^C$($RW.yrs)+C$($RW.ldebt)+C$($RW.pcash)+C$($RW.lnci))/C$($RW.shs)",$fEPS,'')
    vm=@('Value per share - exit multiple',"=(C$($RW.sumpv)+C$($RW.tvm)/(1+WACC!C27)^C$($RW.yrs)+C$($RW.ldebt)+C$($RW.pcash)+C$($RW.lnci))/C$($RW.shs)",$fEPS,'')
  }
  foreach ($e in $V.GetEnumerator()) { $row = $RW[$e.Key]; $a = $e.Value; Lbl $ws "B$row" $a[0]; Put $ws "C$row" $a[1] $a[3]; $ws.Range("C$row").NumberFormat = $a[2] }
  $ws.Range("C$($RW.tvmeth)").Interior.Color = $YELLOW
  foreach ($k in 'ev','eqv','vps','vg','vm') { $ws.Range("B$($RW[$k]):C$($RW[$k])").Font.Bold = $true }
  $ws.Range("B$($RW.vps):C$($RW.vps)").Interior.Color = $DCFF
  Note $ws "D$($RW.tvg)" 'Gordon growth: final-year UFCF x (1+g) / (WACC - g).'
  Note $ws "D$($RW.tvmeth)" 'Switch to 2 to use the exit multiple as the headline value.'
  Note $ws "D$($RW.impm)" 'Gordon TV expressed as an EV/EBITDA multiple - compare with the exit multiple.'
  Sect $ws "B$($RW.secSk):$lastCol$($RW.secSk)" 'Shkreli-style cross-check: NPV of net income ("Discount" / "Maturity" method)'
  $mat = $CFG.wacc.maturity; if ($null -eq $mat) { $mat = $CFG.wacc.g }
  $V2 = [ordered]@{
    disc=@('Discount (cost of equity)','=WACC!C16',$fPCT,'lk'); mat=@('Maturity growth (after final year)',[double]$mat,$fPCT,'in')
    npv=@('NPV of net income ($mm)',"=SUMPRODUCT($p0$($RW.ni):$pN$($RW.ni),$p0$($RW.frac):$pN$($RW.frac),(1+C$($RW.disc))^(-$p0$($RW.per):$pN$($RW.per)))+$pN$($RW.ni)*(1+C$($RW.mat))/(C$($RW.disc)-C$($RW.mat))/(1+C$($RW.disc))^C$($RW.yrs)",$fUSD,'')
    nvps=@('Value per share',"=C$($RW.npv)/C$($RW.shs)",$fEPS,''); nup=@('Upside / (downside)',"=C$($RW.nvps)/C$($RW.price)-1",$fPCT,'')
  }
  foreach ($e in $V2.GetEnumerator()) { $row = $RW[$e.Key]; $a = $e.Value; Lbl $ws "B$row" $a[0]; Put $ws "C$row" $a[1] $a[3]; $ws.Range("C$row").NumberFormat = $a[2] }
  $ws.Range("B$($RW.nvps):C$($RW.nvps)").Font.Bold = $true
  Note $ws "D$($RW.npv)" 'Net income is an equity flow: discounted at the cost of equity, no net-debt adjustment.'
  # scenario data table
  Sect $ws "B$($RW.secT):$lastCol$($RW.secT)" "Scenario comparison (live Excel data table on selector cell $sel)"
  $y1 = if ($scenYears.Count -gt 0) { $pc[$scenYears[0]] } else { $p0 }
  $hdr = @('Scenario',"$($ws.Range("${y1}3").Text) sales","$($Ps[-1].label) sales","$($ws.Range("${y1}3").Text) EPS","$($Ps[-1].label) EPS",'DCF $/sh (Gordon)','DCF $/sh (exit mult.)','Shkreli NPV $/sh')
  $tc = @('B','C','D','E','F','G','H','I')
  for ($i = 0; $i -lt 8; $i++) { $c = $ws.Range("$($tc[$i])$($RW.tHdr)"); SetP $c 'Value2' $hdr[$i]; $c.Font.Bold = $true; $c.WrapText = $true; $c.HorizontalAlignment = -4108 }
  Lbl $ws "B$($RW.tCur)" 'Selected scenario ->'
  $tf = @("=$y1$($RW.sales)","=$pN$($RW.sales)","=$y1$($RW.eps)","=$pN$($RW.eps)","=C$($RW.vg)","=C$($RW.vm)","=C$($RW.nvps)")
  for ($i = 0; $i -lt 7; $i++) { Put $ws "$($tc[$i+1])$($RW.tCur)" $tf[$i] }
  $scnNames = @('Bear','Base','Bull')
  for ($i = 0; $i -lt 3; $i++) { $c = "B$($RW.t1 + $i)"; Put $ws $c ([double]($i+1)) 'in'; $ws.Range($c).NumberFormat = ('0" - {0}"' -f $scnNames[$i]); $ws.Range($c).HorizontalAlignment = -4131 }
  [void]$ws.Range("B$($RW.tCur):I$($RW.t1 + 2)").Table([Type]::Missing, $ws.Range($sel))
  $ws.Range("C$($RW.tCur):D$($RW.t1 + 2)").NumberFormat = $fUSD; $ws.Range("E$($RW.tCur):I$($RW.t1 + 2)").NumberFormat = $fEPS
  # formats
  $L = $lastCol
  foreach ($k in 'sales','adj','gaap','cogs','gp','sga','rd','op','other','pretax','tax','nci','ni','gop','gni','cash','ar','inv','ppe','assets','ap','debt','bsnci','eq','netcash','nwc','ocf','capex','hedge','afcf','dep','amort','sbc','div','buyback','ebit','ncc','aebit','etax','nopat','pda','pcx','pnwc','ufcf','ufcfc','dcf') { $ws.Range("C$($RW[$k]):$L$($RW[$k])").NumberFormat = $fUSD }
  foreach ($row in $RW.seg) { $ws.Range("C${row}:$L$row").NumberFormat = $fUSD }
  foreach ($k in 'eps','epsrep','geps') { $ws.Range("C$($RW[$k]):$L$($RW[$k])").NumberFormat = $fEPS }
  $ws.Range("C$($RW.sh):$L$($RW.sh)").NumberFormat = $fNUM
  $ws.Range("C$($RW.gSales):$L$($RW.adjP)").NumberFormat = $fPCT; $ws.Range("C$($RW.da):$L$($RW.nonc)").NumberFormat = $fPCT
  $ws.Range("C$($RW.frac):$L$($RW.frac)").NumberFormat = '0.00'; $ws.Range("C$($RW.per):$L$($RW.per)").NumberFormat = '0.00'; $ws.Range("C$($RW.df):$L$($RW.df)").NumberFormat = '0.000'
  foreach ($k in 'sales','gaap','gp','op','ni','eps','afcf','ufcf') { $ws.Range("B$($RW[$k]):$L$($RW[$k])").Font.Bold = $true }
  foreach ($k in 'sales','op','ni','ufcf') { $ws.Range("C$($RW[$k]):$L$($RW[$k])").Borders.Item(8).LineStyle = 1 }
  # row notes
  if ($CFG.rowNotes) { foreach ($p in $CFG.rowNotes.PSObject.Properties) { if ($RW.ContainsKey($p.Name)) { Note $ws "$noteCol$($RW[$p.Name])" $p.Value } } }
  Note $ws "${noteCol}1" 'Blue = hardcoded input (filings or assumption); black = formula; green = link to another tab; yellow = key lever.'
  CW $ws 'A' 5; CW $ws 'B' 44; foreach ($c in $qc) { CW $ws $c 9 }; CW $ws 'C' 13; CW $ws (ColL (3 + $nQ)) 2; foreach ($c in $ac) { CW $ws $c 9.5 }; CW $ws (ColL ($yStart + $nY + $nP)) 2; CW $ws $noteCol 110

  # ---------------- Scenarios ----------------
  $w = $wsS
  Lbl $w 'B2' "Scenario drivers (selected with Model!$sel`: 1 = Bear, 2 = Base, 3 = Bull)"; $w.Range('B2').Font.Bold = $true; $w.Range('B2').Font.Size = 12
  Txt $w 'B4' 'Driver'; for ($j = 0; $j -lt $scenYears.Count; $j++) { Txt $w "$(ColL (3+$j))4" $Ps[$scenYears[$j]].label }; $gCol = ColL (3 + $scenYears.Count); Txt $w "${gCol}4" 'Rationale'
  $w.Range("B4:${gCol}4").Font.Bold = $true
  $blocks = @(); for ($s = 0; $s -lt $nS; $s++) { $blocks += @{row=$scRow["seg$s"]; title="$($segs[$s].name) - y/y growth"; sc=$segs[$s].scen} }
  $blocks += @{row=$scRow['gm']; title='Gross margin'; sc=$CFG.gmScen}
  foreach ($b in $blocks) {
    $rr = $b.row; Lbl $w "B$rr" $b.title; $w.Range("B${rr}:${gCol}${rr}").Interior.Color = $SECT; $w.Range("B$rr").Font.Bold = $true
    $nm = @('bear','base','bull'); $lb = @('  Bear','  Base','  Bull')
    for ($i = 0; $i -lt 3; $i++) {
      $x = $rr + 1 + $i; Lbl $w "B$x" $lb[$i]
      $vals = Arr $b.sc.($nm[$i]); for ($j = 0; $j -lt $scenYears.Count; $j++) { Put $w "$(ColL (3+$j))$x" (At $vals $j) 'in' }
      $w.Range("C${x}:$(ColL (2+$scenYears.Count))$x").NumberFormat = $fPCT
      if ($b.sc.notes) { Note $w "$gCol$x" $b.sc.notes.($nm[$i]) }
      [void]$w.Range("B${x}:${gCol}$x").FormatConditions.Add(2, [Type]::Missing, ("=Model!$selCell={0}" -f ($i+1))); $w.Range("B${x}:${gCol}$x").FormatConditions.Item(1).Interior.Color = (BGR 'FFF2CC')
    }
  }
  CW $w 'A' 5; CW $w 'B' 48; for ($j = 0; $j -lt $scenYears.Count; $j++) { CW $w (ColL (3+$j)) 10 }; CW $w $gCol 80

  # ---------------- WACC ----------------
  $w = $wsW; $WD = $CFG.wacc
  Lbl $w 'B2' 'WACC, terminal assumptions & sensitivity'; $w.Range('B2').Font.Bold = $true; $w.Range('B2').Font.Size = 12
  $rows = [ordered]@{
    4=@('#','Market inputs'); 5=@('Share price ($)','=Main!I2',$fEPS,'lk',''); 6=@('Diluted shares (mm)','=Main!I3',$fNUM,'lk',''); 7=@('Market capitalization ($mm)','=C5*C6',$fUSD,'','')
    8=@('Total debt ($mm)','=Main!I6',$fUSD,'lk',''); 9=@('Cash ($mm)','=Main!I5',$fUSD,'lk',''); 10=@('Non-controlling interest ($mm)','=Main!I8',$fUSD,'lk','')
    12=@('#','Cost of equity (CAPM)'); 13=@('Risk-free rate (10-yr UST)',[double]$WD.rf,'0.00%','in',$WD.rfNote); 14=@('Equity risk premium',[double]$WD.erp,'0.00%','in',$WD.erpNote); 15=@('Levered beta',[double]$WD.beta,'0.00','in',$WD.betaNote); 16=@('Cost of equity','=C13+C15*C14','0.00%','','Rf + beta x ERP')
    18=@('#','Cost of debt'); 19=@('Pre-tax cost of debt',[double]$WD.kd,'0.00%','in',$WD.kdNote); 20=@('Tax rate',[double]$WD.tax,$fPCT,'in',$WD.taxNote); 21=@('After-tax cost of debt','=C19*(1-C20)','0.00%','','')
    23=@('#','Capital structure (market values)'); 24=@('Equity weight','=C7/(C7+C8)',$fPCT,'',''); 25=@('Debt weight','=C8/(C7+C8)',$fPCT,'',''); 27=@('WACC','=C24*C16+C25*C21','0.00%','','')
    29=@('#','Terminal value & timing'); 30=@('Terminal growth (g)',[double]$WD.g,$fPCT,'in',$WD.gNote); 31=@('Exit EV/EBITDA multiple',[double]$WD.exitMultiple,$fMULT,'in',$WD.multNote); 32=@('Valuation date',(ToDate $WD.valDate),'m/d/yyyy','in','Cash flows discounted to this date')
  }
  foreach ($e in $rows.GetEnumerator()) { $k = $e.Key; $a = $e.Value; if ($a[0] -eq '#') { Sect $w "B${k}:D$k" $a[1]; continue }; Lbl $w "B$k" $a[0]; Put $w "C$k" $a[1] $a[3]; $w.Range("C$k").NumberFormat = $a[2]; Note $w "D$k" $a[4] }
  foreach ($a in 'C13','C14','C15','C19','C30','C31') { $w.Range($a).Interior.Color = $YELLOW }
  $w.Range('B27:C27').Font.Bold = $true; $w.Range('B27:C27').Interior.Color = $DCFF
  $ufc = "Model!`$$p0`$$($RW.ufcfc):`$$pN`$$($RW.ufcfc)"; $prr = "Model!`$$p0`$$($RW.per):`$$pN`$$($RW.per)"
  $bridge = "+Model!`$C`$$($RW.ldebt)+Model!`$C`$$($RW.pcash)+Model!`$C`$$($RW.lnci))/Model!`$C`$$($RW.shs)"
  Sect $w 'B34:H34' 'Sensitivity: value per share ($) - WACC (rows) vs terminal growth (columns), Gordon growth'
  Lbl $w 'B35' 'WACC  \  g'; Lbl $w 'B45' 'WACC  \  multiple'
  $gc = @('C','D','E','F','G'); $gd = @(-0.01,-0.005,0,0.005,0.01); $md = @(-6,-3,0,3,6); $wDelta = @(-0.015,-0.01,-0.005,0,0.005,0.01,0.015)
  for ($i = 0; $i -lt 5; $i++) { Put $w "$($gc[$i])35" ('=$C$30+({0})' -f (Inv $gd[$i])); $w.Range("$($gc[$i])35").NumberFormat = $fPCT; Put $w "$($gc[$i])45" ('=$C$31+({0})' -f $md[$i]); $w.Range("$($gc[$i])45").NumberFormat = $fMULT }
  for ($j = 0; $j -lt 7; $j++) { foreach ($b in 36,46) { $x = $b + $j; Put $w "B$x" ('=$C$27+({0})' -f (Inv $wDelta[$j])); $w.Range("B$x").NumberFormat = '0.00%' } }
  $w.Range('B35:G35').Font.Bold = $true; $w.Range('B45:G45').Font.Bold = $true; $w.Range('B36:B52').Font.Bold = $true; $w.Range('B35').HorizontalAlignment = -4152; $w.Range('B45').HorizontalAlignment = -4152
  FR $w 'C36:G42' "=(SUMPRODUCT($ufc,(1+`$B36)^(-$prr))+Model!`$$pN`$$($RW.ufcf)*(1+C`$35)/(`$B36-C`$35)/(1+`$B36)^Model!`$C`$$($RW.yrs)$bridge"
  Sect $w 'B44:H44' 'Sensitivity: value per share ($) - WACC (rows) vs exit EV/EBITDA multiple (columns)'
  FR $w 'C46:G52' "=(SUMPRODUCT($ufc,(1+`$B46)^(-$prr))+Model!`$C`$$($RW.ebitda)*C`$45/(1+`$B46)^Model!`$C`$$($RW.yrs)$bridge"
  $w.Range('C36:G42').NumberFormat = $fEPS; $w.Range('C46:G52').NumberFormat = $fEPS   # leave the multiple header row (45) as x
  foreach ($rng in 'C36:G42','C46:G52') { $fc = $w.Range($rng).FormatConditions.Add(1, 7, '=$C$5'); $fc.Interior.Color = (BGR 'C6EFCE') }
  foreach ($a in 'E39','E49') { $w.Range($a).Font.Bold = $true; $w.Range($a).Borders.LineStyle = 1 }
  Note $w 'B54' 'Green = value at or above the current share price. Boxed center cell = base case (ties to the Model valuation).'
  CW $w 'A' 5; CW $w 'B' 34; foreach ($c in $gc) { CW $w $c 11 }

  # ---------------- Contracts ----------------
  $w = $wsC
  Lbl $w 'B2' "$($CFG.company): biggest contracts & partnerships (ranked by estimated economic scale)"; $w.Range('B2').Font.Bold = $true; $w.Range('B2').Font.Size = 12
  $ch = @('#','Counterparty','Announced','Segment','Disclosed value ($mm)','Value type','Term (yrs)','Implied $mm / yr','Scope - what the company supplies','Why it matters for the model','Source')
  $cc = @('B','C','D','E','F','G','H','I','J','K','L')
  for ($i = 0; $i -lt 11; $i++) { $c = $w.Range("$($cc[$i])4"); SetP $c 'Value2' $ch[$i]; $c.Font.Bold = $true; $c.WrapText = $true; $c.Interior.Color = $SECT }
  $x = 5; $n = 1
  foreach ($k in (Arr $CFG.contracts)) {
    Put $w "B$x" ([double]$n); Lbl $w "C$x" $k.counterparty; $w.Range("C$x").Font.Bold = $true
    if ($k.date -match '^\d{4}-\d{2}-\d{2}$') { Put $w "D$x" (ToDate $k.date); $w.Range("D$x").NumberFormat = 'mmm d, yyyy' } else { Txt $w "D$x" $k.date }
    Lbl $w "E$x" $k.segment
    if ($null -ne $k.value) { Put $w "F$x" ([double]$k.value) 'in' } else { Lbl $w "F$x" 'n/d'; $w.Range("F$x").HorizontalAlignment = -4152 }
    Lbl $w "G$x" $k.type
    if ($null -ne $k.term) { Put $w "H$x" ([double]$k.term) 'in' } else { Lbl $w "H$x" 'n/d'; $w.Range("H$x").HorizontalAlignment = -4152 }
    if ($k.implied) { Put $w "I$x" "=F$x/H$x" } else { Lbl $w "I$x" 'n/m'; $w.Range("I$x").HorizontalAlignment = -4152 }
    Lbl $w "J$x" $k.scope; Lbl $w "K$x" $k.why; Lbl $w "L$x" $k.source
    $x++; $n++
  }
  $lastK = $x - 1
  if ($lastK -ge 5) {
    $w.Range("F5:F$lastK").NumberFormat = $fUSD; $w.Range("I5:I$lastK").NumberFormat = $fUSD
    $w.Range("B5:L$lastK").VerticalAlignment = -4160; foreach ($c in 'C','E','G','J','K','L') { $w.Range("${c}5:$c$lastK").WrapText = $true }
    $w.Range("B4:L$lastK").Borders.Item(12).LineStyle = 1; $w.Range("B4:L$lastK").Borders.Item(12).Color = (BGR 'D9D9D9')
  }
  Note $w "C$($lastK + 2)" 'Values mix revenue ceilings, customer investments and grants - not additive. n/d = not disclosed; n/m = not meaningful. Implied $/yr only where value and term are known.'
  CW $w 'A' 5; CW $w 'B' 4; CW $w 'C' 22; CW $w 'D' 13; CW $w 'E' 18; CW $w 'F' 12; CW $w 'G' 26; CW $w 'H' 8; CW $w 'I' 10; CW $w 'J' 60; CW $w 'K' 48; CW $w 'L' 36
  [void]$w.Range("B5:L$([Math]::Max(5,$lastK))").Rows.AutoFit()

  # ---------------- Sources ----------------
  $w = $wsSrc
  Lbl $w 'B2' "Sources, methodology & notes (data as of $($CFG.asOf))"; $w.Range('B2').Font.Bold = $true; $w.Range('B2').Font.Size = 12
  Lbl $w 'B4' 'Source'; Lbl $w 'C4' 'Link'; $w.Range('B4:C4').Font.Bold = $true; $w.Range('B4:C4').Interior.Color = $SECT
  $x = 5
  foreach ($s in (Arr $CFG.sources)) { Lbl $w "B$x" $s[0]; Lbl $w "C$x" $s[1]; if ($s[1] -match '^https?://') { [void]$w.Hyperlinks.Add($w.Range("C$x"), $s[1]) }; $x++ }
  $x++; Sect $w "B${x}:C$x" 'Methodology notes'; $x++
  foreach ($m in (Arr $CFG.methodology)) { Lbl $w "B$x" $m; $w.Range("B$x").WrapText = $true; $x++ }
  CW $w 'A' 5; CW $w 'B' 110; CW $w 'C' 100

  # ---------------- Main ----------------
  $w = $wsMain; $MK = $CFG.market
  Lbl $w 'B2' "$($CFG.company) ($($CFG.exchange): $($CFG.ticker))"; $w.Range('B2').Font.Bold = $true; $w.Range('B2').Font.Size = 14
  Lbl $w 'B3' "Shkreli-style operating model - `$ in millions except per share - data as of $($CFG.asOf)"; $w.Range('B3').Font.Italic = $true
  $box = [ordered]@{ 2=@('Price',[double]$MK.price,$fEPS,'in',$MK.priceTag); 3=@('Shares',[double]$MK.shares,$fNUM,'in',$MK.sharesTag); 4=@('MC','=I2*I3',$fUSD,'',''); 5=@('Cash',$MK.cash,$fUSD,'in',$MK.bsTag); 6=@('Debt',$MK.debt,$fUSD,'in',$MK.bsTag); 7=@('EV','=I4-I5+I6',$fUSD,'',''); 8=@('NCI',[double]$MK.nci,$fUSD,'in',$MK.bsTag); 9=@('Dividend / sh',[double]$MK.dps,$fEPS,'in','annualized'); 10=@('Dividend yield','=I9/I2',$fPCT,'','') }
  foreach ($e in $box.GetEnumerator()) { $k = $e.Key; $a = $e.Value; Lbl $w "H$k" $a[0]; Put $w "I$k" $a[1] $a[3]; $w.Range("I$k").NumberFormat = $a[2]; if ($a[4]) { Lbl $w "J$k" $a[4] } }
  foreach ($b in 7,8,9,10) { $w.Range('H2:J10').Borders.Item($b).LineStyle = 1 }
  $w.Range('H2:H10').Font.Bold = $true; $w.Range('J2:J10').Font.Italic = $true; $w.Range('J2:J10').Font.Color = $NOTE
  Sect $w 'H12:J12' 'Valuation & multiples'
  $vp = [ordered]@{ 13=@('Scenario',"=Model!$nameCol`1",'@','lk'); 14=@('DCF value / sh (Gordon)',"=Model!C$($RW.vg)",$fEPS,'lk'); 15=@('DCF value / sh (exit multiple)',"=Model!C$($RW.vm)",$fEPS,'lk'); 16=@('Shkreli NPV value / sh',"=Model!C$($RW.nvps)",$fEPS,'lk')
    17=@('Upside vs price (Gordon DCF)','=I14/I2-1',$fPCT,''); 18=@('WACC','=WACC!C27','0.00%','lk'); 19=@("P/E $($Ps[0].label)","=I2/Model!$p0$($RW.eps)",$fMULT,'lk'); 20=@("P/E $($Ps[1].label)","=I2/Model!$($pc[1])$($RW.eps)",$fMULT,'lk')
    21=@("EV / EBITDA $($Ps[0].label)","=I7/(Model!$p0$($RW.op)+Model!$p0$($RW.pda))",$fMULT,'lk'); 22=@("Consensus EPS $($Ps[0].label)",$MK.consensusEps,$fEPS,'in'); 23=@('Avg. analyst price target',$MK.avgTarget,$fEPS,'in') }
  foreach ($e in $vp.GetEnumerator()) { $k = $e.Key; $a = $e.Value; Lbl $w "H$k" $a[0]; Put $w "I$k" $a[1] $a[3]; $w.Range("I$k").NumberFormat = $a[2]; $w.Range("I$k").HorizontalAlignment = -4152 }
  if ($MK.consensusSource) { Lbl $w 'J22' $MK.consensusSource; Lbl $w 'J23' $MK.consensusSource; $w.Range('J22:J23').Font.Italic = $true; $w.Range('J22:J23').Font.Color = $NOTE }
  $w.Range('H14:I14').Font.Bold = $true; $w.Range('H14:I14').Interior.Color = $DCFF
  # left column
  $x = 5
  Sect $w "B${x}:F$x" "Segments ($basis sales, `$mm)"; Txt $w "C$x" $Ys[-1].label; Txt $w "D$x" '% of total'; Txt $w "E$x" $Ps[0].label; Txt $w "F$x" 'y/y'; $w.Range("C${x}:F$x").HorizontalAlignment = -4152
  $segTop = $x + 1
  for ($s = 0; $s -lt $nS; $s++) { $x++; $mr = $RW.seg[$s]; Lbl $w "B$x" $segs[$s].name; Put $w "C$x" "=Model!$lastH$mr" 'lk'; Put $w "E$x" "=Model!$p0$mr" 'lk'; Put $w "F$x" "=IFERROR(E$x/C$x-1,"""")" }
  $x++; $tot = $x; Lbl $w "B$x" 'Total'; Put $w "C$x" "=SUM(C${segTop}:C$($x-1))"; Put $w "E$x" "=SUM(E${segTop}:E$($x-1))"; Put $w "F$x" "=E$x/C$x-1"
  for ($y = $segTop; $y -le $tot; $y++) { Put $w "D$y" "=C$y/C`$$tot" }
  $w.Range("B${tot}:F$tot").Font.Bold = $true; $w.Range("C${tot}:F$tot").Borders.Item(8).LineStyle = 1
  $w.Range("C${segTop}:C$tot").NumberFormat = $fUSD; $w.Range("E${segTop}:E$tot").NumberFormat = $fUSD; $w.Range("D${segTop}:D$tot").NumberFormat = $fPCT; $w.Range("F${segTop}:F$tot").NumberFormat = $fPCT
  foreach ($blk in @(@('What each segment does',$CFG.main.descLines), @('Management',$CFG.main.management), @('Plan & guidance',$CFG.main.guidance))) {
    $x += 2; Sect $w "B${x}:F$x" $blk[0]; foreach ($t in (Arr $blk[1])) { $x++; Lbl $w "B$x" $t }
  }
  $x += 2; Sect $w "B${x}:F$x" 'Biggest contracts (details on Contracts tab)'; Txt $w "C$x" 'Value $mm'; Txt $w "D$x" 'Announced'; Txt $w "E$x" 'Type'
  $nK = [Math]::Min(6, (Arr $CFG.contracts).Count)
  for ($i = 0; $i -lt $nK; $i++) { $x++; $cr = 5 + $i; Put $w "B$x" "=Contracts!C$cr" 'lk'; Put $w "C$x" "=Contracts!F$cr" 'lk'; Put $w "D$x" "=Contracts!D$cr" 'lk'; Put $w "E$x" "=Contracts!G$cr" 'lk'; $w.Range("C$x").NumberFormat = $fUSD; $w.Range("C$x").HorizontalAlignment = -4152; $w.Range("D$x").NumberFormat = 'mmm yyyy'; $w.Range("D$x").HorizontalAlignment = -4152 }
  $x += 2; Sect $w "B${x}:F$x" 'Key takeaways (update live with the model)'
  $lead = $RW.seg[0]; $ln = $segs[0].name
  $tk = @(
    "=""- DCF (Gordon; WACC ""&TEXT(WACC!C27,""0.0%"")&"", g ""&TEXT(WACC!C30,""0.0%"")&""): `$""&TEXT(Model!C$($RW.vg),""0"")&""/sh vs. `$""&TEXT(I2,""0.00"")&"" share price (""&TEXT(Model!C$($RW.vg)/I2-1,""+0%;-0%"")&"").""",
    "=""- Exit-multiple DCF (""&TEXT(WACC!C31,""0.0"")&""x EBITDA): `$""&TEXT(Model!C$($RW.vm),""0"")&""/sh; Shkreli-style NPV of net income: `$""&TEXT(Model!C$($RW.nvps),""0"")&""/sh.""",
    "=""- $ln goes from ""&TEXT(Model!$lastH$lead/Model!$lastH$($RW.sales),""0%"")&"" of sales ($($Ys[-1].label)) to ""&TEXT(Model!$pN$lead/Model!$pN$($RW.sales),""0%"")&"" ($($Ps[-1].label)) in the ""&Model!$nameCol`1&"" case; sales `$""&TEXT(Model!$lastH$($RW.sales)/1000,""0.0"")&""B -> `$""&TEXT(Model!$pN$($RW.sales)/1000,""0.0"")&""B.""",
    "=""- Operating margin ""&TEXT(Model!$lastH$($RW.opm),""0.0%"")&"" ($($Ys[-1].label)) -> ""&TEXT(Model!$pN$($RW.opm),""0.0%"")&"" ($($Ps[-1].label)); EPS `$""&TEXT(Model!$lastH$($RW.eps),""0.00"")&"" -> `$""&TEXT(Model!$pN$($RW.eps),""0.00"")&"" ($($Ps[0].label) `$""&TEXT(Model!$p0$($RW.eps),""0.00"")&"" vs. consensus `$""&TEXT(I22,""0.00"")&"").""",
    "=IF(Model!C$($RW.vg)<I2,""- The market price implies more than this case: compare the Bull row of the scenario table (Model) and the WACC sensitivities."",""- The share price is below this case's intrinsic value."")"
  )
  foreach ($t in $tk) { $x++; Put $w "B$x" $t }
  foreach ($t in (Arr $CFG.main.extraTakeaways)) { $x++; Lbl $w "B$x" $t }
  $x += 2; Note $w "B$x" 'Blue = input, black = formula, green = link to another tab. Sources on the Sources tab. Not investment advice.'
  $x++; Note $w "B$x" "Key levers: scenario selector Model!$sel (1 Bear / 2 Base / 3 Bull), drivers on the Scenarios tab, yellow cells on the WACC tab, TV method Model!C$($RW.tvmeth)."
  CW $w 'A' 3; CW $w 'B' 46; CW $w 'C' 12; CW $w 'D' 12; CW $w 'E' 12; CW $w 'F' 9; CW $w 'G' 3; CW $w 'H' 30; CW $w 'I' 12; CW $w 'J' 16
  # charts
  $x += 2; $top = $w.Range("B$x").Top
  $pal = @('2a78d6','eb6834','1baf7a','eda100','e87ba4','008300','6250d6','e34948')
  $co = $w.ChartObjects().Add($w.Range("B$x").Left, $top, 520, 290); $chart = $co.Chart; $chart.ChartType = 52
  while ($chart.SeriesCollection().Count -gt 0) { $chart.SeriesCollection(1).Delete() }
  $fa = $ac[0]
  for ($s = 0; $s -lt [Math]::Min($nS, 8); $s++) { $sr2 = $chart.SeriesCollection().NewSeries(); $mr = $RW.seg[$s]; $sr2.Name = "=Model!`$B`$$mr"; $sr2.Values = "=Model!`$$fa`$${mr}:`$$pN`$$mr"; $sr2.XValues = "=Model!`$$fa`$3:`$$pN`$3"; $sr2.Format.Fill.ForeColor.RGB = (BGR $pal[$s]); $sr2.Format.Line.Visible = -1; $sr2.Format.Line.ForeColor.RGB = $WHITE; $sr2.Format.Line.Weight = 1.5 }
  $chart.ChartGroups(1).GapWidth = 55; $chart.HasTitle = $true; $chart.HasLegend = $true; $chart.Legend.Position = -4107
  $chart.ChartArea.Font.Name = 'Arial'; $chart.ChartArea.Font.Size = 9; $chart.ChartArea.Font.Color = (BGR '52514e'); $chart.ChartArea.Format.Line.Visible = 0
  $chart.ChartTitle.Text = 'Sales by segment ($bn) - selected scenario'; $chart.ChartTitle.Font.Size = 11; $chart.ChartTitle.Font.Bold = $true; $chart.ChartTitle.Font.Color = (BGR '0b0b0b')
  $ax = $chart.Axes(2); $ax.TickLabels.NumberFormat = '$#,##0,"B"'; $ax.MajorGridlines.Format.Line.ForeColor.RGB = (BGR 'e1e0d9'); $ax.Format.Line.Visible = 0
  $co2 = $w.ChartObjects().Add($w.Range("H$x").Left, $top, 420, 290); $c2 = $co2.Chart; $c2.ChartType = 51
  while ($c2.SeriesCollection().Count -gt 0) { $c2.SeriesCollection(1).Delete() }
  $sr2 = $c2.SeriesCollection().NewSeries(); $sr2.Name = 'EPS'; $sr2.Values = "=Model!`$$fa`$$($RW.eps):`$$pN`$$($RW.eps)"; $sr2.XValues = "=Model!`$$fa`$3:`$$pN`$3"; $sr2.Format.Fill.ForeColor.RGB = (BGR '2a78d6')
  for ($p = $nY + 1; $p -le $nY + $nP; $p++) { $sr2.Points($p).Format.Fill.ForeColor.RGB = (BGR '86b6ef') }
  foreach ($p in @($nY, ($nY + $nP))) { $pt = $sr2.Points($p); $pt.HasDataLabel = $true; $pt.DataLabel.NumberFormat = '$0.00'; $pt.DataLabel.Position = 2 }
  $c2.ChartGroups(1).GapWidth = 60; $c2.HasTitle = $true; $c2.HasLegend = $false
  $c2.ChartArea.Font.Name = 'Arial'; $c2.ChartArea.Font.Size = 9; $c2.ChartArea.Font.Color = (BGR '52514e'); $c2.ChartArea.Format.Line.Visible = 0
  $c2.ChartTitle.Text = "EPS ($basis; light bars = estimates)"; $c2.ChartTitle.Font.Size = 11; $c2.ChartTitle.Font.Bold = $true; $c2.ChartTitle.Font.Color = (BGR '0b0b0b')
  $ax = $c2.Axes(2); $ax.TickLabels.NumberFormat = '$0.00'; $ax.MajorGridlines.Format.Line.ForeColor.RGB = (BGR 'e1e0d9'); $ax.Format.Line.Visible = 0

  # tidy & save
  try { $ws.Activate(); $ws.Range('C4').Select() | Out-Null; $xl.ActiveWindow.FreezePanes = $true; $xl.ActiveWindow.Zoom = 90 } catch { }
  try { $wsMain.Activate(); $wsMain.Range('A1').Select() | Out-Null; $xl.ActiveWindow.DisplayGridlines = $false } catch { }
  $xl.CalculateFull()
  if (Test-Path $OutPath) { Remove-Item $OutPath -Force }
  New-Item -ItemType Directory -Force (Split-Path $OutPath) | Out-Null
  $wb.SaveAs($OutPath, 51)
  # write row map for the verifier
  ($RW | ConvertTo-Json -Compress) | Set-Content -Encoding UTF8 ($JsonPath -replace '\.json$', '.rowmap.json')
  "Saved: $OutPath"
}
catch { "ERROR at line $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())"; "  -> $($_.Exception.Message)"; exit 1 }
finally { if ($wb) { $wb.Close($false) }; $xl.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null }
