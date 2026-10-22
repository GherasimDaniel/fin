// fetch_sec.js <TICKER> <out.json> [--quarters 10] [--years 5]
// Pulls GAAP financials from SEC EDGAR XBRL (companyfacts) and writes a build_model.ps1 JSON skeleton.
// Everything that needs judgment (segments, non-GAAP basis, projections, contracts, price) is pre-filled
// with placeholders and listed in "_todo".
const fs = require('fs');
const UA = { 'User-Agent': 'Independent Research Model research-contact@example.org', 'Accept-Encoding': 'identity' };
const args = process.argv.slice(2);
const ticker = (args[0] || '').toUpperCase(); const outPath = args[1];
const opt = (k, d) => { const i = args.indexOf(k); return i >= 0 ? +args[i + 1] : d; };
const NQ = opt('--quarters', 10), NY = opt('--years', 5), NP = 5;
if (!ticker || !outPath) { console.error('usage: node fetch_sec.js TICKER out.json'); process.exit(1); }

const get = async (url, json = true) => { const r = await fetch(url, { headers: UA }); if (!r.ok) throw new Error(`${r.status} ${url}`); return json ? r.json() : r.text(); };
const day = 864e5, D = s => new Date(s + 'T00:00:00Z'), iso = d => d.toISOString().slice(0, 10);
const r0 = x => x == null ? null : Math.round(x * 100) / 100;

const TAGS = {
  sales: ['RevenueFromContractWithCustomerExcludingAssessedTax', 'Revenues', 'SalesRevenueNet', 'RevenueFromContractWithCustomerIncludingAssessedTax'],
  cogs: ['CostOfRevenue', 'CostOfGoodsAndServicesSold', 'CostOfGoodsSold'],
  gp: ['GrossProfit'], sga: ['SellingGeneralAndAdministrativeExpense'], rd: ['ResearchAndDevelopmentExpense', 'ResearchAndDevelopmentExpenseExcludingAcquiredInProcessCost'],
  op: ['OperatingIncomeLoss'], pretax: ['IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest', 'IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments'],
  tax: ['IncomeTaxExpenseBenefit'], ni: ['NetIncomeLoss'], eps: ['EarningsPerShareDiluted'], shares: ['WeightedAverageNumberOfDilutedSharesOutstanding'],
  ocf: ['NetCashProvidedByUsedInOperatingActivities'], capex: ['PaymentsToAcquirePropertyPlantAndEquipment', 'PaymentsToAcquireProductiveAssets'],
  dep: ['DepreciationDepletionAndAmortization', 'DepreciationAndAmortization', 'Depreciation'], amort: ['AmortizationOfIntangibleAssets'], sbc: ['ShareBasedCompensation', 'AllocatedShareBasedCompensationExpense'],
  div: ['PaymentsOfDividendsCommonStock', 'PaymentsOfDividends', 'PaymentsOfOrdinaryDividends'], buyback: ['PaymentsForRepurchaseOfCommonStock'],
  cash: ['CashAndCashEquivalentsAtCarryingValue', 'CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents'], ar: ['AccountsReceivableNetCurrent'], inv: ['InventoryNet'],
  ppe: ['PropertyPlantAndEquipmentNet'], assets: ['Assets'], ap: ['AccountsPayableCurrent', 'AccountsPayableTradeCurrent', 'AccountsPayableAndAccruedLiabilitiesCurrent'], nciBS: ['MinorityInterest'], equity: ['StockholdersEquity'],
  debtCur: ['LongTermDebtCurrent', 'DebtCurrent', 'LongTermDebtAndCapitalLeaseObligationsCurrent', 'ShortTermBorrowings'],
  debtLT: ['LongTermDebtNoncurrent', 'LongTermDebtAndCapitalLeaseObligations', 'LongTermDebt']
};
const INSTANT = new Set(['cash', 'ar', 'inv', 'ppe', 'assets', 'ap', 'nciBS', 'equity', 'debtCur', 'debtLT']);
const NEG = new Set(['capex', 'div', 'buyback']);

(async () => {
  const map = await get('https://www.sec.gov/files/company_tickers.json');
  const hit = Object.values(map).find(x => x.ticker === ticker);
  if (!hit) throw new Error('ticker not found in SEC list: ' + ticker);
  const cik = String(hit.cik_str).padStart(10, '0');
  const facts = await get(`https://data.sec.gov/api/xbrl/companyfacts/CIK${cik}.json`);
  const subs = await get(`https://data.sec.gov/submissions/CIK${cik}.json`);
  const g = facts.facts['us-gaap'] || {};
  const todo = [], used = {};

  // all facts for a key, de-duplicated by (start,end), latest filing wins
  // merge all candidate tags in priority order (per period, the first tag that reports it wins)
  const series = key => {
    const have = new Set(), vals = [], tags = [];
    for (const t of TAGS[key]) {
      const f = g[t]; if (!f) continue;
      const unit = Object.keys(f.units).find(u => u === 'USD' || u === 'shares' || u === 'USD/shares') || Object.keys(f.units)[0];
      const m = new Map();
      for (const v of f.units[unit]) { if (!/^10-[KQ]/.test(v.form || '')) continue; const k = (v.start || '') + '|' + v.end; if (!m.has(k) || v.filed > m.get(k).filed) m.set(k, v); }
      let added = 0; for (const [k, v] of m) if (!have.has(k)) { have.add(k); vals.push(v); added++; }
      if (added) tags.push(t);
    }
    if (tags.length) used[key] = tags.join(' + ');
    return vals;
  };
  const S = {}; for (const k of Object.keys(TAGS)) S[k] = series(k);

  // fiscal years from annual revenue (fallback net income)
  const annualOf = arr => arr.filter(v => v.start && (D(v.end) - D(v.start)) / day > 350 && (D(v.end) - D(v.start)) / day < 380);
  const fyEnds = [...new Set(annualOf(S.sales.length ? S.sales : S.ni).map(v => v.end))].sort().slice(-NY);
  const fyStartOf = {}; for (const v of annualOf(S.sales.length ? S.sales : S.ni)) fyStartOf[v.end] = v.start;
  // quarter ends from 10-Q / 10-K report dates
  const rec = subs.filings.recent; const qEnds = new Set();
  for (let i = 0; i < rec.form.length; i++) if (/^10-[QK]$/.test(rec.form[i]) && rec.reportDate[i]) qEnds.add(rec.reportDate[i]);
  const quarterEnds = [...qEnds].sort().slice(-NQ);

  const flow = (key, s, e) => {
    const arr = S[key]; if (!arr.length) return null;
    const exact = arr.find(v => v.start === s && v.end === e); if (exact) return exact.val;
    const ytd = arr.filter(v => v.end === e && v.start && v.start < s).sort((a, b) => a.start < b.start ? -1 : 1)[0];
    if (!ytd) return null;
    const prevEnd = iso(new Date(D(s) - day));
    if (ytd.start > prevEnd) return null;
    const prior = arr.find(v => v.start === ytd.start && v.end === prevEnd);
    return prior ? ytd.val - prior.val : null;
  };
  const inst = (key, e) => { const v = S[key].find(x => !x.start && x.end === e); return v ? v.val : null; };
  const val = (key, s, e) => {
    let x;
    if (INSTANT.has(key)) x = inst(key, e);
    else if (key === 'eps') { const v = S.eps.find(z => z.start === s && z.end === e); x = v ? v.val : null; }
    else if (key === 'shares') { const v = S.shares.find(z => z.start === s && z.end === e); x = v ? v.val : null; }
    else x = flow(key, s, e);
    if (x == null) return null;
    if (key === 'eps') return r0(x);
    x = key === 'shares' ? x / 1e6 : x / 1e6;
    if (NEG.has(key)) x = -Math.abs(x);
    return Math.round(x * 10) / 10;
  };

  const build = periods => {
    const o = {}; for (const k of Object.keys(TAGS)) o[k] = periods.map(p => val(k, p.s, p.e));
    const n = periods.length, out = {};
    out.gaapSales = o.sales;
    out.gp = o.gp.map((x, i) => x != null ? x : (o.sales[i] != null && o.cogs[i] != null ? o.sales[i] - o.cogs[i] : null));
    out.sga = o.sga; out.rd = o.rd.map(x => x == null ? 0 : x);
    out.gaapOpInc = o.op; out.pretax = o.pretax;
    out.taxRate = o.tax.map((x, i) => x != null && o.pretax[i] ? r0(x / o.pretax[i] * 1000) / 1000 : null);
    out.ni = o.ni; out.gaapNi = o.ni; out.shares = o.shares; out.epsRep = o.eps; out.gaapEps = o.eps;
    for (const k of ['cash', 'ar', 'inv', 'ppe', 'assets', 'ap', 'nciBS', 'equity', 'ocf', 'capex', 'dep', 'amort', 'sbc', 'div', 'buyback']) out[k] = o[k];
    out.debt = o.debtLT.map((x, i) => (x == null && o.debtCur[i] == null) ? null : (x || 0) + (o.debtCur[i] || 0));
    out.hedge = Array(n).fill(null);
    // SG&A missing: back into it so operating income ties (GP - SG&A - R&D = OpInc)
    out.sga = out.sga.map((x, i) => x != null ? x : (out.gp[i] != null && out.gaapOpInc[i] != null ? Math.round((out.gp[i] - out.rd[i] - out.gaapOpInc[i]) * 10) / 10 : null));
    return out;
  };
  const yPeriods = fyEnds.map(e => ({ s: fyStartOf[e], e }));
  const qPeriods = quarterEnds.map((e, i) => { const prev = i ? quarterEnds[i - 1] : iso(new Date(D(e) - 91 * day)); return { s: iso(new Date(D(prev).getTime() + day)), e }; });
  const y = build(yPeriods), q = build(qPeriods);
  // Q4 share counts are not reported separately: use that fiscal year's diluted weighted average (approximation)
  q.shares = q.shares.map((x, i) => { if (x != null) return x; const yi = fyEnds.indexOf(quarterEnds[i]); if (yi >= 0 && y.shares[yi] != null) { todo.push(`q.shares[${i}] (${quarterEnds[i]}) = FY diluted average (approximation)`); return y.shares[yi]; } return null; });

  // labels
  const fyLabel = e => e.slice(0, 4);
  const lastFyEnd = D(fyEnds[fyEnds.length - 1]);
  // fiscal-year ends from 10-K report dates (+ projected next ones) - handles 52/53-week years
  const fyList = [...new Set(rec.form.map((f, i) => f === '10-K' ? rec.reportDate[i] : null).filter(Boolean).concat(fyEnds))].sort();
  for (let k = 1; k <= 2; k++) { const d = D(fyList[fyList.length - 1]); fyList.push(iso(new Date(Date.UTC(d.getUTCFullYear() + 1, d.getUTCMonth(), d.getUTCDate())))); }
  const qLabel = e => {
    const fyE = fyList.find(f => D(f) - D(e) >= -10 * day); const fyEnd = D(fyE);
    const qn = Math.min(4, Math.max(1, 4 - Math.round((fyEnd - D(e)) / day / 91.3)));
    return { label: `Q${qn}'${String(fyEnd.getUTCFullYear()).slice(2)}`, qn, fyEnd };
  };
  const quarters = quarterEnds.map(e => ({ label: qLabel(e).label, end: e }));
  // estimate quarters to finish the current fiscal year
  const lastQ = quarterEnds[quarterEnds.length - 1]; const lq = qLabel(lastQ); const est = [];
  let curEnd = D(lastQ);
  for (let k = lq.qn + 1; k <= 4 && iso(lq.fyEnd) !== fyEnds[fyEnds.length - 1]; k++) {
    curEnd = new Date(Date.UTC(curEnd.getUTCFullYear(), curEnd.getUTCMonth() + 3, 1)); curEnd = new Date(Date.UTC(curEnd.getUTCFullYear(), curEnd.getUTCMonth() + 1, 0));
    est.push({ label: `Q${k}'${String(lq.fyEnd.getUTCFullYear()).slice(2)}E`, end: iso(k === 4 ? lq.fyEnd : curEnd), est: true });
  }
  const years = fyEnds.map(e => ({ label: fyLabel(e), end: e }));
  const proj = []; for (let k = 1; k <= NP; k++) { const d = new Date(Date.UTC(lastFyEnd.getUTCFullYear() + k, lastFyEnd.getUTCMonth(), lastFyEnd.getUTCDate())); proj.push({ label: `${d.getUTCFullYear()}E`, end: iso(d) }); }

  // treasury 10y (latest)
  let rf = null, rfDate = null;
  try {
    const yr = new Date().getUTCFullYear();
    const csv = await get(`https://home.treasury.gov/resource-center/data-chart-center/interest-rates/daily-treasury-rates.csv/${yr}/all?type=daily_treasury_yield_curve&field_tdr_date_value=${yr}&page&_format=csv`, false);
    const lines = csv.trim().split('\n'); const hdr = lines[0].split(',').map(s => s.replace(/"/g, '')); const ci = hdr.indexOf('10 Yr');
    const row = lines[1].split(','); rf = +row[ci] / 100; rfDate = row[0];
  } catch (e) { todo.push('wacc.rf: Treasury fetch failed - enter the 10-yr yield'); }

  const last = a => { for (let i = a.length - 1; i >= 0; i--) if (a[i] != null) return a[i]; return null; };
  const avg = a => { const v = a.filter(x => x != null); return v.length ? v.reduce((s, x) => s + x, 0) / v.length : null; };
  const ly = k => last(y[k]); const sales = ly('gaapSales');
  const ratio = (a, b) => (a != null && b) ? Math.round(a / b * 1000) / 1000 : null;
  const gm = ratio(ly('gp'), sales), sgaP = ratio(ly('sga'), sales), rdP = ratio(ly('rd'), sales), taxR = last(y.taxRate) ?? 0.21;
  const da = ratio(ly('dep'), sales), cx = ratio(-ly('capex'), sales);
  const nEst = est.length, nScen = NP - (nEst ? 1 : 0);
  const rep = (x, n) => Array(n).fill(x);
  const projArr = x => nEst ? [null, ...rep(x, NP - 1)] : rep(x, NP);
  const growth = y.gaapSales.length > 1 ? Math.round((Math.pow(sales / y.gaapSales[0], 1 / (y.gaapSales.length - 1)) - 1) * 1000) / 1000 : 0.05;

  const out = {
    company: hit.title, ticker, exchange: (subs.exchanges || [])[0] || '', asOf: new Date().toISOString().slice(0, 10),
    outFile: `C:\\Users\\gherasid\\work\\workshop\\fin\\${hit.title.replace(/[^A-Za-z0-9]+/g, '_')}\\${ticker}_Model.xlsx`,
    basis: 'GAAP', adjLabel: 'Adjustment to GAAP sales', scenario: 2, tvMethod: 1,
    market: { price: null, priceTag: 'TODO close', shares: last(q.shares), sharesTag: quarters.length ? quarters[quarters.length - 1].label + ' diluted' : '', cash: last(q.cash), debt: last(q.debt), nci: last(q.nciBS) || 0, bsTag: quarters.length ? quarters[quarters.length - 1].label : '', dps: null, consensusEps: null, avgTarget: null, consensusSource: '' },
    wacc: { rf, rfNote: rfDate ? `10-yr Treasury par yield on ${rfDate} (U.S. Treasury)` : '', erp: 0.045, erpNote: 'Long-run US ERP assumption', beta: null, betaNote: '', kd: rf ? Math.round((rf + 0.01) * 10000) / 10000 : null, kdNote: '10-yr UST + ~100bp (adjust for credit rating)', tax: taxR, taxNote: 'Latest effective tax rate', g: 0.03, gNote: 'Long-run nominal growth', exitMultiple: 12, multNote: 'TODO: set vs peers / current trading multiple', maturity: 0.03, valDate: new Date().toISOString().slice(0, 10) },
    periods: { quarters: [...quarters, ...est], years, proj },
    segments: [{ name: 'Total company', note: 'TODO: replace with reported segments (10-K segment note / earnings releases).', q: q.gaapSales, y: y.gaapSales, qEstGrowth: rep(growth, nEst),
      scen: { bear: rep(Math.round(growth * 500) / 1000, nScen), base: rep(growth, nScen), bull: rep(Math.round(growth * 1500) / 1000, nScen), notes: { bear: 'TODO', base: 'TODO', bull: 'TODO' } } }],
    gmScen: { bear: rep(gm != null ? gm - 0.02 : null, nScen), base: rep(gm, nScen), bull: rep(gm != null ? gm + 0.02 : null, nScen), notes: { bear: '', base: 'Latest FY gross margin', bull: '' } },
    q, y,
    qEst: { gm: rep(gm, nEst), sga: rep(sgaP, nEst), rd: rep(rdP, nEst), other: rep(Math.round(avg(q.pretax.map((p, i) => p != null && q.gaapOpInc[i] != null ? p - q.gaapOpInc[i] : null)) || 0), nEst), tax: rep(taxR, nEst), nci: rep(0, nEst), shareGrowth: rep(0, nEst), adjPct: rep(0, nEst) },
    proj: { sga: projArr(sgaP), rd: projArr(rdP), other: projArr(Math.round((ly('pretax') || 0) - (ly('gaapOpInc') || 0))), tax: projArr(taxR), nci: projArr(0), shareGrowth: projArr(0), adjPct: projArr(0),
            da: rep(da, NP), capex: rep(cx, NP), nwc: rep(0.15, NP), noncore: rep(0, NP), fraction: [0.5, 1, 1, 1, 1] },
    rowNotes: { gp: 'GAAP basis from SEC XBRL (companyfacts). Replace with company non-GAAP if it guides on that basis.' },
    main: { descLines: ['TODO: one line per segment'], management: ['TODO'], guidance: ['TODO: latest guidance / long-term targets'], extraTakeaways: [] },
    contracts: [], sources: [['SEC XBRL company facts', `https://data.sec.gov/api/xbrl/companyfacts/CIK${cik}.json`], ['SEC filings index', `https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=${cik}&type=10-K`]],
    methodology: ['GAAP figures from SEC XBRL; quarterly flows derived from year-to-date values where needed.', 'Not investment advice.'],
    _xbrlTagsUsed: used
  };
  // checks / todo list
  for (const k of ['gaapSales','gp','sga','rd','gaapOpInc','pretax','ni','shares','ocf','capex','dep','cash','debt']) if (y[k].every(x => x == null)) todo.push(`y.${k}: no annual values found in standard XBRL tags (company may use a custom tag) - fill from the 10-K`);
  for (const k of Object.keys(TAGS)) if (!used[k]) todo.push(`no XBRL tag found for ${k} (${TAGS[k][0]}...) - fill manually`);
  for (const k of ['gaapSales', 'gp', 'gaapOpInc', 'pretax', 'ni', 'shares']) q[k].forEach((x, i) => { if (x == null) todo.push(`q.${k}[${i}] (${quarters[i].label}) missing`); });
  y.gaapSales.forEach((s, i) => { const t = (y.gp[i] ?? 0) - (y.sga[i] ?? 0) - (y.rd[i] ?? 0); if (y.gaapOpInc[i] != null && Math.abs(t - y.gaapOpInc[i]) > 0.02 * Math.abs(s)) todo.push(`${years[i].label}: GP - SG&A - R&D (${t.toFixed(0)}) != operating income (${y.gaapOpInc[i]}) - other operating lines (amortization, restructuring) exist; fold them into SG&A or switch to the company's adjusted basis`); });
  todo.push('market.price, dps, consensusEps, avgTarget; wacc.beta', 'segments (+ scen, notes, noYoY), basis/adjusted P&L if the company guides on non-GAAP', 'qEst / proj assumptions (placeholders = latest FY ratios)', 'contracts (web research), main.* text, sources, methodology');
  out._todo = todo;
  fs.writeFileSync(outPath, JSON.stringify(out, null, 2));
  console.log(`Wrote ${outPath}: ${hit.title} (CIK ${cik}); ${years.length} years, ${quarters.length} quarters + ${est.length} estimate quarters`);
  console.log('TODO:\n - ' + todo.join('\n - '));
})().catch(e => { console.error('ERROR:', e.stack); process.exit(1); });
