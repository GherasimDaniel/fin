// Shared data pipeline for the month delta: S&P 500 members, SEC share counts, Yahoo daily closes.
// update() keeps a rolling one-month window: the latest completed trading day back to the same date one month earlier.
const fs = require('fs'), path = require('path');

const ROOT = path.join(__dirname, '..', '..');
const DATA = path.join(ROOT, 'data'), RAW = path.join(DATA, 'raw');
const YDIR = path.join(RAW, 'yahoo'), CFDIR = path.join(RAW, 'sec_companyfacts');
const SEC_UA = 'fin-workshop research admin@fin-workshop.local';
const YH_UA = 'Mozilla/5.0';
const WIKI_UA = 'fin-workshop-research/1.0 (personal project)';
const DAY = 86400000;
const REF_MAX_AGE = 7 * DAY;          // refresh the member list and SEC share counts weekly
const sleep = ms => new Promise(r => setTimeout(r, ms));
[DATA, RAW, YDIR, CFDIR].forEach(d => fs.mkdirSync(d, { recursive: true }));

const GICS = ['Information Technology', 'Health Care', 'Financials', 'Consumer Discretionary', 'Communication Services', 'Industrials',
  'Consumer Staples', 'Energy', 'Utilities', 'Real Estate', 'Materials'];
// Wikipedia rows whose columns are shifted
const ROW_FIX = { EME: { cik: '0000105634', sector: 'Industrials', sub: 'Construction & Engineering' },
  KVUE: { cik: '0001944048', sector: 'Consumer Staples', sub: 'Personal Care Products' } };
// Multi-class companies: shares by class from 10-Q cover pages (scripts/fetch_cover.js). [shares, Yahoo symbol, conversion factor]
const COVER = {
  'BRK.B': { date: '2026-08-10', parts: [[488450, 'BRK-A', 1], [1408035161, 'BRK-B', 1]] },
  V: { date: '2026-07-29', parts: [[1704112694, 'V', 1], [2180148 + 486669 + 60589871, 'V', 1.5], [17059152, 'V', 4]] },
  ARES: { date: '2026-08-07', parts: [[223957742 + 3489911, 'ARES', 1]] },
  HSY: { date: '2026-10-29', parts: [[146366875 + 54613514, 'HSY', 1]] },
  STZ: { date: '2026-10-07', parts: [[169198487 + 25733, 'STZ', 1]] },
  ERIE: { date: '2026-10-29', parts: [[46189068, 'ERIE', 1], [2542, 'ERIE', 2400]] },
  TSN: { date: '2026-08-03', parts: [[281792507 + 70009005, 'TSN', 1]] },
  KVUE: { date: '2026-10-30', parts: [[1921768446, 'KVUE', 1]] },
  EME: { date: '2026-10-29', parts: [[43782004, 'EME', 1]] }
};

const ysym = s => s.replace('.', '-');
const isoDay = ms => new Date(ms).toISOString().slice(0, 10);
const ageOf = f => fs.existsSync(f) ? Date.now() - fs.statSync(f).mtimeMs : Infinity;
const readJSON = f => JSON.parse(fs.readFileSync(f, 'utf8'));
function nowET() {
  const p = Object.fromEntries(new Intl.DateTimeFormat('en-US', { timeZone: 'America/New_York', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hourCycle: 'h23' })
    .formatToParts(new Date()).map(x => [x.type, x.value]));
  return { date: `${p.year}-${p.month}-${p.day}`, minutes: +p.hour * 60 + +p.minute };
}
async function get(url, ua, as = 'json', tries = 3) {
  let last;
  for (let k = 0; k < tries; k++) {
    try {
      const r = await fetch(url, { headers: { 'User-Agent': ua } });
      if (r.ok) return as === 'json' ? r.json() : r.text();
      if (r.status === 404) return null;
      last = new Error(`HTTP ${r.status} from ${new URL(url).host}`);
    } catch (e) { last = e; }
    await sleep(1200 * (k + 1));
  }
  throw last;
}
async function pool(items, n, fn) {
  let i = 0;
  await Promise.all(Array.from({ length: n }, async () => { while (i < items.length) { const k = i++; await fn(items[k], k); } }));
}

// ---------- reference data: members, sectors, share counts ----------
function parseWiki(json) {
  const w = json.parse.wikitext, i = w.indexOf('{|'), tbl = w.slice(i, w.indexOf('|}', i));
  return tbl.split(/\n\|-/).slice(2).map(r => {
    const c = r.split(/\n\|\|?/).map(s => s.trim()).filter(Boolean);
    const sym = (c[0].match(/\{\{[^|}]+\|([^}|]+)/) || [])[1];
    const name = (c[1] || '').replace(/\[\[(?:[^\]|]*\|)?([^\]]*)\]\]/g, '$1').replace(/'''?/g, '');
    const row = { sym, name, sector: c[2], sub: c[3], cik: c[6] };
    if (ROW_FIX[sym]) Object.assign(row, ROW_FIX[sym]);
    return row;
  }).filter(r => r.sym);
}
function quarterFrames() { // current and two previous calendar quarters
  const d = new Date(), out = [];
  let y = d.getUTCFullYear(), q = Math.floor(d.getUTCMonth() / 3) + 1;
  for (let k = 0; k < 3; k++) { out.push([y, q]); q--; if (!q) { q = 4; y--; } }
  return out;
}
async function refreshReference(log, force) {
  const consFile = path.join(RAW, 'constituents.json'), framesFile = path.join(RAW, 'sec_frames.json');
  if (force || ageOf(consFile) > REF_MAX_AGE) {
    log('Refreshing the S&P 500 member list from Wikipedia');
    const j = await get('https://en.wikipedia.org/w/api.php?action=parse&page=List_of_S%26P_500_companies&prop=wikitext&format=json&formatversion=2', WIKI_UA);
    fs.writeFileSync(path.join(RAW, 'sp500_wikipedia.json'), JSON.stringify(j));
    const cons = parseWiki(j);
    // repair any CIK that did not parse, using SEC's ticker list
    if (cons.some(c => !/^\d{10}$/.test(c.cik || ''))) {
      const t = Object.values(await get('https://www.sec.gov/files/company_tickers.json', SEC_UA));
      cons.forEach(c => { if (!/^\d{10}$/.test(c.cik || '')) { const r = t.find(x => x.ticker === ysym(c.sym)); if (r) c.cik = String(r.cik_str).padStart(10, '0'); } });
    }
    fs.writeFileSync(consFile, JSON.stringify(cons, null, 1));
  }
  if (force || ageOf(framesFile) > REF_MAX_AGE) {
    log('Refreshing share counts from SEC');
    const frames = {};
    for (const [y, q] of quarterFrames()) {
      for (const f of [`dei/EntityCommonStockSharesOutstanding/shares/CY${y}Q${q}I`, `us-gaap/WeightedAverageNumberOfSharesOutstandingBasic/shares/CY${y}Q${q}`]) {
        const j = await get('https://data.sec.gov/api/xbrl/frames/' + f + '.json', SEC_UA).catch(() => null);
        frames[f] = j ? j.data : [];
        await sleep(250);
      }
    }
    fs.writeFileSync(framesFile, JSON.stringify(frames));
  }
  const cons = readJSON(consFile), frames = readJSON(framesFile);
  cons.forEach(c => { if (ROW_FIX[c.sym]) Object.assign(c, ROW_FIX[c.sym]); });
  // SEC companyfacts for members missing from the frames (cached for 30 days)
  const inFrames = new Set(); for (const k of Object.keys(frames)) for (const d of frames[k]) inFrames.add(String(d.cik).padStart(10, '0'));
  const need = [...new Set(cons.filter(c => /^\d{10}$/.test(c.cik) && !inFrames.has(c.cik) && !COVER[c.sym]).map(c => c.cik))]
    .filter(cik => ageOf(path.join(CFDIR, cik + '.json')) > 30 * DAY);
  if (need.length) log(`Reading SEC filings for ${need.length} companies`);
  for (const cik of need) {
    const j = await get(`https://data.sec.gov/api/xbrl/companyfacts/CIK${cik}.json`, SEC_UA).catch(() => null);
    if (j) {
      const pick = (ns, tag) => { const u = j.facts?.[ns]?.[tag]?.units?.shares; return u ? u.filter(x => x.form && /10-[QK]/.test(x.form)).sort((a, b) => a.end < b.end ? 1 : -1).slice(0, 3).map(x => ({ val: x.val, end: x.end })) : null; };
      fs.writeFileSync(path.join(CFDIR, cik + '.json'), JSON.stringify({ name: j.entityName, dei: pick('dei', 'EntityCommonStockSharesOutstanding'), wab: pick('us-gaap', 'WeightedAverageNumberOfSharesOutstandingBasic') }));
    }
    await sleep(200);
  }
  return { cons, frames };
}
function shareCounts(cons, frames) {
  const dei = {}, wab = {}, cutoff = isoDay(Date.now() - 365 * DAY);
  for (const k of Object.keys(frames)) for (const d of frames[k]) {
    const cik = String(d.cik).padStart(10, '0'), tgt = k.startsWith('dei') ? dei : wab;
    if (!tgt[cik] || d.end > tgt[cik].end) tgt[cik] = { v: d.val, end: d.end };
  }
  const out = {};
  for (const c of cons) {
    if (COVER[c.sym]) { out[c.sym] = { cover: COVER[c.sym], date: COVER[c.sym].date, src: '10-Q cover page, all share classes' }; continue; }
    const f = fs.existsSync(path.join(CFDIR, c.cik + '.json')) ? readJSON(path.join(CFDIR, c.cik + '.json')) : null;
    const fresh = x => x && (x.end || '') >= cutoff;
    let s = null;
    if (fresh(dei[c.cik])) s = { v: dei[c.cik].v, date: dei[c.cik].end, src: 'SEC cover page' };
    else if (f && f.dei && fresh(f.dei[0])) s = { v: f.dei[0].val, date: f.dei[0].end, src: 'SEC cover page' };
    else if (fresh(wab[c.cik])) s = { v: wab[c.cik].v, date: wab[c.cik].end, src: 'SEC weighted-average basic shares' };
    else if (f && f.wab && fresh(f.wab[0])) s = { v: f.wab[0].val, date: f.wab[0].end, src: 'SEC weighted-average basic shares' };
    if (s) out[c.sym] = s;
  }
  return out;
}

// ---------- prices ----------
async function yahoo(sym, fromMs) {
  const j = await get(`https://query1.finance.yahoo.com/v8/finance/chart/${sym}?period1=${Math.floor(fromMs / 1000)}&period2=${Math.floor(Date.now() / 1000) + 3600}&interval=1d&events=split`, YH_UA);
  const r = j && j.chart && j.chart.result && j.chart.result[0];
  if (!r || !r.timestamp) return null;
  const splits = r.events && r.events.splits ? Object.values(r.events.splits).map(x => ({ date: isoDay(x.date * 1000), num: x.numerator, den: x.denominator })) : [];
  return { days: r.timestamp.map(t => isoDay(t * 1000)), close: r.indicators.quote[0].close, splits };
}
function completed(days) { // drop today's bar while the market is still open
  const et = nowET();
  return days.filter(d => d < et.date || (d === et.date && et.minutes >= 16 * 60 + 20));
}
function windowDays(allDays) {
  const latest = allDays[allDays.length - 1];
  const [y, m, d] = latest.split('-').map(Number);
  const back = new Date(Date.UTC(y, m - 2, Math.min(d, new Date(Date.UTC(y, m - 1, 0)).getUTCDate())));
  const start = isoDay(back.getTime());
  const base = allDays.filter(x => x <= start).pop() || allDays[0];
  return allDays.filter(x => x >= base);
}
const splitFactor = (splits, since) => splits.reduce((f, s) => {
  const r = s.num / s.den, clean = Number.isInteger(r) || Number.isInteger(1 / r);
  return clean && s.date > since ? f * r : f;
}, 1);

// ---------- main update ----------
async function update({ onProgress = () => {}, force = false } = {}) {
  const log = m => onProgress({ phase: 'reference', message: m });
  const metaFile = path.join(DATA, 'meta.json');
  const meta = fs.existsSync(metaFile) ? readJSON(metaFile) : {};
  onProgress({ phase: 'checking', message: 'Checking for new trading days' });
  const from = Date.now() - 400 * DAY; // long enough to see splits since each share-count date
  const spy = await yahoo('SPY', from);
  if (!spy) throw new Error('Yahoo returned no data for SPY');
  const all = completed(spy.days);
  const latest = all[all.length - 1];
  if (!force && meta.latest === latest) return { added: [], latest, upToDate: true };

  const { cons, frames } = await refreshReference(log, false);
  const shares = shareCounts(cons, frames);
  const days = windowDays(all);
  const syms = [...new Set(cons.map(c => ysym(c.sym)).concat(Object.values(COVER).flatMap(c => c.parts.map(p => p[1]))))];
  const px = {};
  let done = 0;
  onProgress({ phase: 'prices', done, total: syms.length, message: 'Downloading daily closes' });
  await pool(syms, 6, async s => {
    try { px[s] = await yahoo(s, from); } catch (e) { px[s] = null; }
    done++; if (done % 10 === 0 || done === syms.length) onProgress({ phase: 'prices', done, total: syms.length, message: 'Downloading daily closes' });
    await sleep(80);
  });
  // keep only the one-month window on disk
  fs.readdirSync(YDIR).forEach(f => fs.unlinkSync(path.join(YDIR, f)));
  for (const [s, p] of Object.entries(px)) if (p) {
    const keep = p.days.map((d, k) => [d, p.close[k]]).filter(([d]) => d >= days[0] && d <= days[days.length - 1]);
    fs.writeFileSync(path.join(YDIR, s + '.json'), JSON.stringify({ sym: s, days: keep.map(x => x[0]), close: keep.map(x => x[1]), splits: p.splits }));
  }
  const closesOn = s => { const p = px[s]; if (!p) return null; const m = {}; p.days.forEach((d, k) => { m[d] = p.close[k]; }); return days.map(d => m[d]); };

  onProgress({ phase: 'building', message: 'Computing market values' });
  const seen = new Set(), companies = [], notes = [];
  for (const c of cons) {
    if (!GICS.includes(c.sector)) { notes.push(`${c.sym} left out: no sector in the member list`); continue; }
    if (seen.has(c.cik)) continue; // second listed share class (GOOG, FOX, NWS)
    seen.add(c.cik);
    const sh = shares[c.sym];
    if (!sh) { notes.push(`${c.sym} left out: no share count found at SEC`); continue; }
    let mc;
    if (sh.cover) {
      const series = sh.cover.parts.map(([n, s, conv]) => { const cl = closesOn(s); return cl && cl.map(p => p == null ? null : p * n * conv); });
      mc = series.some(x => !x) ? null : days.map((_, k) => series.every(s => s[k] != null) ? series.reduce((a, s) => a + s[k], 0) : null);
    } else {
      const cl = closesOn(ysym(c.sym));
      const f = cl ? splitFactor(px[ysym(c.sym)].splits, sh.date) : 1;
      mc = cl && cl.map(p => p == null ? null : p * sh.v * f);
      if (f !== 1) sh.src += ` (adjusted ${f}x for a split)`;
    }
    if (!mc) { notes.push(`${c.sym} left out: no prices from Yahoo`); continue; }
    for (let k = 1; k < mc.length; k++) if (mc[k] == null) mc[k] = mc[k - 1];
    if (mc[0] == null) { notes.push(`${c.sym} left out: not trading on ${days[0]}`); continue; }
    companies.push({ sym: c.sym, name: c.name, sector: c.sector, sub: c.sub, src: sh.src, sdate: sh.date, mc });
  }

  const csv = ['date,symbol,name,gics_sector,gics_sub_industry,market_value_usd_bn,share_source,share_date'];
  companies.forEach(c => c.mc.forEach((v, k) => csv.push([days[k], c.sym, '"' + c.name.replace(/"/g, '') + '"', c.sector, '"' + c.sub + '"', (v / 1e9).toFixed(3), '"' + c.src + '"', c.sdate].join(','))));
  fs.writeFileSync(path.join(DATA, 'daily_market_values.csv'), csv.join('\n') + '\n');
  const consDate = isoDay(fs.statSync(path.join(RAW, 'constituents.json')).mtimeMs);
  const page = { days, meta: { updatedAt: new Date().toISOString(), membersAsOf: consDate },
    companies: companies.map(c => ({ s: c.sym, n: c.name, g: c.sector, sub: c.sub, v: c.mc.map(x => Math.round(x / 1e7) / 100) })), notes };
  fs.writeFileSync(path.join(DATA, 'month.js'), 'const DATA=' + JSON.stringify(page) + ';\n');
  const added = meta.latest ? days.filter(d => d > meta.latest) : days.slice();
  fs.writeFileSync(metaFile, JSON.stringify({ latest, first: days[0], updatedAt: page.meta.updatedAt, companies: companies.length, notes }, null, 1));
  return { added, latest, first: days[0], companies: companies.length, notes };
}

module.exports = { update, ROOT, DATA };
