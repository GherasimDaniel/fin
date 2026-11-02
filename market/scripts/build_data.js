const fs = require('fs');
const [,, indCsv, cpiCsv, outFile] = process.argv;
const lines = fs.readFileSync(indCsv, 'utf8').split(/\r?\n/);
function section(title) {
  const i = lines.findIndex(l => l.trim() === title);
  if (i < 0) throw new Error('missing ' + title);
  const header = lines[i + 1].split(',').slice(1).map(s => s.trim());
  const rows = [];
  for (let k = i + 2; k < lines.length; k++) {
    const l = lines[k];
    if (!/^\s*\d{6}\s*,/.test(l)) break;
    const p = l.split(',').map(s => s.trim());
    rows.push({ ym: +p[0], v: p.slice(1).map(Number) });
  }
  return { header, rows };
}
const nf = section('Number of Firms in Portfolios');
const sz = section('Average Firm Size');
if (nf.rows.length !== sz.rows.length) throw new Error('row mismatch');
const cpiRows = fs.readFileSync(cpiCsv, 'utf8').trim().split(/\r?\n/).slice(1).map(l => l.split(','));
const cpi = new Map(cpiRows.map(r => [+r[0].slice(0, 4) * 100 + +r[0].slice(5, 7), +r[1]]));
const months = nf.rows.map(r => r.ym);
const ind = nf.header;
const val = ind.map(() => []), firms = ind.map(() => []);
let bad = 0;
nf.rows.forEach((r, t) => {
  if (sz.rows[t].ym !== r.ym) throw new Error('ym mismatch');
  ind.forEach((_, j) => {
    let n = r.v[j], s = sz.rows[t].v[j];
    if (!(n >= 0) || !(s >= 0)) { bad++; n = Math.max(0, n || 0); s = Math.max(0, s || 0); }
    firms[j].push(n);
    val[j].push(Math.round(n * s));       // $ millions
  });
});
const cpiArr = months.map((m, t) => cpi.get(m));
for (let t = 0; t < cpiArr.length; t++) if (!cpiArr[t] && cpiArr[t-1] && cpi.get(months[t+1])) { cpiArr[t] = +Math.sqrt(cpiArr[t-1] * cpi.get(months[t+1])).toFixed(3); console.log('interpolated CPI', months[t], cpiArr[t]); }
if (cpiArr.some(x => !x)) throw new Error('cpi gap');
const out = { start: months[0], end: months[months.length - 1], n: months.length, ind, val, firms, cpi: cpiArr };
fs.writeFileSync(outFile, 'const DATA=' + JSON.stringify(out) + ';\n');
console.log('months', months.length, months[0], months.at(-1), 'bad cells', bad, 'bytes', fs.statSync(outFile).size);
// quick story stats
const tot = t => val.reduce((a, v) => a + v[t], 0);
const idx = ym => months.indexOf(ym);
for (const ym of [192607, 192909, 193206, 194506, 196206, 196212, 197212, 197301, 198001, 199912, 200003, 200210, 200710, 200903, 202002, 202112, 202609]) {
  const t = idx(ym); const T = tot(t);
  const top = ind.map((k, j) => [k, val[j][t] / T]).sort((a, b) => b[1] - a[1]).slice(0, 5).map(([k, s]) => k + ' ' + (s * 100).toFixed(1)).join(', ');
  const real = T * cpiArr.at(-1) / cpiArr[t];
  console.log(ym, 'total $' + (T / 1e6).toFixed(3) + 'T', 'real2026 $' + (real / 1e6).toFixed(2) + 'T', 'firms', firms.reduce((a, f) => a + f[t], 0), '|', top);
}
