// node list_facts.js file.json <regex>  -> lists all taxonomies' tags matching regex with annual 10-K values 2020+
const fs = require('fs');
const j = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
const re = new RegExp(process.argv[3], 'i');
for (const tax of Object.keys(j.facts)) {
  for (const tag of Object.keys(j.facts[tax])) {
    if (!re.test(tag)) continue;
    const f = j.facts[tax][tag];
    for (const unit of Object.keys(f.units)) {
      const out = {};
      for (const v of f.units[unit]) {
        if (!v.form || !v.form.startsWith('10-K')) continue;
        if (v.start) { const d = (new Date(v.end) - new Date(v.start)) / 864e5; if (d < 350 || d > 380) continue; }
        const yr = v.end.slice(0, 4); if (+yr < 2020) continue;
        if (!out[yr] || v.filed > out[yr].filed) out[yr] = { val: v.val, filed: v.filed };
      }
      const yrs = Object.keys(out).sort(); if (!yrs.length) continue;
      console.log((tax + ':' + tag).padEnd(85), unit.padEnd(6), yrs.map(y => y + ':' + (Math.abs(out[y].val) >= 1e5 ? Math.round(out[y].val / 1e6) : out[y].val)).join('  '));
    }
  }
}
