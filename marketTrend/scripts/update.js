// Command-line update: fetch new trading days, keep one month, rebuild the page files.
// usage: node scripts/update.js [--force]
const market = require('./lib/market'), page = require('./lib/page');
const force = process.argv.includes('--force');

market.update({ force, onProgress: p => { if (p.total) { if (p.done === p.total || p.done % 100 === 0) console.log(`${p.message}: ${p.done}/${p.total}`); } else console.log(p.message); } })
  .then(r => {
    if (r.upToDate) console.log(`Up to date: latest close ${r.latest}`);
    else console.log(`Window ${r.first} to ${r.latest}, ${r.companies} companies. Added: ${r.added.join(', ') || 'none'}`), r.notes.forEach(n => console.log('note:', n));
    page.write();
    console.log('Wrote month-delta.html and build/month-delta.body.html');
  })
  .catch(e => { console.error('Update failed:', e.message); process.exit(1); });
