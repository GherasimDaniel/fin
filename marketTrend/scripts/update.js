// Command-line update: fetch new trading days, keep one month, rebuild the page files.
// usage: node scripts/update.js [--force]
const fs = require('fs'), path = require('path');
const market = require('./lib/market'), page = require('./lib/page');
const force = process.argv.includes('--force');

market.update({ force, onProgress: p => { if (p.total) { if (p.done === p.total || p.done % 100 === 0) console.log(`${p.message}: ${p.done}/${p.total}`); } else console.log(p.message); } })
  .then(r => {
    if (r.upToDate) console.log(`Up to date: latest close ${r.latest}`);
    else console.log(`Window ${r.first} to ${r.latest}, ${r.companies} companies. Added: ${r.added.join(', ') || 'none'}`), r.notes.forEach(n => console.log('note:', n));
    const root = market.ROOT;
    fs.mkdirSync(path.join(root, 'build'), { recursive: true });
    const body = page.body({ live: false });
    for (const m of body.matchAll(/<script>([\s\S]*?)<\/script>/g)) new Function(m[1]);
    fs.writeFileSync(path.join(root, 'build', 'month-delta.body.html'), body);
    fs.writeFileSync(path.join(root, 'month-delta.html'), page.full({ live: false }));
    console.log('Wrote month-delta.html and build/month-delta.body.html');
  })
  .catch(e => { console.error('Update failed:', e.message); process.exit(1); });
