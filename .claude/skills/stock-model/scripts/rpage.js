// node rpage.js <accessionNoDashes> <R#.htm> [outfile]  -> fetch SEC R page and print table rows as " | " separated text
const https = require('https');
const fs = require('fs');
const [acc, page, out] = process.argv.slice(2);
const url = `https://www.sec.gov/Archives/edgar/data/${process.env.CIK||"24741"}/${acc}/${page}`;
https.get(url, { headers: { 'User-Agent': 'Independent Research Model research-contact@example.org' } }, res => {
  let s = ''; res.setEncoding('utf8');
  res.on('data', d => s += d).on('end', () => {
    const dec = t => t.replace(/<br\s*\/?>/gi, ' ').replace(/<[^>]+>/g, '').replace(/&nbsp;|&#160;/g, ' ').replace(/&amp;/g, '&').replace(/&#8217;|&rsquo;/g, "'").replace(/&#8212;|&mdash;/g, '-').replace(/&#8211;/g, '-').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/\s+/g, ' ').trim();
    const rows = [];
    const trRe = /<tr[^>]*>([\s\S]*?)<\/tr>/gi; let m;
    while ((m = trRe.exec(s))) {
      const cells = []; const tdRe = /<t[hd][^>]*>([\s\S]*?)<\/t[hd]>/gi; let c;
      while ((c = tdRe.exec(m[1]))) cells.push(dec(c[1]));
      const line = cells.filter(x => x !== '').join(' | ');
      if (line) rows.push(line);
    }
    const txt = `URL: ${url}\n` + rows.join('\n');
    if (out) { fs.writeFileSync(out, txt); console.log('wrote', out, rows.length, 'rows'); } else console.log(txt);
  });
}).on('error', e => console.error(e));
