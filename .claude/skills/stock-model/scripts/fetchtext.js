// node fetchtext.js <url> [maxChars] -> fetch SEC doc and print plain text
const https = require('https');
const [url, max] = process.argv.slice(2);
https.get(url, { headers: { 'User-Agent': 'Independent Research Model research-contact@example.org' } }, res => {
  let s = ''; res.setEncoding('utf8');
  res.on('data', d => s += d).on('end', () => {
    let t = s.replace(/<script[\s\S]*?<\/script>/gi, '').replace(/<style[\s\S]*?<\/style>/gi, '')
      .replace(/<\/(p|div|tr|li|h\d|table)>/gi, '\n').replace(/<br\s*\/?>/gi, '\n').replace(/<\/t[dh]>/gi, ' | ')
      .replace(/<[^>]+>/g, '').replace(/&nbsp;|&#160;|&#xa0;/gi, ' ').replace(/&amp;/g, '&').replace(/&#8217;|&rsquo;|&#x2019;/gi, "'")
      .replace(/&#8220;|&#8221;|&ldquo;|&rdquo;|&#x201c;|&#x201d;/gi, '"').replace(/&#8212;|&#8211;|&mdash;|&ndash;|&#x2014;|&#x2013;/gi, '-').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&#\d+;/g, ' ')
      .replace(/[ \t]+/g, ' ').replace(/\n\s*\n+/g, '\n').trim();
    console.log(max ? t.slice(0, +max) : t);
  });
}).on('error', e => console.error(e));
