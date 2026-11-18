// Puts data/month.js into scripts/template.html.
// live = true adds window.LIVE so the page asks the local helper for new trading days when it opens.
const fs = require('fs'), path = require('path');
const ROOT = path.join(__dirname, '..', '..');

function body({ live = false } = {}) {
  const tpl = fs.readFileSync(path.join(ROOT, 'scripts', 'template.html'), 'utf8');
  const data = fs.readFileSync(path.join(ROOT, 'data', 'month.js'), 'utf8').trim();
  return tpl.replace('/*LIVE*/', () => (live ? 'window.LIVE = true;' : '')).replace('/*DATA*/', () => data);
}
function full(opts) {
  return '<!doctype html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n</head>\n<body>\n' + body(opts) + '\n</body>\n</html>\n';
}
// Rewrites the static copies (month-delta.html and build/month-delta.body.html) from the current data.
function write() {
  const b = body({ live: false });
  for (const m of b.matchAll(/<script>([\s\S]*?)<\/script>/g)) new Function(m[1]); // syntax check before overwriting
  fs.mkdirSync(path.join(ROOT, 'build'), { recursive: true });
  fs.writeFileSync(path.join(ROOT, 'build', 'month-delta.body.html'), b);
  fs.writeFileSync(path.join(ROOT, 'month-delta.html'), full({ live: false }));
}
module.exports = { body, full, write };
