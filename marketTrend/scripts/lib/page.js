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
module.exports = { body, full };
