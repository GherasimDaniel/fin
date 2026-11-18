// Local helper for the month delta.
// usage: node serve.js   then open http://localhost:8787
// Each time the page opens it asks this helper for new trading days; the helper downloads them,
// keeps one month of history and the page reloads with the new data.
const http = require('http'), { exec } = require('child_process');
const market = require('./scripts/lib/market');
const page = require('./scripts/lib/page');

const PORT = +process.env.PORT || 8787;
let job = { state: 'idle' }, running = null, lastCheck = 0;

// keep month-delta.html and build/month-delta.body.html in step with data/month.js
function writeStatic() {
  try { page.write(); } catch (e) { console.error('could not rewrite month-delta.html:', e.message); }
}

function startUpdate() {
  if (running) return running;
  // a page reload right after an update does not need a second check
  if (job.state === 'done' && Date.now() - lastCheck < 60000) return Promise.resolve();
  job = { state: 'running', message: 'Checking for new trading days', done: 0, total: 0 };
  running = market.update({ onProgress: p => { job = { state: 'running', message: p.message, done: p.done || 0, total: p.total || 0 }; } })
    .then(r => {
      job = { state: 'done', added: r.added, latest: r.latest };
      lastCheck = Date.now();
      writeStatic();
      console.log(new Date().toLocaleTimeString(), r.added.length ? `added ${r.added.join(', ')}; window now ${r.first} to ${r.latest}` : `up to date (latest close ${r.latest})`);
    })
    .catch(e => { job = { state: 'error', message: e.message }; console.error('update failed:', e.message); })
    .finally(() => { running = null; });
  return running;
}

const server = http.createServer((req, res) => {
  const send = (code, type, body) => { res.writeHead(code, { 'Content-Type': type, 'Cache-Control': 'no-store' }); res.end(body); };
  try {
    if (req.method === 'GET' && (req.url === '/' || req.url.startsWith('/?'))) return send(200, 'text/html; charset=utf-8', page.full({ live: true }));
    if (req.method === 'POST' && req.url === '/api/update') { startUpdate(); return send(202, 'application/json', JSON.stringify(job)); }
    if (req.method === 'GET' && req.url === '/api/status') return send(200, 'application/json', JSON.stringify(job));
    send(404, 'text/plain', 'Not found');
  } catch (e) {
    send(500, 'text/plain', 'Error: ' + e.message);
  }
});
writeStatic();
server.listen(PORT, '127.0.0.1', () => {
  const url = `http://localhost:${PORT}`;
  console.log(`Month delta running at ${url}  (Ctrl+C to stop)`);
  if (!process.env.NO_OPEN) exec(process.platform === 'win32' ? `start "" "${url}"` : `open "${url}"`);
});
