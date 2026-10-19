// Saves the RM / CM KPI rating changes made on the dashboard (live from the change store) to build/out/mgr-changes.json,
// so the goal sheets can show them as the Manager's Rating. Run before export-goal-sheets.js.
const fs = require('fs'), path = require('path'), X = require('xlsx');
const ROOT = path.join(__dirname, '..');
const html = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');
const script = html.slice(html.indexOf('<script>') + 8, html.lastIndexOf('</script>'));

// minimal browser stand-ins: the page code only needs them to load; the calculations use none of them
const el = new Proxy(function () {}, { get: (t, k) => k === 'value' ? '' : k === 'checked' ? true : k === 'classList' ? { add() {}, remove() {}, toggle() {} } : k === 'style' || k === 'dataset' ? {} : k === Symbol.toPrimitive ? () => '' : el, set: () => true, apply: () => el });
const document = { querySelector: () => el, querySelectorAll: () => [], addEventListener() {}, createElement: () => el, documentElement: { dataset: {}, style: {} }, head: { appendChild() {} }, body: { appendChild() {} } };
const window = {}, location = { reload() {} }, matchMedia = () => ({ matches: false });

const data = JSON.parse(fs.readFileSync(path.join(__dirname, 'out/data.json')));
const pass = fs.readFileSync(path.join(__dirname, 'passcode.txt'), 'utf8').trim();
const run = `;return (async () => { D = __data; API_KEY = await sha256hex('pms-h1-key|' + __pass); compute(); await loadChanges(); compute();
  return mgrList().map(l => mgrRate(l)).flatMap(r => r.rows.filter(x => x.ovr).map(x => ({ id: r.id, level: r.level, name: r.name, rm: r.rm, kpi: x.k, sys: x.sys, rating: x.r, reason: String(x.ovr.reason), by: x.ovr.by, ts: x.ovr.ts }))); })();`;
(async () => {
  const out = await new Function('document', 'window', 'location', 'matchMedia', '__data', '__pass', script + run)(document, window, location, matchMedia, data, pass);
  fs.writeFileSync(path.join(__dirname, 'out/mgr-changes.json'), JSON.stringify(out, null, 1));
  console.log(out.length, 'RM/CM rating changes saved to build/out/mgr-changes.json');
})();
