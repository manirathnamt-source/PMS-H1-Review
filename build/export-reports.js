// Writes RM-wise and cluster-wise Excel reports using the dashboard's own code (index.html), so numbers match the site:
// calibration format sheets (SM / ASM / SSA-SA) + a Summary sheet. Includes saved KPI changes and reviews (live from the change store).
// usage: node build/export-reports.js "<RM name>" ["<RM name>" ...]     (no names = every RM)
const fs = require('fs'), path = require('path'), X = require('xlsx');
const ROOT = path.join(__dirname, '..');
const OUT = 'C:/Users/Lenovo/OneDrive/Desktop/Category Q1/26-27 iteamwise report H1/RM reports';
const html = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');
const script = html.slice(html.indexOf('<script>') + 8, html.lastIndexOf('</script>'));

// minimal browser stand-ins: the page code only needs them to load; the calculations use none of them
const el = new Proxy(function () {}, { get: (t, k) => k === 'value' ? '' : k === 'checked' ? true : k === 'classList' ? { add() {}, remove() {}, toggle() {} } : k === 'style' || k === 'dataset' ? {} : k === Symbol.toPrimitive ? () => '' : el, set: () => true, apply: () => el });
const document = { querySelector: () => el, querySelectorAll: () => [], addEventListener() {}, createElement: () => el, documentElement: { dataset: {}, style: {} }, head: { appendChild() {} }, body: { appendChild() {} } };
const window = {}, location = { reload() {} }, matchMedia = () => ({ matches: false });

const runner = `
;return (async () => {
  D = __data;
  API_KEY = await sha256hex('pms-h1-key|' + __pass);
  compute(); await loadChanges(); compute();
  return { PEOPLE, storeH1, exportCalibration, rOf, hasOvr, needsReview, tc, sync: SYNC, changes: CHANGES.length };
})();`;
const data = JSON.parse(fs.readFileSync(path.join(__dirname, 'out/data.json')));
const pass = fs.readFileSync(path.join(__dirname, 'passcode.txt'), 'utf8').trim();

(async () => {
  const api = await new Function('document', 'window', 'location', 'matchMedia', '__data', '__pass', script + runner)(document, window, location, matchMedia, data, pass);
  const { PEOPLE, storeH1, exportCalibration, rOf, hasOvr, needsReview, tc } = api;
  console.log('change store:', api.sync, '| saved changes + reviews:', api.changes);
  const elig = PEOPLE.filter(p => p.eligible);
  const want = process.argv.slice(2), rms = [...new Set(elig.map(p => p.rmN))].filter(r => r && (!want.length || want.some(w => r.toLowerCase().includes(w.toLowerCase()))));
  fs.mkdirSync(OUT, { recursive: true });
  const avg = a => a.length ? +(a.reduce((x, y) => x + y, 0) / a.length).toFixed(2) : '';
  const pct = v => v == null || !isFinite(v) ? '' : +(v / 100).toFixed(4);
  const stats = l => {
    const rated = l.filter(p => rOf(p) != null), c = r => l.filter(p => rOf(p) === r).length;
    return [l.length, l.filter(p => p.role === 'SM').length, l.filter(p => p.role === 'ASM').length, l.filter(p => p.role === 'SSA').length,
      c(4), c(3), c(2), c(1), l.length - rated.length, avg(rated.map(p => p.score ?? p.prov)),
      l.filter(p => needsReview(p) && p.review).length, l.filter(p => needsReview(p) && !p.review).length, l.filter(hasOvr).length];
  };
  const STAT_HEAD = ['Staff (eligible)', 'SM', 'ASM', 'SA / SSA', 'Rating 4', 'Rating 3', 'Rating 2', 'Rating 1', 'Not rated', 'Avg score', 'Reviewed', 'Review pending', 'Calibrated'];
  const summary = (title, l, byCluster) => {
    const aoa = [[title], [`H1 FY 2026-27 (Apr–Sep 2026) · eligible staff · generated ${new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' })}`], []];
    if (byCluster) {
      aoa.push(['CLUSTER WISE'], ['Cluster Manager', 'Stores', ...STAT_HEAD]);
      for (const cm of [...new Set(l.map(p => p.cmN))].sort()) { const g = l.filter(p => p.cmN === cm); aoa.push([cm, new Set(g.map(p => p.outlet)).size, ...stats(g)]); }
      aoa.push(['Total', new Set(l.map(p => p.outlet)).size, ...stats(l)], []);
    }
    aoa.push(['STORE WISE'], ['Store', 'Cluster Manager', 'H1 sales (excl. e-com)', 'Target', 'Sales ach.', 'Conversion', 'Conv. target', 'Conv. ach.', 'Ops score', 'Shrinkage %', 'Exits', ...STAT_HEAD]);
    const first = aoa.length;
    for (const k of [...new Set(l.map(p => p.outlet))].sort((a, b) => (l.find(p => p.outlet === a).cmN + a).localeCompare(l.find(p => p.outlet === b).cmN + b))) {
      const g = l.filter(p => p.outlet === k), s = storeH1(k);
      aoa.push([k, g[0].cmN, s.sales || '', s.target || '', pct(s.salesPct), pct(s.conv), pct(s.convT), pct(s.convPct), pct(s.audit), s.shrink == null ? '' : +s.shrink.toFixed(3), s.left ?? '', ...stats(g)]);
    }
    const ws = X.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{ wch: 34 }, { wch: 20 }, ...Array(22).fill({ wch: 13 })];
    for (let r = first; r < aoa.length; r++) {
      for (const c of [4, 5, 6, 7, 8]) { const a = X.utils.encode_cell({ r, c }); if (ws[a] && typeof ws[a].v === 'number') ws[a].z = '0.0%'; }
      for (const c of [2, 3]) { const a = X.utils.encode_cell({ r, c }); if (ws[a] && typeof ws[a].v === 'number') ws[a].z = '#,##0'; }
    }
    return ws;
  };
  const write = (file, title, l, byCluster) => {
    let wb; exportCalibration({ ...X, writeFile: w => { wb = w; } }, l, 'x');
    // Summary first, then the HR-format sheets
    const out = X.utils.book_new();
    X.utils.book_append_sheet(out, summary(title, l, byCluster), 'Summary');
    for (const n of wb.SheetNames) X.utils.book_append_sheet(out, wb.Sheets[n], n);
    X.writeFile(out, path.join(OUT, file));
    console.log(' ', file, '|', l.length, 'staff');
  };
  const safe = s => s.replace(/[\\/:*?"<>|]/g, '').trim();
  for (const rm of rms) {
    const l = elig.filter(p => p.rmN === rm);
    console.log('RM', rm);
    write(`PMS H1 - RM ${safe(rm)}.xlsx`, `RM report: ${rm}`, l, true);
    for (const cm of [...new Set(l.map(p => p.cmN))].sort()) write(`PMS H1 - RM ${safe(rm)} - Cluster ${safe(cm)}.xlsx`, `Cluster report: ${cm} (RM ${rm})`, l.filter(p => p.cmN === cm), false);
  }
  console.log('saved in', OUT);
})().catch(e => { console.error('FAILED', e); process.exit(1); });
