// RM and cluster performance report (their own numbers, not staff ratings): sales target vs achievement, conversion,
// attrition and operational scorecard, rolled up from the same store data the dashboard uses (build/out/data.json).
// usage: node build/export-rm-performance.js "<RM name>" ["<RM name>" ...]     (no names = every RM)
const fs = require('fs'), path = require('path'), X = require('xlsx');
const OUT = 'C:/Users/Lenovo/OneDrive/Desktop/Category Q1/26-27 iteamwise report H1/RM reports';
const D = JSON.parse(fs.readFileSync(path.join(__dirname, 'out/data.json')));
const M = D.months, ML = m => new Date(m + '-01').toLocaleString('en-GB', { month: 'short' }).slice(0, 3);
const tc = s => String(s || '').toLowerCase().replace(/\b\w/g, c => c.toUpperCase()).trim();
const sum = (a, f) => a.reduce((x, y) => x + (f(y) || 0), 0);
const div = (a, b) => b ? a / b : null;
const r4 = v => v == null || !isFinite(v) ? '' : +v.toFixed(4);

// every real store with its RM / CM (attrition sheet mapping first, then the target sheet)
const stores = Object.entries(D.stores).filter(([, s]) => s.code).map(([k, s]) => ({ k, s, rm: tc(s.rmA || s.rm), cm: tc(s.cmA || s.cm) }));
const want = process.argv.slice(2);
const rms = [...new Set(stores.map(x => x.rm))].filter(r => r && (!want.length || want.some(w => r.toLowerCase().includes(w.toLowerCase())))).sort();

// roll a group of stores up for a set of months (same rules as the dashboard: e-commerce excluded; months without a target skipped)
function roll(group, months) {
  const cells = group.flatMap(x => months.map(m => ({ ...(x.s.m[m] || {}), m })));
  const sm = cells.filter(c => c.target > 0), cv = cells.filter(c => c.wiT > 0 && c.cvT > 0 && c.wi > 0);
  const sales = sum(sm, c => c.netI), target = sum(sm, c => c.target);
  const wi = sum(cv, c => c.wi), bills = sum(cv, c => c.cBills), tB = sum(cv, c => c.wiT * c.cvT), tW = sum(cv, c => c.wiT);
  const conv = div(bills, wi), convT = div(tB, tW);
  const scores = group.flatMap(x => ((x.s.inp || {}).scores || []).filter(z => months.includes(z.m)).map(z => z.v));
  return { sales, target, ach: div(sales, target), wi, bills, conv, convT, convAch: conv != null && convT ? conv / convT : null,
    ops: scores.length ? scores.reduce((a, b) => a + b, 0) / scores.length / 100 : null };
}
function attr(group) {
  const g = group.map(x => x.s.inp || {}).filter(i => i.left != null);
  const open = sum(g, i => i.open), join = sum(g, i => i.joiners), exits = sum(g, i => i.left), close = sum(g, i => i.close);
  return { open, join, exits, close, pct: div(exits, (open + close) / 2) };   // same formula as the attrition report: exits ÷ average headcount
}
function shrink(group) {
  const a = group.flatMap(x => (x.s.inp || {}).audits || []);
  const sale = sum(a, z => z.sale), val = sum(a, z => z.value);
  return sale ? Math.max(0, -val) / sale * 100 : null;
}
const levels = [];   // rows in order: RM, its clusters, next RM...
for (const rm of rms) {
  const g = stores.filter(x => x.rm === rm);
  levels.push({ level: 'RM', name: rm, rm, g });
  for (const cm of [...new Set(g.map(x => x.cm))].sort()) levels.push({ level: 'Cluster', name: cm, rm, g: g.filter(x => x.cm === cm) });
}
const fmt = (ws, aoa, from, pctCols, numCols) => {
  for (let r = from; r < aoa.length; r++) {
    for (const c of pctCols) { const a = X.utils.encode_cell({ r, c }); if (ws[a] && typeof ws[a].v === 'number') ws[a].z = '0.0%'; }
    for (const c of numCols) { const a = X.utils.encode_cell({ r, c }); if (ws[a] && typeof ws[a].v === 'number') ws[a].z = '#,##0'; }
  }
};
const label = l => l.level === 'RM' ? `RM: ${l.name}` : `   Cluster: ${l.name}`;
const wb = X.utils.book_new();
const stamp = `H1 FY 2026-27 (Apr–Sep 2026) · sales exclude e-commerce · generated ${new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' })}`;

// 1. Summary: one row per RM / cluster, all four KPI groups
{
  const aoa = [['RM and cluster performance: target vs achievement, conversion, attrition, operational scorecard'], [stamp], [],
    ['', '', 'SALES TARGET VS ACHIEVEMENT', '', '', 'CONVERSION', '', '', '', '', 'ATTRITION (Apr–Sep)', '', '', '', '', 'OPERATIONAL SCORECARD', '', '', '', '', '', 'SHRINKAGE'],
    ['RM / Cluster', 'Stores', 'Target', 'Achievement', 'Ach %', 'Walk-ins', 'Bills', 'Conversion', 'Conv. target', 'Conv. ach %', 'Opening HC', 'Joiners', 'Exits', 'Closing HC', 'Attrition %',
      ...M.slice(0, 5).map(ML), 'H1 avg', 'Shrinkage %']];
  for (const l of levels) {
    const h = roll(l.g, M), a = attr(l.g), sh = shrink(l.g);
    aoa.push([label(l), l.g.length, h.target, h.sales, r4(h.ach), h.wi, h.bills, r4(h.conv), r4(h.convT), r4(h.convAch), a.open, a.join, a.exits, a.close, r4(a.pct),
      ...M.slice(0, 5).map(m => r4(roll(l.g, [m]).ops)), r4(h.ops), sh == null ? '' : +sh.toFixed(3)]);
  }
  aoa.push([], ['Notes'], ['Sales: store Item Net Amt excluding e-commerce ÷ store target, months with a target only.'],
    ['Conversion: bills ÷ walk-ins (daily conversion report); target = walk-in-target weighted CV%.'],
    ['Attrition %: exits ÷ average of opening and closing headcount (same as the Retail Attrition report).'],
    ['Operational scorecard: average of the stores\u2019 monthly scores (file has Apr–Aug; no September).'],
    ['Shrinkage %: net loss value ÷ store sale across the Q1 and Q2 stock audits.']);
  const ws = X.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [{ wch: 34 }, { wch: 8 }, { wch: 15 }, { wch: 15 }, ...Array(18).fill({ wch: 11 })];
  ws['!merges'] = [{ s: { r: 3, c: 2 }, e: { r: 3, c: 4 } }, { s: { r: 3, c: 5 }, e: { r: 3, c: 9 } }, { s: { r: 3, c: 10 }, e: { r: 3, c: 14 } }, { s: { r: 3, c: 15 }, e: { r: 3, c: 20 } }];
  fmt(ws, aoa, 5, [4, 7, 8, 9, 14, 15, 16, 17, 18, 19, 20], [2, 3, 5, 6]);
  X.utils.book_append_sheet(wb, ws, 'Summary');
}
// 2. Sales by month
{
  const aoa = [['Sales target vs achievement by month (excluding e-commerce)'], [stamp], [], ['', ...M.flatMap(m => [ML(m), '', '']), 'H1', '', ''],
    ['RM / Cluster', ...[...M, 'H1'].flatMap(() => ['Target', 'Achievement', 'Ach %'])]];
  for (const l of levels) aoa.push([label(l), ...[...M.map(m => roll(l.g, [m])), roll(l.g, M)].flatMap(x => [x.target, x.sales, r4(x.ach)])]);
  const ws = X.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [{ wch: 34 }, ...Array(21).fill({ wch: 13 })];
  ws['!merges'] = [...M, 'H1'].map((_, i) => ({ s: { r: 3, c: 1 + i * 3 }, e: { r: 3, c: 3 + i * 3 } }));
  fmt(ws, aoa, 5, [...M, 'H1'].map((_, i) => 3 + i * 3), [...M, 'H1'].flatMap((_, i) => [1 + i * 3, 2 + i * 3]));
  X.utils.book_append_sheet(wb, ws, 'Sales by month');
}
// 3. Conversion by month
{
  const aoa = [['Conversion by month'], [stamp], [], ['', ...M.flatMap(m => [ML(m), '', '', '', '']), 'H1', '', '', '', ''],
    ['RM / Cluster', ...[...M, 'H1'].flatMap(() => ['Walk-ins', 'Bills', 'Conversion', 'Target', 'Ach %'])]];
  for (const l of levels) aoa.push([label(l), ...[...M.map(m => roll(l.g, [m])), roll(l.g, M)].flatMap(x => [x.wi, x.bills, r4(x.conv), r4(x.convT), r4(x.convAch)])]);
  const ws = X.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [{ wch: 34 }, ...Array(35).fill({ wch: 10 })];
  ws['!merges'] = [...M, 'H1'].map((_, i) => ({ s: { r: 3, c: 1 + i * 5 }, e: { r: 3, c: 5 + i * 5 } }));
  fmt(ws, aoa, 5, [...M, 'H1'].flatMap((_, i) => [3 + i * 5, 4 + i * 5, 5 + i * 5]), [...M, 'H1'].flatMap((_, i) => [1 + i * 5, 2 + i * 5]));
  X.utils.book_append_sheet(wb, ws, 'Conversion by month');
}
// 4. Store wise
{
  const aoa = [['Store-wise detail'], [stamp], [],
    ['Regional Manager', 'Cluster Manager', 'Store', 'Code', 'Target', 'Achievement', 'Ach %', 'Walk-ins', 'Bills', 'Conversion', 'Conv. target', 'Conv. ach %', 'Opening HC', 'Joiners', 'Exits', 'Closing HC', 'Attrition %',
      ...M.slice(0, 5).map(m => 'Ops ' + ML(m)), 'Ops H1 avg', 'Shrinkage %', ...M.map(m => 'Sales ach ' + ML(m))]];
  for (const l of levels.filter(x => x.level === 'Cluster')) for (const x of l.g.slice().sort((a, b) => a.k.localeCompare(b.k))) {
    const h = roll([x], M), a = attr([x]), sh = shrink([x]), i = x.s.inp || {};
    aoa.push([l.rm, l.name, x.k, x.s.code, h.target || '', h.sales || '', r4(h.ach), h.wi || '', h.bills || '', r4(h.conv), r4(h.convT), r4(h.convAch),
      i.left != null ? a.open : '', i.left != null ? a.join : '', i.left != null ? a.exits : '', i.left != null ? a.close : '', i.left != null ? r4(a.pct) : '',
      ...M.slice(0, 5).map(m => r4(roll([x], [m]).ops)), r4(h.ops), sh == null ? '' : +sh.toFixed(3), ...M.map(m => r4(roll([x], [m]).ach))]);
  }
  const ws = X.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [{ wch: 18 }, { wch: 20 }, { wch: 34 }, { wch: 9 }, { wch: 14 }, { wch: 14 }, ...Array(24).fill({ wch: 11 })];
  ws['!autofilter'] = { ref: `A4:AD${aoa.length}` };
  fmt(ws, aoa, 4, [6, 9, 10, 11, 16, 17, 18, 19, 20, 21, 22, 24, 25, 26, 27, 28, 29], [4, 5, 7, 8]);
  X.utils.book_append_sheet(wb, ws, 'Store wise');
}
fs.mkdirSync(OUT, { recursive: true });
const file = path.join(OUT, `PMS H1 - RM & Cluster performance - ${rms.join(', ').replace(/[\\/:*?"<>|]/g, '')}.xlsx`);
X.writeFile(wb, file);
console.log('written', file);
for (const l of levels) { const h = roll(l.g, M), a = attr(l.g); console.log(`${label(l).padEnd(34)} stores ${String(l.g.length).padStart(2)} | sales ${(h.sales / 1e7).toFixed(2)} of ${(h.target / 1e7).toFixed(2)} Cr = ${(h.ach * 100).toFixed(1)}% | conv ${(h.conv * 100).toFixed(1)}% vs ${(h.convT * 100).toFixed(1)}% = ${(h.convAch * 100).toFixed(0)}% | attrition ${a.exits} exits, ${(a.pct * 100).toFixed(1)}% | ops ${(h.ops * 100).toFixed(1)}%`); }
