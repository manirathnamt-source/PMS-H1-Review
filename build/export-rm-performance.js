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

/* ---- RM and CM ratings (Retail Goal FY 2025-26 KPI sheets: "CM KPI" and "RM KPI" with their Rating tabs) ---- */
const rateAch = p => p == null ? null : p <= 70 ? 1 : p <= 90 ? 2 : p <= 100 ? 3 : 4;                 // sales / conversion achievement %
const rateAudit = p => p == null ? null : p < 80 ? 1 : p < 90 ? 2 : p < 95 ? 3 : 4;                    // store ops scorecard %
const rateShrink = p => p == null ? null : p > 0.125 ? 1 : p > 0.075 ? 2 : p >= 0.055 ? 3 : 4;         // store shrinkage %
const rateStores = p => p == null ? null : p < 50 ? 1 : p < 65 ? 2 : p <= 80 ? 3 : 4;                  // CM: % of stores at rating 3 or better
const rateAttr = p => p == null ? null : p > 40 ? 0 : p > 30 ? 1 : p > 25 ? 2 : p >= 20 ? 3 : 4;       // attrition %
const finalOf = s => s == null ? null : Math.min(4, Math.max(1, Math.round(+s.toFixed(4))));
const SHRINK_EXEMPT = ['BLR - COMMERCIAL STREET'];   // "Except commercial street store" on the KPI sheets
// one store's H1 KPI ratings (the same rules the dashboard applies to stores)
// shrinkage is rated quarter by quarter (that quarter's net loss ÷ store sale), then the quarter ratings are averaged
const shrinkQ = x => {
  const g = {}; for (const a of (x.s.inp || {}).audits || []) (g[a.q || 'Audit'] = g[a.q || 'Audit'] || []).push(a);
  const r = Object.values(g).map(a => { const sale = sum(a, z => z.sale), val = sum(a, z => z.value); return rateShrink(sale ? Math.max(0, -val) / sale * 100 : null); }).filter(v => v != null);
  return r.length ? +(r.reduce((p, v) => p + v, 0) / r.length).toFixed(2) : rateShrink((x.s.inp || {}).shrink ?? null);
};
const storeRt = x => {
  const h = roll([x], M);
  return { conv: rateAch(h.convAch == null ? null : h.convAch * 100), ops: rateAudit(h.ops == null ? null : h.ops * 100), shrink: SHRINK_EXEMPT.includes(x.k) ? null : shrinkQ(x) };
};
// share of a group's stores that reach rating 3+ on a KPI (stores without data for it are left out)
const share = (g, k) => { const r = g.map(x => storeRt(x)[k]).filter(v => v != null); return r.length ? { ok: r.filter(v => v >= 3).length, n: r.length, pct: r.filter(v => v >= 3).length / r.length * 100 } : null; };
// ops scorecard, month by month: how many of the group's stores achieved it (90% or above = Green zone) out of the stores scored that month
const OPS_MONTHS = M.filter(m => stores.some(x => ((x.s.inp || {}).scores || []).some(z => z.m === m)));
const opsMonth = (g, m) => { const v = g.map(x => (((x.s.inp || {}).scores || []).find(z => z.m === m) || {}).v).filter(z => z != null); return { ok: v.filter(z => z >= 90).length, n: v.length }; };
const opsAll = g => { const a = OPS_MONTHS.map(m => opsMonth(g, m)); const ok = a.reduce((p, x) => p + x.ok, 0), n = a.reduce((p, x) => p + x.n, 0); return n ? { ok, n, pct: ok / n * 100, byMonth: a } : null; };
function rating(l) {
  const h = roll(l.g, M), a = attr(l.g), aPct = a.pct == null ? null : a.pct * 100, rows = [];
  const add = (kpi, w, basis, ach, r) => rows.push({ kpi, w, basis, ach, r });
  add('Sales Target', l.level === 'RM' ? .40 : .35, `₹${(h.sales / 1e7).toFixed(2)} Cr of ₹${(h.target / 1e7).toFixed(2)} Cr`, h.ach == null ? '' : (h.ach * 100).toFixed(1) + '%', rateAch(h.ach == null ? null : h.ach * 100));
  if (l.level === 'RM') {
    add('Selling Effectiveness (Conversion)', .25, `${(h.conv * 100).toFixed(1)}% vs target ${(h.convT * 100).toFixed(1)}%`, (h.convAch * 100).toFixed(1) + '%', rateAch(h.convAch * 100));
    add('Process Improvement & Initiative Impact', .15, 'Manager assessment (min. 2 impactful initiatives a year)', 'Pending', null);
    add('Salary to Sales Ratio', .10, 'Target 10%; salary data not available', 'Pending', null);
    add('Team Development & Retention', .10, `${a.exits} exits, attrition ${aPct == null ? '–' : aPct.toFixed(1) + '%'} (promotion % not available)`, aPct == null ? '' : aPct.toFixed(1) + '%', rateAttr(aPct));
  } else {
    const c = share(l.g, 'conv'), o = opsAll(l.g), s = share(l.g, 'shrink');
    const sh = (x, what) => x ? `${x.ok} of ${x.n} stores at rating 3+ on ${what}` : 'No data';
    add('Selling Effectiveness (Conversion)', .35, sh(c, 'conversion'), c ? c.pct.toFixed(0) + '% of stores' : '', rateStores(c && c.pct));
    add('Retail Operational Scorecard', .15, o ? OPS_MONTHS.map((m, i) => `${ML(m)} ${o.byMonth[i].ok}/${o.byMonth[i].n}`).join(', ') + ` = ${o.ok} of ${o.n} store-months achieved (90%+)` : 'No data', o ? o.pct.toFixed(0) + '% achieved' : '', rateStores(o && o.pct));
    add('Team Development', .10, `${a.exits} exits, attrition ${aPct == null ? '–' : aPct.toFixed(1) + '%'} (promotion % not available)`, aPct == null ? '' : aPct.toFixed(1) + '%', rateAttr(aPct));
    add('Shrinkage Control', .05, sh(s, 'shrinkage') + (l.g.some(x => SHRINK_EXEMPT.includes(x.k)) ? ' (Commercial Street excluded)' : ''), s ? s.pct.toFixed(0) + '% of stores' : '', rateStores(s && s.pct));
  }
  const have = rows.filter(x => x.r != null), sw = have.reduce((p, x) => p + x.w, 0), sc = have.reduce((p, x) => p + x.w * x.r, 0);
  const full = have.length === rows.length;
  return { rows, score: sw ? (full ? sc : sc / sw) : null, full, pending: rows.filter(x => x.r == null).map(x => x.kpi) };
}

// 0. Ratings: one block per RM / CM with every KPI, its basis, rating and weighted points
{
  const aoa = [['RM and Cluster Manager ratings, H1 FY 2026-27'], [stamp], [],
    ['Name', 'Role', 'Stores', 'Weighted score', 'Rating', 'Status', 'Pending KPIs']];
  const res = levels.map(l => ({ l, R: rating(l) }));
  for (const { l, R } of res) aoa.push([l.name, l.level === 'RM' ? 'Regional Manager' : `Cluster Manager (RM ${l.rm})`, l.g.length, R.score == null ? '' : +R.score.toFixed(2), finalOf(R.score) ?? '', R.full ? 'Final' : 'Provisional', R.pending.join('; ')]);
  aoa.push([], ['KPI DETAIL'], ['Name', 'Role', 'KPI', 'Weight', 'Basis', 'Achievement', 'Rating', 'Weighted']);
  const d0 = aoa.length;
  for (const { l, R } of res) {
    for (const x of R.rows) aoa.push([l.name, l.level === 'RM' ? 'RM' : 'CM', x.kpi, x.w, x.basis, x.ach, x.r ?? 'Pending', x.r == null ? '' : +(x.w * x.r).toFixed(2)]);
    aoa.push([l.name, '', R.full ? 'Final score' : 'Provisional score (re-weighted over KPIs with data)', '', '', '', finalOf(R.score) ?? '', R.score == null ? '' : +R.score.toFixed(2)], []);
  }
  aoa.push(['How these are rated (Retail Goal FY 2025-26 KPI sheets)'],
    ['Sales and (RM) conversion: achievement % 70 or below = 1, 70.1–90 = 2, 90.1–100 = 3, above 100 = 4.'],
    ['CM conversion, ops scorecard, shrinkage: % of the cluster’s stores at rating 3 or better: below 50% = 1, 50–64% = 2, 65–80% = 3, above 80% = 4.'],
    ['A store reaches rating 3 on: conversion above 90% of target; shrinkage rated each quarter (0.075% or less = 3) and the quarter ratings averaged to 3 or more. Commercial Street is left out of shrinkage.'],
    ['Ops scorecard is counted month by month, not averaged: each month, stores scoring 90% or above (Green zone) out of the stores scored; the rating uses total achieved ÷ total store-months (Apr–Aug).'],
    ['Team development: attrition % above 40 = 0, 30–40 = 1, 25–30 = 2, 20–25 = 3, below 20 = 4. Promotion % is not available, so attrition alone is used.'],
    ['RM Process Improvement (15%) and Salary to Sales Ratio (10%) need manual input; until then the RM score is provisional: weighted points ÷ weights with data.'],
    ['RM Selling Effectiveness uses conversion only (no UPT targets). Final rating = score rounded to the nearest whole number (x.5 rounds up).']);
  const ws = X.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [{ wch: 22 }, { wch: 32 }, { wch: 44 }, { wch: 14 }, { wch: 52 }, { wch: 16 }, { wch: 9 }, { wch: 10 }];
  for (let r = d0; r < aoa.length; r++) { const a = X.utils.encode_cell({ r, c: 3 }); if (ws[a] && typeof ws[a].v === 'number') ws[a].z = '0%'; }
  X.utils.book_append_sheet(wb, ws, 'Ratings');
  for (const { l, R } of res) console.log(`${label(l).padEnd(34)} score ${R.score == null ? '–' : R.score.toFixed(2)} → rating ${finalOf(R.score) ?? '–'} ${R.full ? '' : '(provisional)'} | ` + R.rows.map(x => `${x.kpi.split(' ')[0]} ${x.r ?? 'P'}`).join(', '));
}

// 1. Summary: one row per RM / cluster, all four KPI groups
{
  const aoa = [['RM and cluster performance: target vs achievement, conversion, attrition, operational scorecard'], [stamp], [],
    ['', '', 'SALES TARGET VS ACHIEVEMENT', '', '', 'CONVERSION', '', '', '', '', 'ATTRITION (Apr–Sep)', '', '', '', '', 'OPS SCORECARD: STORES ACHIEVED (90%+) OF STORES SCORED', '', '', '', '', '', 'SHRINKAGE'],
    ['RM / Cluster', 'Stores', 'Target', 'Achievement', 'Ach %', 'Walk-ins', 'Bills', 'Conversion', 'Conv. target', 'Conv. ach %', 'Opening HC', 'Joiners', 'Exits', 'Closing HC', 'Attrition %',
      ...OPS_MONTHS.map(ML), 'Achieved %', 'Shrinkage %']];
  for (const l of levels) {
    const h = roll(l.g, M), a = attr(l.g), sh = shrink(l.g);
    aoa.push([label(l), l.g.length, h.target, h.sales, r4(h.ach), h.wi, h.bills, r4(h.conv), r4(h.convT), r4(h.convAch), a.open, a.join, a.exits, a.close, r4(a.pct),
      ...OPS_MONTHS.map(m => { const o = opsMonth(l.g, m); return o.n ? `${o.ok} of ${o.n}` : ''; }), r4(opsAll(l.g) ? opsAll(l.g).pct / 100 : null), sh == null ? '' : +sh.toFixed(3)]);
  }
  aoa.push([], ['Notes'], ['Sales: store Item Net Amt excluding e-commerce ÷ store target, months with a target only.'],
    ['Conversion: bills ÷ walk-ins (daily conversion report); target = walk-in-target weighted CV%.'],
    ['Attrition %: exits ÷ average of opening and closing headcount (same as the Retail Attrition report).'],
    ['Operational scorecard: each month, the number of stores scoring 90% or above (Green zone) out of the stores scored; Achieved % = total achieved / total store-months (file has Apr–Aug; no September).'],
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
      ...M.slice(0, 5).map(m => 'Ops ' + ML(m)), 'Ops H1 avg', 'Shrinkage %', ...M.map(m => 'Sales ach ' + ML(m)), 'Conv. rating', 'Ops rating', 'Shrink rating']];
  for (const l of levels.filter(x => x.level === 'Cluster')) for (const x of l.g.slice().sort((a, b) => a.k.localeCompare(b.k))) {
    const h = roll([x], M), a = attr([x]), sh = shrink([x]), i = x.s.inp || {};
    aoa.push([l.rm, l.name, x.k, x.s.code, h.target || '', h.sales || '', r4(h.ach), h.wi || '', h.bills || '', r4(h.conv), r4(h.convT), r4(h.convAch),
      i.left != null ? a.open : '', i.left != null ? a.join : '', i.left != null ? a.exits : '', i.left != null ? a.close : '', i.left != null ? r4(a.pct) : '',
      ...M.slice(0, 5).map(m => r4(roll([x], [m]).ops)), r4(h.ops), sh == null ? '' : +sh.toFixed(3), ...M.map(m => r4(roll([x], [m]).ach)), storeRt(x).conv ?? '', storeRt(x).ops ?? '', SHRINK_EXEMPT.includes(x.k) ? 'exempt' : storeRt(x).shrink ?? '']);
  }
  const ws = X.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [{ wch: 18 }, { wch: 20 }, { wch: 34 }, { wch: 9 }, { wch: 14 }, { wch: 14 }, ...Array(24).fill({ wch: 11 })];
  ws['!autofilter'] = { ref: `A4:AG${aoa.length}` };
  fmt(ws, aoa, 4, [6, 9, 10, 11, 16, 17, 18, 19, 20, 21, 22, 24, 25, 26, 27, 28, 29], [4, 5, 7, 8]);
  X.utils.book_append_sheet(wb, ws, 'Store wise');
}
fs.mkdirSync(OUT, { recursive: true });
const file = path.join(OUT, `PMS H1 - RM & Cluster performance - ${rms.join(', ').replace(/[\\/:*?"<>|]/g, '')}.xlsx`);
X.writeFile(wb, file);
console.log('written', file);
for (const l of levels) { const h = roll(l.g, M), a = attr(l.g); console.log(`${label(l).padEnd(34)} stores ${String(l.g.length).padStart(2)} | sales ${(h.sales / 1e7).toFixed(2)} of ${(h.target / 1e7).toFixed(2)} Cr = ${(h.ach * 100).toFixed(1)}% | conv ${(h.conv * 100).toFixed(1)}% vs ${(h.convT * 100).toFixed(1)}% = ${(h.convAch * 100).toFixed(0)}% | attrition ${a.exits} exits, ${(a.pct * 100).toFixed(1)}% | ops ${(h.ops * 100).toFixed(1)}%`); }
