// Goal sheets (HR "Retail Goal" KPI forms) for Regional Managers and their Cluster Managers, with back-up data.
// Step 1 (this script): works out every KPI, writes one back-up workbook per person (one sheet per KPI, numbered like the form)
//                       and build/out/goal-sheets.json describing what to put in each form.
// Step 2 (fill-goal-sheets.ps1): copies the HR form, fills it and adds the back-up sheets behind it using Excel.
// usage: node build/export-goal-sheets.js "<RM name>" ["<RM name>" ...]
const fs = require('fs'), path = require('path'), X = require('xlsx');
const OUT = 'C:/Users/Lenovo/OneDrive/Desktop/Category Q1/26-27 iteamwise report H1/RM reports';
const TMP = path.join(__dirname, 'out/goal-backup');
const D = JSON.parse(fs.readFileSync(path.join(__dirname, 'out/data.json')));
const M = D.months, ML = m => new Date(m + '-01').toLocaleString('en-GB', { month: 'short' }).slice(0, 3);
const tc = s => String(s || '').toLowerCase().replace(/\b\w/g, c => c.toUpperCase()).trim();
const sum = (a, f) => a.reduce((x, y) => x + (f(y) || 0), 0);
const div = (a, b) => b ? a / b : null;
const r4 = v => v == null || !isFinite(v) ? '' : +v.toFixed(4);
const L = v => '₹' + (v / 1e5).toLocaleString('en-IN', { maximumFractionDigits: 1 }) + ' L';
const P = (v, d = 1) => v == null ? '–' : (v * 100).toFixed(d) + '%';

const stores = Object.entries(D.stores).filter(([, s]) => s.code).map(([k, s]) => ({ k, s, i: s.inp || {}, rm: tc(s.rmA || s.rm), cm: tc(s.cmA || s.cm) }));
const want = process.argv.slice(2);
const rms = [...new Set(stores.map(x => x.rm))].filter(r => r && want.some(w => r.toLowerCase().includes(w.toLowerCase()))).sort();
if (!rms.length) { console.error('Give at least one RM name'); process.exit(1); }

// rating scales from the "Rating" tab of the KPI forms
const rateAch = p => p == null ? null : p <= 70 ? 1 : p <= 90 ? 2 : p <= 100 ? 3 : 4;
const rateShrink = p => p == null ? null : p > 0.125 ? 1 : p > 0.075 ? 2 : p >= 0.055 ? 3 : 4;
const rateStores = p => p == null ? null : p < 50 ? 1 : p < 65 ? 2 : p <= 80 ? 3 : 4;
const rateAttr = p => p == null ? null : p > 40 ? 0 : p > 30 ? 1 : p > 25 ? 2 : p >= 20 ? 3 : 4;
const SHRINK_EXEMPT = ['BLR - COMMERCIAL STREET'];

// sales and conversion for a group of stores over a set of months (e-commerce excluded; months without a target skipped)
function roll(g, months) {
  const cells = g.flatMap(x => months.map(m => x.s.m[m] || {}));
  const sm = cells.filter(c => c.target > 0), cv = cells.filter(c => c.wiT > 0 && c.cvT > 0 && c.wi > 0);
  const sales = sum(sm, c => c.netI), target = sum(sm, c => c.target);
  const wi = sum(cv, c => c.wi), bills = sum(cv, c => c.cBills), conv = div(bills, wi), convT = div(sum(cv, c => c.wiT * c.cvT), sum(cv, c => c.wiT));
  return { sales, target, ach: div(sales, target), wi, bills, conv, convT, convAch: conv != null && convT ? conv / convT : null };
}
const attr = g => { const a = g.map(x => x.i).filter(i => i.left != null); const open = sum(a, i => i.open), join = sum(a, i => i.joiners), exits = sum(a, i => i.left), close = sum(a, i => i.close); return { open, join, exits, close, pct: div(exits, (open + close) / 2) }; };
const OPS_M = M.filter(m => stores.some(x => (x.i.scores || []).some(z => z.m === m)));
const score = (x, m) => ((x.i.scores || []).find(z => z.m === m) || {}).v;
// shrinkage: each quarter's audits rated on their own, then the quarter ratings averaged
const QS = ['Q1', 'Q2'];
function shrinkOf(x) {
  const q = QS.map(n => { const a = (x.i.audits || []).filter(z => z.q === n); if (!a.length) return null; const sale = sum(a, z => z.sale), value = sum(a, z => z.value), pct = sale ? Math.max(0, -value) / sale * 100 : null; return { sale, value, pct, rt: rateShrink(pct) }; });
  const r = q.filter(z => z && z.rt != null).map(z => z.rt);
  return { q, rt: r.length ? +(r.reduce((a, b) => a + b, 0) / r.length).toFixed(2) : null };
}

// ---- back-up sheets: every builder returns { name, aoa, tables:[{head,last,totals}], pct:[cols], num:[cols], widths } ----
const sheet = (name, title, note) => ({ name, aoa: [[title], [note], []], tables: [], fmt: [] });
const table = (sh, head, rows, totalRows = []) => {
  const h = sh.aoa.length; sh.aoa.push(head); rows.forEach(r => sh.aoa.push(r));
  sh.tables.push({ head: h + 1, last: sh.aoa.length, cols: head.length, totals: totalRows.map(i => h + 2 + i) });
  sh.aoa.push([]);
};
const groupsOf = l => l.level === 'RM' ? [...new Set(l.g.map(x => x.cm))].sort().map(cm => ({ cm, g: l.g.filter(x => x.cm === cm) })) : [{ cm: l.name, g: l.g }];
const sorted = g => [...g].sort((a, b) => a.k.localeCompare(b.k));

function salesSheet(l, no) {
  const sh = sheet(`${no} Sales`, `${no}. Sales Target vs Achievement: ${l.title}`, 'Store sales exclude e-commerce. Months without a target are left out. Amounts in ₹.');
  const head = ['Store', 'Cluster Manager', ...M.flatMap(m => [ML(m) + ' target', ML(m) + ' sales', ML(m) + ' ach %']), 'H1 target', 'H1 sales', 'H1 ach %'];
  const line = (lab, cm, g) => { const h = roll(g, M); return [lab, cm, ...M.flatMap(m => { const r = roll(g, [m]); return [r.target || '', r.target ? r.sales : '', r4(r.ach)]; }), h.target || '', h.target ? h.sales : '', r4(h.ach)]; };
  const rows = [], tot = [];
  for (const { cm, g } of groupsOf(l)) {
    sorted(g).forEach(x => rows.push(line(tc(x.k), cm, [x])));
    if (l.level === 'RM') { tot.push(rows.length); rows.push(line('Cluster total: ' + cm, cm, g)); }
  }
  tot.push(rows.length); rows.push(line(l.level === 'RM' ? 'REGION TOTAL' : 'CLUSTER TOTAL', '', l.g));
  table(sh, head, rows, tot);
  const h = roll(l.g, M);
  sh.aoa.push([`Rating: ${L(h.sales)} of ${L(h.target)} = ${P(h.ach)} → rating ${rateAch(h.ach * 100)}`], ['Scale: 70% or below = 1, 70.1–90% = 2, 90.1–100% = 3, above 100% = 4']);
  sh.pct = M.map((_, i) => 4 + i * 3).concat(4 + M.length * 3 - 0 - 0).map(c => c); sh.pct[sh.pct.length - 1] = 2 + M.length * 3 + 2;
  sh.num = M.flatMap((_, i) => [2 + i * 3, 3 + i * 3]).concat(2 + M.length * 3, 3 + M.length * 3);
  return sh;
}
function convSheetCM(l, no) {
  const sh = sheet(`${no} Conversion`, `${no}. Selling Effectiveness (Conversion): ${l.title}`, "Counted month by month, not on the half-year total: a store achieves the month when its conversion is above 90% of that month's target. Conversion = bills ÷ walk-ins.");
  const rows = sorted(l.g).map(x => { const v = M.map(m => roll([x], [m])), have = v.filter(z => z.convAch != null); return [tc(x.k), ...v.flatMap(z => z.convAch == null ? ['', '', ''] : [r4(z.conv), r4(z.convT), r4(z.convAch)]), have.filter(z => z.convAch > 0.9).length, have.length]; });
  const c = l.R.conv;
  rows.push(['Stores achieved (above 90%)', ...c.byMonth.flatMap(z => ['', '', z.ok]), c.ok, '']);
  rows.push(['Stores with a target', ...c.byMonth.flatMap(z => ['', '', z.n]), '', c.n]);
  table(sh, ['Store', ...M.flatMap(m => [ML(m) + ' conv.', ML(m) + ' target', ML(m) + ' ach %']), 'Months achieved', 'Months counted'], rows, [rows.length - 2, rows.length - 1]);
  sh.aoa.push([`Rating: ${c.ok} of ${c.n} store-months achieved = ${c.pct.toFixed(0)}% → rating ${rateStores(c.pct)}`], ["Scale (share achieved): below 50% = 1, 50–64% = 2, 65–80% = 3, above 80% = 4. Achieved = conversion above 90% of the month's target (store rating 3 or better)."]);
  sh.pctRows = { cols: M.flatMap((_, i) => [1 + i * 3, 2 + i * 3, 3 + i * 3]), upto: rows.length - 1, f: '0.0%' }; sh.pct = []; sh.num = [];
  sh.mark = { cols: M.map((_, i) => 4 + i * 3), rows: rows.length - 2 };   // ach % cells shaded green when above 90%
  return sh;
}
function convSheet(l, no) {
  const cmRule = l.level !== 'RM';
  const sh = sheet(`${no} Conversion`, `${no}. Selling Effectiveness (Conversion): ${l.title}`, 'Conversion = bills ÷ walk-ins. Target = the store target conversion, weighted by target walk-ins. Achievement = actual ÷ target.');
  const head = ['Store', 'Cluster Manager', 'Walk-ins', 'Bills', 'Conversion', 'Target conversion', 'Achievement %', 'Store rating', ...(cmRule ? ['Achieved (rating 3+)'] : []), ...M.map(m => ML(m) + ' conv.')];
  const line = (lab, cm, g, one) => { const h = roll(g, M), rt = rateAch(h.convAch == null ? null : h.convAch * 100); return [lab, cm, h.wi || '', h.bills || '', r4(h.conv), r4(h.convT), r4(h.convAch), one ? rt ?? '' : '', ...(cmRule ? [one ? (rt == null ? 'No data' : rt >= 3 ? 'Yes' : 'No') : ''] : []), ...M.map(m => r4(roll(g, [m]).conv))]; };
  const rows = [], tot = [];
  for (const { cm, g } of groupsOf(l)) {
    sorted(g).forEach(x => rows.push(line(tc(x.k), cm, [x], true)));
    if (l.level === 'RM') { tot.push(rows.length); rows.push(line('Cluster total: ' + cm, cm, g)); }
  }
  tot.push(rows.length); rows.push(line(l.level === 'RM' ? 'REGION TOTAL' : 'CLUSTER TOTAL', '', l.g));
  table(sh, head, rows, tot);
  const h = roll(l.g, M);
  if (cmRule) { const c = l.R.convShare; sh.aoa.push([`Rating: ${c.ok} of ${c.n} stores at rating 3 or better = ${c.pct.toFixed(0)}% of stores → rating ${rateStores(c.pct)}`], ['Scale (share of stores at rating 3+): below 50% = 1, 50–64% = 2, 65–80% = 3, above 80% = 4. A store is at rating 3 when its achievement is above 90%.']); }
  else sh.aoa.push([`Rating: conversion ${P(h.conv)} vs target ${P(h.convT)} = ${P(h.convAch)} achievement → rating ${rateAch(h.convAch * 100)}`], ['Scale: 70% or below = 1, 70.1–90% = 2, 90.1–100% = 3, above 100% = 4. UPT targets are not available, so conversion alone is used.']);
  const first = cmRule ? 9 : 8;
  sh.pct = [4, 5, 6, ...M.map((_, i) => first + i)]; sh.num = [2, 3];
  return sh;
}
function opsSheet(l, no) {
  const sh = sheet(`${no} Ops scorecard`, `${no}. Retail Operational Scorecard: ${l.title}`, 'Counted month by month, not averaged: a store achieves the month when it scores 90% or above (Green zone).');
  const rows = sorted(l.g).map(x => { const v = OPS_M.map(m => score(x, m)), have = v.filter(z => z != null); return [tc(x.k), ...v.map(z => z == null ? '' : z / 100), have.filter(z => z >= 90).length, have.length]; });
  const o = l.R.ops;
  rows.push(['Stores achieved (90%+)', ...o.byMonth.map(z => z.ok), o.ok, '']);
  rows.push(['Stores scored', ...o.byMonth.map(z => z.n), '', o.n]);
  table(sh, ['Store', ...OPS_M.map(ML), 'Months achieved', 'Months scored'], rows, [rows.length - 2, rows.length - 1]);
  sh.aoa.push([`Rating: ${o.ok} of ${o.n} store-months achieved = ${o.pct.toFixed(0)}% → rating ${rateStores(o.pct)}`], ['Scale (share achieved): below 50% = 1, 50–64% = 2, 65–80% = 3, above 80% = 4']);
  sh.pctRows = { cols: OPS_M.map((_, i) => 1 + i), upto: rows.length - 2 }; sh.pct = []; sh.num = [];
  return sh;
}
function attrSheet(l, no) {
  const sh = sheet(`${no} Attrition`, `${no}. Team Development (Attrition): ${l.title}`, 'Attrition % = exits in H1 ÷ average headcount (opening + closing) ÷ 2. From the Retail Attrition report.');
  const line = (lab, cm, g) => { const a = attr(g); return [lab, cm, a.open, a.join, a.exits, a.close, r4(a.pct)]; };
  const rows = [], tot = [];
  for (const { cm, g } of groupsOf(l)) {
    sorted(g).filter(x => x.i.left != null).forEach(x => rows.push(line(tc(x.k), cm, [x])));
    if (l.level === 'RM') { tot.push(rows.length); rows.push(line('Cluster total: ' + cm, cm, g)); }
  }
  tot.push(rows.length); rows.push(line(l.level === 'RM' ? 'REGION TOTAL' : 'CLUSTER TOTAL', '', l.g));
  table(sh, ['Store', 'Cluster Manager', 'Opening headcount', 'Joiners', 'Exits', 'Closing headcount', 'Attrition %'], rows, tot);
  const a = attr(l.g);
  sh.aoa.push([`Rating: ${a.exits} exits ÷ average headcount ${((a.open + a.close) / 2).toFixed(1)} = ${P(a.pct)} → rating ${rateAttr(a.pct * 100)}`], ['Scale: above 40% = 0, 30–40% = 1, 25–30% = 2, 20–25% = 3, below 20% = 4. Promotion % is not available, so attrition alone is used.']);
  sh.pct = [6]; sh.num = [];
  return sh;
}
function shrinkSheet(l, no) {
  const sh = sheet(`${no} Shrinkage`, `${no}. Shrinkage Control: ${l.title}`, "Counted quarter by quarter: a store achieves the quarter when that quarter's shrinkage (net loss ÷ store sale of its stock audits) is 0.075% or less. A net excess counts as 0%.");
  const yes = q => q == null || q.rt == null ? '' : q.rt >= 3 ? 'Yes' : 'No';
  const rows = sorted(l.g).map(x => {
    const ex = SHRINK_EXEMPT.includes(x.k), s = shrinkOf(x), have = s.q.filter(q => q && q.rt != null);
    return [tc(x.k), ...s.q.flatMap(q => q ? [q.sale, q.value, q.pct == null ? '' : +(q.pct / 100).toFixed(6), q.rt ?? '', ex ? 'Exempt' : yes(q)] : ['', '', '', '', ex ? 'Exempt' : 'No audit']), ex ? 'Exempt' : have.filter(q => q.rt >= 3).length, ex ? 'Exempt' : have.length];
  });
  const s = l.R.shrinkShare;
  if (s) {
    rows.push(['Stores achieved (0.075% or less)', ...s.qs.flatMap(q => ['', '', '', '', q.ok]), s.ok, '']);
    rows.push(['Stores audited', ...s.qs.flatMap(q => ['', '', '', '', q.n]), '', s.n]);
  }
  table(sh, ['Store', ...QS.flatMap(q => [q + ' store sale', q + ' shrink value', q + ' shrink %', q + ' rating', q + ' achieved']), 'Quarters achieved', 'Quarters audited'], rows, s ? [rows.length - 2, rows.length - 1] : []);
  sh.aoa.push([s ? `Rating: ${s.ok} of ${s.n} store-quarters achieved = ${s.pct.toFixed(0)}% → rating ${rateStores(s.pct)}` : 'Rating: no stock audits'],
    ['Quarter scale: above 0.125% = 1, 0.076–0.125% = 2, 0.055–0.075% = 3, below 0.055% = 4. Cluster scale (share achieved): below 50% = 1, 50–64% = 2, 65–80% = 3, above 80% = 4. Commercial Street is left out.']);
  sh.pct3 = [3, 8]; sh.pct = []; sh.num = [1, 2, 6, 7];
  return sh;
}

// ---- ratings ----
function rate(l) {
  const h = roll(l.g, M), a = attr(l.g), ap = a.pct == null ? null : a.pct * 100, R = { h, a };
  const share = f => { const r = l.g.map(f).filter(v => v != null); return r.length ? { ok: r.filter(v => v >= 3).length, n: r.length, pct: r.filter(v => v >= 3).length / r.length * 100 } : null; };
  { const cb = M.map(m => { const v = l.g.map(x => roll([x], [m]).convAch).filter(z => z != null); return { ok: v.filter(z => z > 0.9).length, n: v.length }; }); R.conv = { byMonth: cb, ok: sum(cb, z => z.ok), n: sum(cb, z => z.n) }; R.conv.pct = R.conv.n ? R.conv.ok / R.conv.n * 100 : null; }
  { const qs = QS.map((n, i) => { const v = l.g.filter(x => !SHRINK_EXEMPT.includes(x.k)).map(x => shrinkOf(x).q[i]).filter(q => q && q.rt != null); return { q: n, ok: v.filter(q => q.rt >= 3).length, n: v.length }; }), ok = sum(qs, z => z.ok), n = sum(qs, z => z.n); R.shrinkShare = n ? { qs, ok, n, pct: ok / n * 100 } : null; }
  const bm = OPS_M.map(m => { const v = l.g.map(x => score(x, m)).filter(z => z != null); return { ok: v.filter(z => z >= 90).length, n: v.length }; });
  R.ops = { byMonth: bm, ok: sum(bm, z => z.ok), n: sum(bm, z => z.n) }; R.ops.pct = R.ops.n ? R.ops.ok / R.ops.n * 100 : null;
  const sales = { remark: `${l.g.length} stores. ${L(h.sales)} of ${L(h.target)} (excl. e-commerce)`, ach: P(h.ach), r: rateAch(h.ach * 100) };
  const team = { remark: `${a.exits} exits, opening ${a.open}, closing ${a.close}. Promotion % not available.`, ach: 'Attrition ' + P(a.pct), r: rateAttr(ap) };
  if (l.level === 'RM') R.rows = {
    10: sales,
    11: { remark: `Conversion ${P(h.conv)} vs target ${P(h.convT)}. UPT targets not available.`, ach: P(h.convAch), r: rateAch(h.convAch * 100) },
    12: { remark: 'To be filled: list the initiatives taken in H1', ach: '', r: null },
    13: { remark: 'To be filled: salary to sales ratio for H1', ach: '', r: null },
    14: team };
  else R.rows = {
    10: sales,
    11: { remark: M.map((m, i) => `${ML(m)} ${R.conv.byMonth[i].ok}/${R.conv.byMonth[i].n}`).join(', ') + `. ${R.conv.ok} of ${R.conv.n} store-months above 90% of target. Cluster conversion ${P(h.conv)} vs target ${P(h.convT)}.`, ach: R.conv.pct.toFixed(0) + '% achieved', r: rateStores(R.conv.pct) },
    12: { remark: OPS_M.map((m, i) => `${ML(m)} ${bm[i].ok}/${bm[i].n}`).join(', ') + `. ${R.ops.ok} of ${R.ops.n} store-months at 90% or above.`, ach: R.ops.pct.toFixed(0) + '% achieved', r: rateStores(R.ops.pct) },
    13: team,
    14: { remark: R.shrinkShare ? R.shrinkShare.qs.map(q => `${q.q} ${q.ok}/${q.n}`).join(', ') + `. ${R.shrinkShare.ok} of ${R.shrinkShare.n} store-quarters at 0.075% or less${l.g.some(x => SHRINK_EXEMPT.includes(x.k)) ? '. Commercial Street excluded.' : '.'}` : 'No stock audits', ach: R.shrinkShare ? R.shrinkShare.pct.toFixed(0) + '% achieved' : '', r: R.shrinkShare ? rateStores(R.shrinkShare.pct) : null } };
  return R;
}

fs.mkdirSync(TMP, { recursive: true });
const forms = [];
for (const rm of rms) {
  const g = stores.filter(x => x.rm === rm);
  const levels = [{ level: 'RM', name: rm, g, title: `RM ${rm}` }, ...[...new Set(g.map(x => x.cm))].sort().map(cm => ({ level: 'CM', name: cm, rm, g: g.filter(x => x.cm === cm), title: `Cluster ${cm} (RM ${rm})` }))];
  for (const l of levels) {
    l.R = rate(l);
    const sheets = l.level === 'RM' ? [salesSheet(l, 1), convSheet(l, 2), attrSheet(l, 5)] : [salesSheet(l, 1), convSheetCM(l, 2), opsSheet(l, 3), attrSheet(l, 4), shrinkSheet(l, 5)];
    const wb = X.utils.book_new();
    for (const sh of sheets) {
      const ws = X.utils.aoa_to_sheet(sh.aoa);
      for (const t of sh.tables) for (let r = t.head; r < t.last; r++) {
        const z = (c, f) => { const a = X.utils.encode_cell({ r, c }); if (ws[a] && typeof ws[a].v === 'number') ws[a].z = f; };
        sh.pct.forEach(c => z(c, '0.0%')); sh.num.forEach(c => z(c, '#,##0')); (sh.pct3 || []).forEach(c => z(c, '0.000%'));
        if (sh.pctRows && r < t.head + sh.pctRows.upto) sh.pctRows.cols.forEach(c => z(c, sh.pctRows.f || '0%'));
      }
      X.utils.book_append_sheet(wb, ws, sh.name);
    }
    const role = l.level === 'RM' ? 'Regional Manager' : 'Cluster Manager', backup = path.join(TMP, `${role} - ${l.name}.xlsx`);
    X.writeFile(wb, backup);
    forms.push({
      role, name: l.name, kpiSheet: l.level === 'RM' ? 'RM KPI' : 'CM KPI', template: `Retail Goal - ${role} - FY 2025-26.xlsx`,
      file: path.join(OUT, `Retail Goal - ${role} - ${l.name} - H1 FY 2026-27.xlsx`).replace(/\//g, '\\'), backup: backup.replace(/\//g, '\\'),
      manager: l.level === 'RM' ? 'Praveen Kumar - Retail Business Head' : `${rm} - Regional Manager`,
      rows: Object.entries(l.R.rows).map(([row, v]) => ({ row: +row, ...v })),
      sheets: sheets.map(sh => ({ name: sh.name, tables: sh.tables, rows: sh.aoa.length, mark: sh.mark ? { cols: sh.mark.cols.map(c => c + 1), first: sh.tables[0].head + 1, last: sh.tables[0].head + sh.mark.rows } : null })),
    });
    const sc = Object.values(l.R.rows);
    console.log(`${role.padEnd(17)} ${l.name.padEnd(18)} ${sc.map(v => v.r ?? '-').join(' ')}  | ${sc.map(v => v.ach || '-').join(' | ')}`);
  }
}
fs.writeFileSync(path.join(__dirname, 'out/goal-sheets.json'), JSON.stringify(forms, null, 1));
console.log(forms.length, 'forms described in build/out/goal-sheets.json');
