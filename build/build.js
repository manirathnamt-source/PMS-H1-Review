// Builds the H1 PMS dataset: sales aggregates + targets + conversion + calibration roster.
// usage: node build/build.js            -> writes build/out/data.json and data.enc (encrypted with build/passcode.txt)
const fs = require('fs'), path = require('path'), crypto = require('crypto'), X = require('xlsx');
const DL = 'C:/Users/Lenovo/Downloads/';
const SRC = {
  calib: DL + 'PMS Retail Calibration _ KA _KL _ Mide Year Review.xlsx',
  targets: DL + 'Targets.xlsx',
  conv: DL + 'conversion-data (1).csv',
};
const MONTHS = ['2026-04', '2026-05', '2026-06', '2026-07', '2026-08', '2026-09'];
const MLAB = { '2026-04': 'Apr', '2026-05': 'May', '2026-06': 'Jun', '2026-07': 'Jul', '2026-08': 'Aug', '2026-09': 'Sep' };
const here = p => path.join(__dirname, p);
// one key for a store across all sources: upper-case, drop "(CODE)", collapse spaces
// calibration-sheet spellings that differ from the sales/target/conversion files
const ALIAS = { 'BLR - KR PURAM': 'BLR - K R PURAM', 'INFANTRY ROAD - BELLARI': 'BLY - INFANTRY ROAD' };
const key = s => { const k = String(s || '').replace(/\([^)]*\)\s*$/, '').replace(/\s+/g, ' ').trim().toUpperCase(); return ALIAS[k] || k; };
const num = v => typeof v === 'number' ? v : +String(v || '').replace(/[₹,\s]/g, '') || 0;
const xdate = v => typeof v === 'number' ? new Date(Math.round((v - 25569) * 864e5)).toISOString().slice(0, 10)
  : v ? new Date(v + ' UTC').toISOString().slice(0, 10) : '';

const stores = {};
const S = k => stores[k] || (stores[k] = { name: k, m: {} });
const SM = (k, m) => S(k).m[m] || (S(k).m[m] = {});

// 1. sales (item-wise aggregates)
const empSales = {};
for (const m of MONTHS) {
  const a = JSON.parse(fs.readFileSync(here(`agg/${m}.json`)));
  for (const [o, v] of Object.entries(a.stores)) Object.assign(SM(key(o), m), { sales: v.all.net, netI: v.inst.net, qtyI: v.inst.qty, billsI: v.inst.bills });
  for (const [id, byO] of Object.entries(a.emps)) for (const [o, v] of Object.entries(byO)) {
    const e = (empSales[id] = empSales[id] || {}), r = (e[key(o)] = e[key(o)] || { net: 0, qty: 0, bills: 0 });
    r.net += v.net; r.qty += v.qty; r.bills += v.bills;
  }
}

// 2. targets: sales, walk-in and conversion targets per month
{
  const rows = X.utils.sheet_to_json(X.readFile(SRC.targets).Sheets.Sheet1, { header: 1, defval: null });
  const h = rows[0].map(x => String(x || '').replace(/\s+/g, ' ').trim().toUpperCase());
  const col = (pfx, mon) => h.findIndex(x => x === `${pfx} ${mon.toUpperCase()}` || (mon === 'Jul' && x === `${pfx} JULY`));
  for (const r of rows.slice(1)) {
    if (!r[4]) continue;
    const s = S(key(r[4]));
    Object.assign(s, { code: (/\(([^)]+)\)\s*$/.exec(r[4]) || [])[1] || '', kfj: r[0], market: r[1], rm: r[2], cm: r[3], state: r[5] });
    for (const m of MONTHS) {
      const L = MLAB[m];
      Object.assign(SM(s.name, m), { target: num(r[col('TARGET', L)]), wiT: num(r[col('WI', L)]), cvT: num(r[col('CV%', L)]) });
    }
  }
}

// 3. conversion: daily walk-ins & bills -> store-month
{
  const lines = fs.readFileSync(SRC.conv, 'utf8').split(/\r?\n/);
  const H = lines[0].split(','), iO = H.indexOf('Outlet Name'), iD = H.indexOf('Date'), iW = H.indexOf('Walk-ins'), iB = H.indexOf('Bills');
  for (const l of lines.slice(1)) {
    const c = l.split(','); if (c.length < H.length) continue;
    const [d, mo, y] = c[iD].split(' ')[0].split('/'); const m = `${y}-${mo}`;
    if (!MONTHS.includes(m)) continue;
    const r = SM(key(c[iO]), m); r.wi = (r.wi || 0) + num(c[iW]); r.cBills = (r.cBills || 0) + num(c[iB]);
    r.lastDay = Math.max(r.lastDay || 0, +d);
  }
}

// 3b. shrinkage audits + monthly operational scorecard
SRC.ops = DL + "Shrinkage and operational score card deratil April to Sep'26.xlsx";
{
  const wb = X.readFile(SRC.ops);
  const inp = k => (S(k).inp = S(k).inp || {});
  // shrinkage: one row per stock audit; H1 % = net loss value ÷ store sale over all audits (net excess counts as 0)
  const seen = new Set();
  for (const r of X.utils.sheet_to_json(wb.Sheets['Shrinkage report'], { defval: null })) {
    const k = key(r.AB); if (!k) continue;
    const dup = [k, r['Shrinkage Value'], r['Shrinkage Qty'], r['Store Sale']].join('|'); if (seen.has(dup)) continue; seen.add(dup);
    const i = inp(k); (i.audits = i.audits || []).push({ date: xdate(r.Date), month: String(r.Month || '').slice(0, 3), auditor: r.Auditor, pct: num(r['Shrinkage %']), value: num(r['Shrinkage Value']), qty: num(r['Shrinkage Qty']), sale: num(r['Store Sale']) });
  }
  for (const s of Object.values(stores)) if (s.inp && s.inp.audits) {
    const v = s.inp.audits.reduce((a, x) => a + x.value, 0), sale = s.inp.audits.reduce((a, x) => a + x.sale, 0);
    s.inp.shrink = sale ? Math.max(0, -v) / sale * 100 : null;
  }
  // scorecard: header row holds month dates (Excel serials) followed by a Zone column
  const rows = X.utils.sheet_to_json(wb.Sheets['Operational score card'], { header: 1, defval: null });
  const mcols = rows[0].map((h, i) => typeof h === 'number' ? { i, m: xdate(h).slice(0, 7) } : null).filter(Boolean);
  for (const r of rows.slice(1)) {
    const k = key(r[0]); if (!k) continue;
    const sc = mcols.map(({ i, m }) => ({ m, v: typeof r[i] === 'number' ? r[i] * (r[i] <= 1.5 ? 100 : 1) : null, zone: r[i + 1] })).filter(x => x.v != null);
    const i = inp(k); i.scores = sc; i.audit = sc.length ? sc.reduce((a, x) => a + x.v, 0) / sc.length : null;
  }
}

// 3c. manual store inputs (attrition; can also override audit/shrinkage) — optional template HR fills
SRC.inputs = 'C:/Users/Lenovo/OneDrive/Desktop/Category Q1/26-27 iteamwise report H1/PMS H1 Inputs - Audit Shrinkage Attrition.xlsx';
if (fs.existsSync(SRC.inputs)) {
  const rows = X.utils.sheet_to_json(X.readFile(SRC.inputs).Sheets.Inputs, { defval: null });
  const val = v => v == null || v === '' ? null : (typeof v === 'number' ? v : parseFloat(String(v).replace('%', '')));
  for (const r of rows) {
    const s = stores[key(r['Store Name'])]; if (!s) continue;
    const i = (s.inp = s.inp || {}), set = (f, v) => { if (v != null && !isNaN(v)) i[f] = v; };
    set('audit', val(r['Audit Score %'])); set('shrink', val(r['Shrinkage %'])); set('left', val(r['People Left (H1)']));
  }
}

// 4. calibration roster (eligibility list) — parse by header label, never fixed index
const people = [];
{
  const wb = X.readFile(SRC.calib);
  const ROLE = { 'SM sheet': 'SM', 'ASM sheet': 'ASM', 'SSA-SA sheet': 'SSA' };
  for (const sn of wb.SheetNames) {
    const role = ROLE[sn.trim()]; if (!role) continue;
    const rows = X.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: null });
    const h = rows[1].map(x => String(x || '').trim());
    const g = (r, label) => r[h.indexOf(label)];
    for (const r of rows.slice(2)) {
      const id = String(g(r, 'Emp ID') || '').trim(); if (!id) continue;
      const outlet = key(g(r, 'Outlet Name'));
      const sales = empSales[id] || {};
      people.push({
        id, role, name: String(g(r, 'Names') || '').trim(), desig: g(r, 'Designation'), grade: g(r, 'Group'),
        outlet, gender: g(r, 'Gender'), doj: xdate(g(r, 'DOJ')), conf: String(g(r, 'Confirmation') || '').trim(),
        state: g(r, 'State'), status: g(r, 'Working Status'),
        eligible: /^eligible$/i.test(String(g(r, 'Eligible for  review') || '').trim()),
        cm: g(r, 'Curr.Clustermanager') || S(outlet).cm || '', am: g(r, 'Curr.Areamanager') || S(outlet).rm || '',
        sales: role === 'SSA' ? sales : undefined,
      });
    }
  }
}

// 5. coverage report
const used = new Set(people.map(p => p.outlet));
const gaps = [...used].map(k => ({ k, s: stores[k] })).filter(({ s }) => !s || !s.code || MONTHS.some(m => !(s.m[m] || {}).sales || !(s.m[m] || {}).wi))
  .map(({ k, s }) => `${k}: ${!s ? 'not in any source' : (!s.code ? 'no target row; ' : '') + MONTHS.filter(m => !(s.m[m] || {}).sales).map(m => 'no sales ' + MLAB[m]).concat(MONTHS.filter(m => !(s.m[m] || {}).wi).map(m => 'no walk-ins ' + MLAB[m])).join(', ')}`);
console.log('people', people.length, 'stores used', used.size, '\ncoverage gaps:\n ' + gaps.join('\n '));

// 6. write + encrypt (AES-256-GCM, PBKDF2-SHA256 key) — decrypted in the browser with WebCrypto
const data = { generated: new Date().toISOString(), months: MONTHS, stores, people };
fs.mkdirSync(here('out'), { recursive: true });
fs.writeFileSync(here('out/data.json'), JSON.stringify(data));
const pass = fs.readFileSync(here('passcode.txt'), 'utf8').trim();
const salt = crypto.randomBytes(16), iv = crypto.randomBytes(12), ITER = 250000;
const k = crypto.pbkdf2Sync(pass, salt, ITER, 32, 'sha256');
const ci = crypto.createCipheriv('aes-256-gcm', k, iv);
const ct = Buffer.concat([ci.update(JSON.stringify(data), 'utf8'), ci.final(), ci.getAuthTag()]);
fs.writeFileSync(path.join(__dirname, '..', 'data.enc'), JSON.stringify({ v: 1, iter: ITER, salt: salt.toString('base64'), iv: iv.toString('base64'), ct: ct.toString('base64') }));
console.log('data.enc written', (ct.length / 1024).toFixed(0), 'KB');
