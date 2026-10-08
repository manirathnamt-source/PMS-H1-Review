// Builds the H1 PMS dataset: sales aggregates + targets + conversion + calibration roster.
// usage: node build/build.js            -> writes build/out/data.json and data.enc (encrypted with build/passcode.txt)
const fs = require('fs'), path = require('path'), crypto = require('crypto'), X = require('xlsx');
const DL = 'C:/Users/Lenovo/Downloads/';
const SRC = {
  roster: DL + 'Eligibility List _ Pan India_ 26-27.xlsx',
  attrition: DL + 'Retail Attrition Data _ Pan India.xlsx',
  targets: DL + 'Targets.xlsx',
  conv: DL + 'conversion-data (1).csv',
};
const ELIG_DOJ_CUTOFF = '2026-07-31';
const MONTHS = ['2026-04', '2026-05', '2026-06', '2026-07', '2026-08', '2026-09'];
const MLAB = { '2026-04': 'Apr', '2026-05': 'May', '2026-06': 'Jun', '2026-07': 'Jul', '2026-08': 'Aug', '2026-09': 'Sep' };
const here = p => path.join(__dirname, p);
// one key for a store across all sources: upper-case, drop "(CODE)", collapse spaces, normalise the first dash to " - "
// HR-roster spellings that differ from the sales/target/conversion files
const ALIAS = {
  'BLR - AECS LAYOUT': 'BLR - BROOKEFIELD', 'BLR - KR PURAM': 'BLR - K R PURAM', 'INFANTRY ROAD - BELLARI': 'BLY - INFANTRY ROAD',
  'DLF MIDTOWN - DELHI': 'DEL - DLF MIDTOWN MALL', 'M3M, SECTOR 65 - GURUGRAM': 'GUR - M3M ROUTE 65', 'ERODE - PERUNDURAI ROAD': 'ERD - PERUNDURAI ROAD',
  'BAILEY SQUARE - PATNA': 'PTN - BAILEY SQUARE MALL', 'MUM - NEXUS SEAWOOD MALL': 'MUM - NEXUS SEAWOODS MALL', 'MODEL TOWN - JALANDHAR': 'JLR - ANIKI TOWER MODEL TOWN',
};
const key = s => {
  const k = String(s || '').replace(/\([^)]*\)\s*$/, '').replace(/\s+/g, ' ').trim().toUpperCase().replace(/\s*[-–]\s*/, ' - ');
  return ALIAS[k] || k;
};
const num = v => typeof v === 'number' ? v : +String(v || '').replace(/[₹,\s]/g, '') || 0;
const xdate = v => typeof v === 'number' ? new Date(Math.round((v - 25569) * 864e5)).toISOString().slice(0, 10)
  : v ? new Date(v + ' UTC').toISOString().slice(0, 10) : '';

const stores = {};
const S = k => stores[k] || (stores[k] = { name: k, m: {} });
const SM = (k, m) => S(k).m[m] || (S(k).m[m] = {});

// 1. sales (item-wise aggregates)
const empSales = {}, empMonth = {};   // empMonth[id][month][store] = bills
for (const m of MONTHS) {
  const a = JSON.parse(fs.readFileSync(here(`agg/${m}.json`)));
  for (const [o, v] of Object.entries(a.stores)) Object.assign(SM(key(o), m), { sales: v.all.net, netI: v.inst.net, qtyI: v.inst.qty, billsI: v.inst.bills });
  for (const [id, byO] of Object.entries(a.emps)) for (const [o, v] of Object.entries(byO)) {
    const e = (empSales[id] = empSales[id] || {}), r = (e[key(o)] = e[key(o)] || { net: 0, qty: 0, bills: 0 });
    r.net += v.net; r.qty += v.qty; r.bills += v.bills;
    const em = ((empMonth[id] = empMonth[id] || {})[m] = empMonth[id][m] || {}); em[key(o)] = (em[key(o)] || 0) + v.bills;
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

// 3d. attrition: store-wise exits Apr–Sep (SM KPI uses the exit count) + exit list for the drill-down
{
  const wb = X.readFile(SRC.attrition);
  for (const r of X.utils.sheet_to_json(wb.Sheets['Attrition Analysis Report'], { defval: null })) {
    const k = key(r['Store Wise']); if (!k || /GRAND TOTAL/.test(k)) continue;
    const s = S(k), i = (s.inp = s.inp || {});
    Object.assign(i, { left: num(r['Exit Counts']), open: num(r['Opening Balance']), joiners: num(r['New Joiners']), close: num(r['Closing Balance']), attrPct: num(r['Attrition %']) * 100 });
    if (r.CM) s.cmA = String(r.CM).trim();
    if (r.RM) s.rmA = String(r.RM).trim();
  }
  for (const r of X.utils.sheet_to_json(wb.Sheets['Exit List'], { defval: null })) {
    const s = S(key(r['Curr.Branch'])), i = (s.inp = s.inp || {});
    (i.exits = i.exits || []).push({ id: String(r['E Code']), name: r['Employee Name'], desig: r['Curr.Designation'], doj: xdate(r['Date Of Joining']), left: xdate(r['Leaving Date']) });
  }
}

// 4. roster: pan-India eligibility list (HR). Eligible = its "Eligible for review" column (= joined on/before 31 Jul 2026)
const ROLE = d => /^store manager$/i.test(d) ? 'SM' : /^assistant store manager$/i.test(d) ? 'ASM' : /style associate/i.test(d) ? 'SSA' : null;
const people = [], skipped = {};
{
  const rows = X.utils.sheet_to_json(X.readFile(SRC.roster).Sheets.Sheet1, { defval: null });
  for (const r of rows) {
    const id = String(r['Employee Number'] || '').trim(); if (!id) continue;
    const desig = String(r['Curr.Designation'] || '').trim(), role = ROLE(desig);
    if (!role) { skipped[desig] = (skipped[desig] || 0) + 1; continue; }
    const outlet = key(r['Curr.Branch']), s = S(outlet), doj = xdate(r['Date Of Joining']);
    const eligible = /^eligible$/i.test(String(r['Eligible for  review'] || '').trim());
    if (eligible !== (doj <= ELIG_DOJ_CUTOFF)) console.warn('eligibility differs from DOJ rule:', id, r['Employee Name'], doj);
    people.push({
      id, role, name: String(r['Employee Name'] || '').trim(), desig, grade: r['Curr.Band'] || r['Curr.Grade'], outlet, doj,
      state: r['Curr.Location'], eligible, cm: s.cmA || s.cm || '', am: s.rmA || s.rm || '',
      sales: role === 'SSA' ? (empSales[id] || {}) : undefined,
    });
  }
}
console.log('roster roles skipped (no KPI template here):', skipped);

// 4b. store for each person, month by month. Priority:
//   HR edit in the store-map sheet > attendance (SM/ASM: > ATT_MIN_DAYS swipe days at one store) > billing (store with most bills, ≥ MIN_MONTH_BILLS) > roster branch.
// Anyone whose months don't all match their roster branch gets monthStores and is rated month by month on the store(s) they ran.
const MIN_MONTH_BILLS = 10, ATT_MIN_DAYS = 15, NR = 'NOT RATED';   // NR: month left out of the rating
SRC.storeMap = 'C:/Users/Lenovo/OneDrive/Desktop/Category Q1/26-27 iteamwise report H1/PMS H1 Inputs - Monthly Store Map.xlsx';
SRC.swipes = 'C:/Users/Lenovo/OneDrive/Desktop/Category Q1/H1 attendance swipes/';
// the values the sheet was pre-filled with (written by make-store-map.js); a cell only counts as an HR edit if it differs from these
const baseline = fs.existsSync(here('out/store-map-baseline.json')) ? JSON.parse(fs.readFileSync(here('out/store-map-baseline.json'))) : null;
const hrMap = {}, hrAlso = {};
if (fs.existsSync(SRC.storeMap)) {
  for (const r of X.utils.sheet_to_json(X.readFile(SRC.storeMap).Sheets['Store Map'], { defval: null })) {
    const id = String(r['Emp ID'] || '').trim(); if (!id) continue;
    for (const m of MONTHS) {
      const v = r[MLAB[m]] ? key(r[MLAB[m]]) : ''; if (!v) continue;
      if (!baseline || !baseline[id] || baseline[id][m] !== v) (hrMap[id] = hrMap[id] || {})[m] = v;
    }
    if (r['Also Handles']) hrAlso[id] = key(r['Also Handles']);   // second store run together with the main one, whole H1
  }
}
// attendance swipes (SM/ASM): distinct swipe days per store per month
const doorMap = JSON.parse(fs.readFileSync(here('door-map.json')));
const attDays = {};   // attDays[id][month][store] = days
if (fs.existsSync(SRC.swipes)) {
  const unmapped = {};
  for (const f of fs.readdirSync(SRC.swipes).filter(f => /\.xlsx$/i.test(f) && !f.startsWith('~$'))) {
    for (const r of X.utils.sheet_to_json(X.readFile(SRC.swipes + f).Sheets.Sheet0 || X.readFile(SRC.swipes + f).Sheets[X.readFile(SRC.swipes + f).SheetNames[0]], { defval: null, raw: false })) {
      const id = String(r['Employee No'] || '').trim(), door = String(r['Door/Address'] || '').trim(), dt = new Date(String(r['Swipe Date'] || '').slice(0, 11) + ' UTC');
      if (!id || isNaN(dt)) continue;
      const m = dt.toISOString().slice(0, 7), day = dt.toISOString().slice(0, 10), st = doorMap[door];
      if (!st) { unmapped[door] = (unmapped[door] || 0) + 1; continue; }
      if (!MONTHS.includes(m)) continue;
      const o = (((attDays[id] = attDays[id] || {})[m] = attDays[id][m] || {})[st] = attDays[id][m][st] || new Set());
      o.add(day);
    }
  }
  for (const id in attDays) for (const m in attDays[id]) for (const st in attDays[id][m]) attDays[id][m][st] = attDays[id][m][st].size;
  if (Object.keys(unmapped).length) console.warn('swipe doors not in door-map.json (add them):', unmapped);
}
const isStore = k => !!(stores[k] && stores[k].code);                         // a real store with a target row
const hasH1 = k => Object.values((stores[k] || {}).m || {}).some(x => x.sales);  // new stores have no H1 sales
const autoMap = [];
for (const p of people) {
  const months = {}, defaults = {}, dsrc = {}, src = {}, detect = hasH1(p.outlet);   // staff of new stores stay unrated unless HR maps them
  const useAtt = (p.role === 'SM' || p.role === 'ASM') && attDays[p.id];
  if (useAtt) p.att = attDays[p.id];   // shown in the drill-down
  for (const m of MONTHS) {
    const bills = (empMonth[p.id] || {})[m] || {};
    const [top, n] = Object.entries(bills).filter(([k]) => isStore(k)).sort((a, b) => b[1] - a[1])[0] || [];
    const [aTop, aDays] = useAtt ? Object.entries(attDays[p.id][m] || {}).sort((a, b) => b[1] - a[1])[0] || [] : [];
    let def, defSrc;
    if (aDays > ATT_MIN_DAYS && (isStore(aTop) || aTop === p.outlet)) { def = aTop; defSrc = 'attendance'; }
    // SM/ASM with swipe data but no store above the threshold that month: the month is not rated
    else if (useAtt) { def = NR; defSrc = 'attendance'; }
    else if (detect && n >= MIN_MONTH_BILLS) { def = top; defSrc = 'billing'; }
    else { def = p.outlet; defSrc = 'roster'; }
    if (def === p.outlet && defSrc !== 'attendance') defSrc = 'roster';
    const hr = (hrMap[p.id] || {})[m];
    const hrOk = hr && (hr === NR || isStore(hr) || hr === p.outlet);
    if (hr && !hrOk) console.warn('store map: unknown store', p.id, m, hr);
    defaults[m] = def; dsrc[m] = defSrc;
    if (hrOk) { months[m] = hr; src[m] = hr === def ? defSrc : 'HR sheet'; }
    else { months[m] = def; src[m] = defSrc; }
  }
  autoMap.push({ id: p.id, name: p.name, desig: p.desig, role: p.role, outlet: p.outlet, eligible: p.eligible, months: defaults, src: dsrc });
  const also = hrAlso[p.id];
  if (also && !isStore(also)) console.warn('store map: unknown "Also Handles" store', p.id, also);
  // stores per month: main store, plus the clubbed store when one person runs both
  const per = Object.fromEntries(MONTHS.map(m => [m, months[m] === NR ? [] : also && isStore(also) && also !== months[m] ? [months[m], also] : [months[m]]]));
  if (MONTHS.some(m => per[m].length !== 1 || per[m][0] !== p.outlet)) {
    p.monthStores = per;
    p.moveSrc = [...new Set([...MONTHS.filter(m => months[m] !== p.outlet).map(m => src[m]), ...(also && isStore(also) ? ['HR sheet'] : [])])].join(' + ');
  }
}
fs.mkdirSync(here('out'), { recursive: true });
fs.writeFileSync(here('out/auto-store-map.json'), JSON.stringify(autoMap));
const multi = people.filter(p => p.monthStores);
console.log('people rated on other / combined stores:', multi.length, multi.reduce((c, p) => (c[p.role] = (c[p.role] || 0) + 1, c), {}));

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
