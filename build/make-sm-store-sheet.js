// Writes "PMS H1 Inputs - SM Store Mapping.xlsx": one row per store, one column per month (Jan–Sep 2026),
// each cell = the Store Manager who ran the store that month (dropdown "EmpID - Name", or type the Emp ID).
// Sheet "SM List" = every Store Manager incl. those who left during the year (from the attrition exit list).
// Pre-fill: Apr–Sep from attendance swipes (> 15 days) / billing / roster as the dashboard uses them, plus SMs who have since left;
//           Jan–Mar estimated from joining / leaving dates (no swipes for those months) — HR should check them.
// usage: node build/build.js && node build/make-sm-store-sheet.js [--force]
const fs = require('fs'), X = require('xlsx'), JSZip = require('jszip');
const OUT = process.env.SM_MAP_OUT || 'C:/Users/Lenovo/OneDrive/Desktop/Category Q1/26-27 iteamwise report H1/PMS H1 Inputs - SM Store Mapping.xlsx';   // SM_MAP_OUT: write a test copy (baseline untouched)
if (fs.existsSync(OUT) && !process.argv.includes('--force')) { console.log('exists, not overwritten (use --force):', OUT); process.exit(0); }
const d = require('./out/data.json');
const att = fs.existsSync(__dirname + '/out/att-days.json') ? require('./out/att-days.json') : {};
const MONTHS = ['2026-01', '2026-02', '2026-03', '2026-04', '2026-05', '2026-06', '2026-07', '2026-08', '2026-09'];
const LAB = m => new Date(m + '-01').toLocaleString('en-GB', { month: 'short' }).slice(0, 3) + '-26';
const ATT_MIN_DAYS = 15;

// everyone who was a Store Manager this year: current roster + SMs in the exit list
const current = d.people.filter(p => p.role === 'SM').map(p => ({ ...p, left: '' }));
const exited = [];
for (const [k, s] of Object.entries(d.stores)) for (const e of (s.inp && s.inp.exits) || [])
  if (/^store manager$/i.test(e.desig) && !current.some(p => p.id === e.id)) exited.push({ id: e.id, name: e.name, outlet: k, doj: e.doj, left: e.left, state: '', cm: s.cmA || s.cm || '', am: s.rmA || s.rm || '', eligible: false });
const sms = [...current, ...exited].sort((a, b) => a.name.localeCompare(b.name));
const label = p => `${p.id} - ${p.name}`;
const storeKeys = Object.entries(d.stores).filter(([k, s]) => (s.code || d.people.some(p => p.outlet === k)) && k !== 'BLR - HO').map(([k]) => k).sort();

// days of month m covered by the person's employment (joining → leaving)
const covered = (p, m) => {
  const first = new Date(m + '-01T00:00:00Z'), last = new Date(Date.UTC(first.getUTCFullYear(), first.getUTCMonth() + 1, 0));
  const from = p.doj ? new Date(p.doj + 'T00:00:00Z') : first, to = p.left ? new Date(p.left + 'T00:00:00Z') : last;
  const a = from > first ? from : first, b = to < last ? to : last;
  return b < a ? 0 : Math.round((b - a) / 864e5) + 1;
};
// candidates for store k in month m, with a score (swipe days where we have swipes, else days employed that month)
const candidates = (k, m) => {
  const inH1 = d.months.includes(m), out = [];
  const storesOf = (p, mm) => p.monthStores ? p.monthStores[mm] || [] : [p.outlet];
  for (const p of sms) {
    const sw = (att[p.id] || {})[m], days = sw ? sw[k] || 0 : null;
    if (!p.left) {
      // current SM. Apr–Sep: the store(s) the dashboard rates them on (swipes > 15 days / HR sheet / billing / roster).
      // Jan–Mar: no swipes, so assume they were where they were in April (if they had joined by then).
      const ks = inH1 ? storesOf(p, m) : storesOf(p, '2026-04');
      const c = covered(p, m);
      if (ks.includes(k) && c >= 1) out.push({ p, score: days != null ? Math.max(days, 1) : c, how: inH1 ? 'dashboard' : 'April store' });
      continue;
    }
    // SM who has since left: swipes when we have them, otherwise their exit branch while employed
    if (sw) { if (days > ATT_MIN_DAYS) out.push({ p, score: days, how: 'swipes' }); continue; }
    if (p.outlet !== k) continue;
    const c = covered(p, m); if (c >= 1) out.push({ p, score: c, how: 'dates' });
  }
  return out.sort((a, b) => b.score - a.score || (a.p.left ? 1 : 0) - (b.p.left ? 1 : 0));
};
const runner = (k, m) => { const c = candidates(k, m)[0]; return c ? label(c.p) : ''; };

const rows = storeKeys.map(k => {
  const s = d.stores[k];
  return { 'Store Code': s.code || '', 'Store Name': k, 'State': (current.find(p => p.outlet === k) || d.people.find(p => p.outlet === k) || {}).state || s.state || '',
    'Cluster Manager': s.cmA || s.cm || '', ...Object.fromEntries(MONTHS.map(m => [LAB(m), runner(k, m)])) };
});
const ws = X.utils.json_to_sheet(rows);
ws['!cols'] = [{ wch: 10 }, { wch: 34 }, { wch: 14 }, { wch: 18 }, ...MONTHS.map(() => ({ wch: 28 }))];
ws['!autofilter'] = { ref: `A1:M${rows.length + 1}` };

const list = sms.map(p => ({ 'Emp ID': p.id, 'SM Name': p.name, 'Pick value': label(p), 'Status': p.left ? `Left on ${p.left}` : 'Active',
  'Branch': p.outlet, 'Date of Joining': p.doj, 'State': p.state, 'Cluster Manager': p.cm, 'Regional Manager': p.am, 'Eligible for H1 review': p.left ? 'No (left)' : p.eligible ? 'Yes' : 'No' }));
const ws2 = X.utils.json_to_sheet(list);
ws2['!cols'] = [{ wch: 8 }, { wch: 30 }, { wch: 36 }, { wch: 20 }, { wch: 32 }, { wch: 14 }, { wch: 14 }, { wch: 18 }, { wch: 16 }, { wch: 12 }];
ws2['!autofilter'] = { ref: `A1:J${list.length + 1}` };

// cases HR should settle: two SMs for one store in a month, and trading stores with no SM found
const check = [];
for (const k of storeKeys) {
  const groups = {};
  for (const m of MONTHS) {
    const c = candidates(k, m).filter(x => x.score > ATT_MIN_DAYS);
    if (c.length > 1) { const g = c.map(x => label(x.p)).join(' / '); (groups[g] = groups[g] || []).push(LAB(m)); }
  }
  for (const [g, ms] of Object.entries(groups)) check.push({ 'Store Name': k, 'Months': ms.join(', '), 'Issue': 'Two SMs found for this store', 'SMs found': g, 'Pre-filled with': runner(k, MONTHS.find(m => LAB(m) === ms[0])) });
  const none = d.months.filter(m => !runner(k, m) && (d.stores[k].m[m] || {}).sales);
  if (none.length) check.push({ 'Store Name': k, 'Months': none.map(LAB).join(', '), 'Issue': 'Store traded but no SM found', 'SMs found': '', 'Pre-filled with': '' });
}
const ws3 = X.utils.json_to_sheet(check.length ? check : [{ 'Store Name': 'Nothing to check' }]);
ws3['!cols'] = [{ wch: 34 }, { wch: 46 }, { wch: 30 }, { wch: 70 }, { wch: 34 }];

const notes = X.utils.aoa_to_sheet([
  ['How to fill'],
  ['Sheet "Store x Month": one row per store. Under each month put the Store Manager who ran that store that month.'],
  ['Click a cell and pick "EmpID - Name" from the dropdown, or simply type the Emp ID (e.g. 1234).'],
  ['Apr-26 to Sep-26 are pre-filled from attendance swipes (store with more than 15 swipe days), billing and the roster, including SMs who have since left.'],
  ['Jan-26 to Mar-26 are ESTIMATES from joining / leaving dates (no swipe data for those months): please check them. They are used to see which manager a Q1 stock audit belongs to.'],
  ['One SM running two stores in the same month: put their Emp ID under both stores. They are then rated on both stores combined.'],
  ['Leave a cell blank if the store had no SM that month. An SM who appears under no store in a month is not rated for that month.'],
  ['Sheet "SM List" has every Store Manager this year with Emp ID, status (Active / Left on date), branch and joining date.'],
  ['Sheet "Please check" lists stores where two SMs were found in the same month, and trading stores with no SM found.'],
  ['A month column (Apr–Sep) is used for the rating once anything in it is changed; untouched months keep the dashboard\u2019s automatic store.'],
]);
notes['!cols'] = [{ wch: 140 }];

const wb = X.utils.book_new();
X.utils.book_append_sheet(wb, ws, 'Store x Month');
X.utils.book_append_sheet(wb, ws2, 'SM List');
X.utils.book_append_sheet(wb, ws3, 'Please check');
X.utils.book_append_sheet(wb, notes, 'How to fill');
const buf = X.write(wb, { type: 'buffer', bookType: 'xlsx' });
(async () => {
  const zip = await JSZip.loadAsync(buf);
  const f = 'xl/worksheets/sheet1.xml';
  let xml = await zip.file(f).async('string');
  // dropdown from SM List column C; typing a plain Emp ID is allowed too (no error popup)
  const dv = `<dataValidations count="1"><dataValidation type="list" allowBlank="1" showErrorMessage="0" sqref="E2:M${rows.length + 1}"><formula1>'SM List'!$C$2:$C$${list.length + 1}</formula1></dataValidation></dataValidations>`;
  // Excel requires <dataValidations> before these elements (schema order); placing it later makes Excel drop the sheet data
  const after = ['<hyperlinks', '<printOptions', '<pageMargins', '<pageSetup', '<headerFooter', '<ignoredErrors', '<drawing', '<tableParts', '<extLst', '</worksheet>'];
  const at = Math.min(...after.map(t => xml.indexOf(t)).filter(i => i >= 0));
  xml = xml.slice(0, at) + dv + xml.slice(at);
  zip.file(f, xml);
  fs.writeFileSync(OUT, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
  // baseline = what was pre-filled for Apr–Sep; the build only applies month columns HR has changed from this
  const base = {};
  for (const r of rows) for (const m of d.months) { const id = (/^(T?\d+)/.exec(r[LAB(m)] || '') || [])[1]; if (id) (base[m] = base[m] || {})[r['Store Name']] = id; }
  if (!process.env.SM_MAP_OUT) fs.writeFileSync(__dirname + '/out/sm-map-baseline.json', JSON.stringify(base));
  const filled = rows.reduce((a, r) => a + MONTHS.filter(m => r[LAB(m)]).length, 0);
  console.log('written', rows.length, 'stores x', MONTHS.length, 'months;', filled, 'cells pre-filled;', list.length, 'SMs in list (', exited.length, 'left );', check.length, 'to check');
})();
