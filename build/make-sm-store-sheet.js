// Writes "PMS H1 Inputs - SM Store Mapping.xlsx": one row per store, one column per month (Jan–Sep 2026),
// each cell = the Store Manager who ran the store that month (dropdown "EmpID - Name", or type the Emp ID).
// Sheet 2 = list of all Store Managers with Emp ID. Apr–Sep are pre-filled with what the dashboard uses today.
// usage: node build/build.js && node build/make-sm-store-sheet.js [--force]
const fs = require('fs'), X = require('xlsx'), JSZip = require('jszip');
const OUT = 'C:/Users/Lenovo/OneDrive/Desktop/Category Q1/26-27 iteamwise report H1/PMS H1 Inputs - SM Store Mapping.xlsx';
if (fs.existsSync(OUT) && !process.argv.includes('--force')) { console.log('exists, not overwritten (use --force):', OUT); process.exit(0); }
const d = require('./out/data.json');
const MONTHS = ['2026-01', '2026-02', '2026-03', '2026-04', '2026-05', '2026-06', '2026-07', '2026-08', '2026-09'];
const LAB = m => new Date(m + '-01').toLocaleString('en-GB', { month: 'short' }).slice(0, 3) + '-26';
const tc = s => String(s || '').toLowerCase().replace(/\b\w/g, c => c.toUpperCase());

const sms = d.people.filter(p => p.role === 'SM').sort((a, b) => a.name.localeCompare(b.name));
const label = p => `${p.id} - ${p.name}`;
const storeKeys = Object.entries(d.stores).filter(([k, s]) => (s.code || d.people.some(p => p.outlet === k)) && k !== 'BLR - HO').map(([k]) => k).sort();

// who runs store k in month m today: SMs whose month store(s) include k; most swipe days there wins
const runner = (k, m) => {
  if (!d.months.includes(m)) return '';
  const c = sms.filter(p => (p.monthStores ? p.monthStores[m] || [] : [p.outlet]).includes(k));
  c.sort((a, b) => (((b.att || {})[m] || {})[k] || 0) - (((a.att || {})[m] || {})[k] || 0) || b.eligible - a.eligible);
  return c.length ? label(c[0]) : '';
};
const rows = storeKeys.map(k => {
  const s = d.stores[k];
  return { 'Store Code': s.code || '', 'Store Name': k, 'State': (sms.find(p => p.outlet === k) || d.people.find(p => p.outlet === k) || {}).state || s.state || '',
    'Cluster Manager': s.cmA || s.cm || '', ...Object.fromEntries(MONTHS.map(m => [LAB(m), runner(k, m)])) };
});
const ws = X.utils.json_to_sheet(rows);
ws['!cols'] = [{ wch: 10 }, { wch: 34 }, { wch: 14 }, { wch: 18 }, ...MONTHS.map(() => ({ wch: 28 }))];
ws['!autofilter'] = { ref: `A1:M${rows.length + 1}` };

const list = sms.map(p => ({ 'Emp ID': p.id, 'SM Name': p.name, 'Pick value': label(p), 'Current Branch': p.outlet, 'Date of Joining': p.doj,
  'State': p.state, 'Cluster Manager': p.cm, 'Regional Manager': p.am, 'Eligible for H1 review': p.eligible ? 'Yes' : 'No' }));
const ws2 = X.utils.json_to_sheet(list);
ws2['!cols'] = [{ wch: 8 }, { wch: 28 }, { wch: 34 }, { wch: 32 }, { wch: 14 }, { wch: 14 }, { wch: 18 }, { wch: 16 }, { wch: 10 }];
ws2['!autofilter'] = { ref: `A1:I${list.length + 1}` };

const notes = X.utils.aoa_to_sheet([
  ['How to fill'],
  ['Sheet "Store x Month": one row per store. Under each month put the Store Manager who ran that store that month.'],
  ['Click a cell and pick "EmpID - Name" from the dropdown, or simply type the Emp ID (e.g. 1234).'],
  ['Apr-26 to Sep-26 are pre-filled with what the H1 dashboard uses today (attendance swipes > 15 days / current branch). Correct anything that is wrong.'],
  ['Jan-26 to Mar-26 are blank: fill them if needed (not used in the H1 rating).'],
  ['One SM running two stores in the same month: put their Emp ID under both stores. They are then rated on both stores combined.'],
  ['Leave a cell blank if the store had no SM that month. An SM who appears under no store in a month is not rated for that month.'],
  ['Sheet "SM List" has every Store Manager with Emp ID, current branch and joining date.'],
  ['Sheet "Please check" lists stores where two SMs were found in the same month (e.g. two SMs on the roster for one store) – please confirm who ran it.'],
  ['A month column is used for the rating once anything in it is changed; untouched months keep the dashboard’s automatic store (attendance / roster).'],
]);
notes['!cols'] = [{ wch: 130 }];

// cases HR should settle: two SMs on the same store in a month, and stores with no SM pre-filled
const check = [];
for (const k of storeKeys) {
  const groups = {};
  for (const m of d.months) {
    const c = sms.filter(p => (p.monthStores ? p.monthStores[m] || [] : [p.outlet]).includes(k));
    if (c.length > 1) { const g = c.map(label).join(' / '); (groups[g] = groups[g] || []).push(LAB(m)); }
  }
  for (const [g, ms] of Object.entries(groups)) check.push({ 'Store Name': k, 'Months': ms.join(', '), 'Issue': 'Two SMs found for this store', 'SMs found': g, 'Pre-filled with': runner(k, d.months.find(m => LAB(m) === ms[0])) });
  const none = d.months.filter(m => !runner(k, m) && (d.stores[k].m[m] || {}).sales);
  if (none.length) check.push({ 'Store Name': k, 'Months': none.map(LAB).join(', '), 'Issue': 'Store traded but no SM found', 'SMs found': '', 'Pre-filled with': '' });
}
const ws3 = X.utils.json_to_sheet(check.length ? check : [{ 'Store Name': 'Nothing to check' }]);
ws3['!cols'] = [{ wch: 34 }, { wch: 40 }, { wch: 30 }, { wch: 70 }, { wch: 34 }];

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
  xml = xml.includes('<pageMargins') ? xml.replace('<pageMargins', dv + '<pageMargins') : xml.replace('</worksheet>', dv + '</worksheet>');
  zip.file(f, xml);
  fs.writeFileSync(OUT, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
  // baseline = what was pre-filled; the build only applies month columns HR has changed from this
  const base = {};
  for (const r of rows) for (const m of d.months) { const id = (/^(T?\d+)/.exec(r[LAB(m)] || '') || [])[1]; if (id) (base[m] = base[m] || {})[r['Store Name']] = id; }
  fs.writeFileSync(__dirname + '/out/sm-map-baseline.json', JSON.stringify(base));
  console.log('please check:', check.length, 'items');
  const filled = rows.reduce((a, r) => a + MONTHS.filter(m => r[LAB(m)]).length, 0);
  console.log('written', rows.length, 'stores x', MONTHS.length, 'months;', filled, 'cells pre-filled;', list.length, 'SMs in list');
})();
