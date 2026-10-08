// Writes the HR "Monthly Store Map" sheet, pre-filled from billing / roster, with a store dropdown in every month cell.
// usage: node build/build.js && node build/make-store-map.js [--force]
// Never overwrites an existing sheet unless --force is passed (HR edits live in it).
const fs = require('fs'), path = require('path'), X = require('xlsx'), JSZip = require('jszip');
const OUT = 'C:/Users/Lenovo/OneDrive/Desktop/Category Q1/26-27 iteamwise report H1/PMS H1 Inputs - Monthly Store Map.xlsx';
const LEGACY = 'C:/Users/Lenovo/OneDrive/Desktop/Category Q1/26-27 iteamwise report H1/PMS H1 Inputs - Store Transfers.xlsx';
const MONTHS = ['2026-04', '2026-05', '2026-06', '2026-07', '2026-08', '2026-09'];
const MLAB = { '2026-04': 'Apr', '2026-05': 'May', '2026-06': 'Jun', '2026-07': 'Jul', '2026-08': 'Aug', '2026-09': 'Sep' };

if (fs.existsSync(OUT) && !process.argv.includes('--force')) { console.log('exists, not overwritten (use --force):', OUT); process.exit(0); }
const data = require('./out/data.json');
const auto = require('./out/auto-store-map.json');

// carry over rows from the earlier transfers sheet so nothing typed there is lost
const legacy = {};
if (fs.existsSync(LEGACY)) {
  const MON = { APR: '04', MAY: '05', JUN: '06', JUL: '07', AUG: '08', SEP: '09' };
  const ym = v => { const m = /([A-Za-z]{3})\w*[\s'-]*(\d{2,4})/.exec(String(v || '')); return m ? `20${m[2].slice(-2)}-${MON[m[1].toUpperCase()]}` : ''; };
  for (const r of X.utils.sheet_to_json(X.readFile(LEGACY).Sheets.Transfers, { defval: null })) {
    const id = String(r['Emp ID'] || '').trim(), st = String(r['Store Name'] || '').trim().toUpperCase(), from = ym(r['From Month']), to = ym(r['To Month']);
    MONTHS.filter(m => m >= from && m <= to).forEach(m => ((legacy[id] = legacy[id] || {})[m] = st));
  }
}

const storeList = [...new Set([...Object.entries(data.stores).filter(([, s]) => s.code).map(([k]) => k), ...data.people.map(p => p.outlet)])].filter(k => k !== 'BLR - HO').sort();
const ORDER = { SM: 0, ASM: 1, SSA: 2 };
const rows = auto
  .filter(p => p.eligible && (p.role !== 'SSA' || MONTHS.some(m => p.months[m] !== p.outlet) || legacy[p.id]))
  .sort((a, b) => ORDER[a.role] - ORDER[b.role] || a.outlet.localeCompare(b.outlet) || a.name.localeCompare(b.name))
  .map(p => {
    const months = Object.fromEntries(MONTHS.map(m => [MLAB[m], (legacy[p.id] || {})[m] || p.months[m]]));
    const moved = MONTHS.some(m => months[MLAB[m]] !== p.outlet);
    const why = legacy[p.id] ? 'Entered earlier by HR' : moved ? 'Moved per billing – please confirm' : '';
    return { 'Emp ID': p.id, 'Employee Name': p.name, 'Designation': p.desig, 'Current Branch': p.outlet, ...months, 'Check': why };
  });

const ws = X.utils.json_to_sheet(rows);
ws['!cols'] = [{ wch: 8 }, { wch: 26 }, { wch: 24 }, { wch: 32 }, ...MONTHS.map(() => ({ wch: 30 })), { wch: 34 }];
ws['!autofilter'] = { ref: `A1:K${rows.length + 1}` };
const notes = X.utils.aoa_to_sheet([
  ['How to use this sheet'],
  ['Every eligible Store Manager and ASM is listed, plus any SA/SSA whose billing shows a different store in some month.'],
  ['Each month (Apr–Sep) already shows the store the person is rated on: from their own billing where they bill, otherwise their current branch.'],
  ['Only change a month if it is wrong: click the cell and pick the correct store from the dropdown. No typing needed.'],
  ['Rows marked "Moved per billing – please confirm" were detected automatically from the POS data. Leave them if correct.'],
  ['Anyone not in this sheet is rated on their billing store (SA/SSA) or current branch for every month.'],
  ['Save the file and tell Claude to rebuild the dashboard.'],
]);
notes['!cols'] = [{ wch: 120 }];
const st = X.utils.aoa_to_sheet([['Store'], ...storeList.map(s => [s])]);
st['!cols'] = [{ wch: 36 }];
const wb = X.utils.book_new();
X.utils.book_append_sheet(wb, ws, 'Store Map');
X.utils.book_append_sheet(wb, notes, 'How to use');
X.utils.book_append_sheet(wb, st, 'Stores');
const buf = X.write(wb, { type: 'buffer', bookType: 'xlsx' });

// SheetJS can't write data validation, so add the dropdown XML to the Store Map sheet directly
(async () => {
  const zip = await JSZip.loadAsync(buf);
  const f = 'xl/worksheets/sheet1.xml';
  let xml = await zip.file(f).async('string');
  const dv = `<dataValidations count="1"><dataValidation type="list" allowBlank="1" showErrorMessage="1" errorTitle="Pick a store" error="Choose a store from the list." sqref="E2:J${rows.length + 1}"><formula1>Stores!$A$2:$A$${storeList.length + 1}</formula1></dataValidation></dataValidations>`;
  xml = xml.includes('<pageMargins') ? xml.replace('<pageMargins', dv + '<pageMargins') : xml.replace('</worksheet>', dv + '</worksheet>');
  zip.file(f, xml);
  fs.writeFileSync(OUT, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
  const c = rows.reduce((a, r) => (a[r.Check || 'no change'] = (a[r.Check || 'no change'] || 0) + 1, a), {});
  console.log('written', rows.length, 'rows,', storeList.length, 'stores in dropdown', c);
})();
