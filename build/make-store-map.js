// Writes the HR "Monthly Store Map" sheet: one row per person, a store dropdown for every month (Apr–Sep)
// and an "Also Handles" dropdown for a second store run together with the main one (e.g. one SM for two stores).
// Pre-filled from billing / roster. Anything already in the existing sheet is KEPT, so it is safe to re-run.
// usage: node build/build.js && node build/make-store-map.js [--set <empId>:also=<STORE>]...
const fs = require('fs'), X = require('xlsx'), JSZip = require('jszip');
const OUT = 'C:/Users/Lenovo/OneDrive/Desktop/Category Q1/26-27 iteamwise report H1/PMS H1 Inputs - Monthly Store Map.xlsx';
const MONTHS = ['2026-04', '2026-05', '2026-06', '2026-07', '2026-08', '2026-09'];
const MLAB = { '2026-04': 'Apr', '2026-05': 'May', '2026-06': 'Jun', '2026-07': 'Jul', '2026-08': 'Aug', '2026-09': 'Sep' };
const ALSO = 'Also Handles';
const data = require('./out/data.json');
const auto = require('./out/auto-store-map.json');
const up = v => String(v || '').trim().toUpperCase();

// existing HR edits win over the automatic values. A cell is an HR edit only if it differs from what the sheet was pre-filled with
// (the baseline written last time); untouched cells are refreshed with the latest automatic value.
const BASE = __dirname + '/out/store-map-baseline.json';
const baseline = fs.existsSync(BASE) ? JSON.parse(fs.readFileSync(BASE)) : {};
const existing = {};
if (fs.existsSync(OUT)) for (const r of X.utils.sheet_to_json(X.readFile(OUT).Sheets['Store Map'], { defval: null })) {
  const id = String(r['Emp ID'] || '').trim(); if (!id) continue;
  const edited = MONTHS.map(m => [m, up(r[MLAB[m]])]).filter(([m, v]) => v && (!baseline[id] || baseline[id][m] !== v));
  existing[id] = { months: Object.fromEntries(edited), also: up(r[ALSO]) };
}
// --set 1234:also=BLR - ARS COMPLEX MALLESHWARAM  (values given in chat, applied as if HR picked them)
process.argv.forEach((a, i, all) => {
  if (a !== '--set') return;
  const m = /^(\w+):(\w+)=(.+)$/.exec(all[i + 1] || ''); if (!m) return;
  const e = (existing[m[1]] = existing[m[1]] || { months: {} });
  if (m[2] === 'also') e.also = up(m[3]); else if (MONTHS.includes(m[2])) e.months[m[2]] = up(m[3]);
});

const storeList = ['NOT RATED', ...[...new Set([...Object.entries(data.stores).filter(([, s]) => s.code).map(([k]) => k), ...data.people.map(p => p.outlet)])].filter(k => k !== 'BLR - HO').sort()];
const ORDER = { SM: 0, ASM: 1, SSA: 2 };
const rows = auto
  .filter(p => p.eligible && (p.role !== 'SSA' || MONTHS.some(m => p.months[m] !== p.outlet) || existing[p.id]))
  .sort((a, b) => ORDER[a.role] - ORDER[b.role] || a.outlet.localeCompare(b.outlet) || a.name.localeCompare(b.name))
  .map(p => {
    const ex = existing[p.id] || { months: {} };
    const months = Object.fromEntries(MONTHS.map(m => [MLAB[m], ex.months[m] || p.months[m]]));
    const byHR = ex.also || MONTHS.some(m => ex.months[m] && ex.months[m] !== p.months[m]);
    const moved = MONTHS.some(m => months[MLAB[m]] !== p.outlet);
    const how = [...new Set(MONTHS.filter(m => p.months[m] !== p.outlet).map(m => (p.src || {})[m]).filter(Boolean))].join(' + ') || 'billing';
    const check = byHR ? 'Changed by HR' : moved ? `Moved per ${how} – please confirm` : '';
    return { 'Emp ID': p.id, 'Employee Name': p.name, 'Designation': p.desig, 'Current Branch': p.outlet, ...months, [ALSO]: ex.also || '', 'Check': check };
  });

const ws = X.utils.json_to_sheet(rows);
ws['!cols'] = [{ wch: 8 }, { wch: 26 }, { wch: 24 }, { wch: 32 }, ...MONTHS.map(() => ({ wch: 30 })), { wch: 32 }, { wch: 34 }];
ws['!autofilter'] = { ref: `A1:L${rows.length + 1}` };
const notes = X.utils.aoa_to_sheet([
  ['How to use this sheet'],
  ['Every eligible Store Manager and ASM is listed, plus any SA/SSA whose billing shows a different store in some month.'],
  ['Each month (Apr–Sep) already shows the store the person is rated on: from their own billing where they bill, otherwise their current branch.'],
  ['Only change a month if it is wrong: click the cell and pick the correct store from the dropdown. No typing needed.'],
  ['NOT RATED (top of the dropdown): leave that month out of the person’s rating. Store Managers / ASMs who did not swipe at one store on more than 15 days in a month are pre-filled NOT RATED.'],
  ['"Also Handles": if the person runs a second store together with their main one for the whole of H1 (e.g. one SM for two stores), pick that store. Both stores are then rated together.'],
  ['Rows marked "Moved per billing – please confirm" were detected automatically from the POS data. Leave them if correct.'],
  ['Anyone not in this sheet is rated on their billing store (SA/SSA) or current branch for every month.'],
  ['Save the file and tell Claude to rebuild the dashboard. Re-generating the sheet keeps everything entered here.'],
]);
notes['!cols'] = [{ wch: 130 }];
const st = X.utils.aoa_to_sheet([['Store'], ...storeList.map(s => [s])]);
st['!cols'] = [{ wch: 36 }];
const wb = X.utils.book_new();
X.utils.book_append_sheet(wb, ws, 'Store Map');
X.utils.book_append_sheet(wb, notes, 'How to use');
X.utils.book_append_sheet(wb, st, 'Stores');
const buf = X.write(wb, { type: 'buffer', bookType: 'xlsx' });

// SheetJS can't write data validation, so add the dropdown XML to the Store Map sheet directly (month columns E:J and Also Handles K)
(async () => {
  const zip = await JSZip.loadAsync(buf);
  const f = 'xl/worksheets/sheet1.xml';
  let xml = await zip.file(f).async('string');
  const dv = `<dataValidations count="1"><dataValidation type="list" allowBlank="1" showErrorMessage="1" errorTitle="Pick a store" error="Choose a store from the list." sqref="E2:K${rows.length + 1}"><formula1>Stores!$A$2:$A$${storeList.length + 1}</formula1></dataValidation></dataValidations>`;
  // Excel requires <dataValidations> before these elements (schema order); placing it later makes Excel drop the sheet data
  const after = ['<hyperlinks', '<printOptions', '<pageMargins', '<pageSetup', '<headerFooter', '<ignoredErrors', '<drawing', '<tableParts', '<extLst', '</worksheet>'];
  const at = Math.min(...after.map(t => xml.indexOf(t)).filter(i => i >= 0));
  xml = xml.slice(0, at) + dv + xml.slice(at);
  zip.file(f, xml);
  fs.writeFileSync(OUT, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
  fs.writeFileSync(BASE, JSON.stringify(Object.fromEntries(auto.map(p => [p.id, p.months]))));   // what untouched cells now hold
  const c = rows.reduce((a, r) => (a[r.Check || 'no change'] = (a[r.Check || 'no change'] || 0) + 1, a), {});
  console.log('written', rows.length, 'rows,', storeList.length, 'stores in dropdown', c, 'also-handles:', rows.filter(r => r[ALSO]).map(r => `${r['Emp ID']} → ${r[ALSO]}`));
})();
