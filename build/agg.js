// Stream one month of Item Wise sales (xlsx or csv) into store / employee aggregates.
// usage: node agg.js <file> <monthKey> <outJson>
const fs = require('fs'), { spawn, execFileSync } = require('child_process');
const [file, month, out] = process.argv.slice(2);
const C = { outlet: 0, bill: 2, date: 3, item: 6, qty: 12, net: 39, sm: 40, btype: 56, wot: 112 };
const isEcom = s => !s || /^(e-?com|ecommerce|online)/i.test(s) || !/^\d+$/.test(s);

const stores = {}, emps = {}; let lines = 0, hdr = false;
function add(o, net, wot, qty, bk) { o.net += net; o.wot += wot; o.qty += qty; (o._b[bk] = (o._b[bk] || 0) + net); }
const blank = () => ({ net: 0, wot: 0, qty: 0, _b: {} });
function row(r) {
  if (!hdr) { if (r[0] === 'Outlet Name') hdr = true; return; }
  const outlet = (r[C.outlet] || '').trim(); if (!outlet || !r[C.bill]) return;
  const net = +r[C.net] || 0, wot = +r[C.wot] || 0, qty = +r[C.qty] || 0;
  const sm = String(r[C.sm] || '').trim(), bk = outlet + '|' + r[C.bill];
  lines++;
  const s = stores[outlet] || (stores[outlet] = { all: blank(), inst: blank() });
  add(s.all, net, wot, qty, bk);
  if (!isEcom(sm)) {
    add(s.inst, net, wot, qty, bk);
    const e = (emps[sm] || (emps[sm] = {}));
    add(e[outlet] || (e[outlet] = blank()), net, wot, qty, bk);
  }
}
function fin(o) { // bills = distinct bills with positive value for that scope
  o.bills = Object.values(o._b).filter(v => v > 0).length; delete o._b;
  o.net = Math.round(o.net); o.wot = Math.round(o.wot); return o;
}
function done() {
  for (const s of Object.values(stores)) { fin(s.all); fin(s.inst); }
  for (const e of Object.values(emps)) for (const o of Object.values(e)) fin(o);
  fs.writeFileSync(out, JSON.stringify({ month, lines, stores, emps }));
  console.log(month, 'lines', lines, 'stores', Object.keys(stores).length, 'emps', Object.keys(emps).length);
}

if (/\.csv$/i.test(file)) {
  // RFC4180 parser over a stream (quoted fields may hold commas/newlines)
  let f = [], cur = '', q = false;
  const st = fs.createReadStream(file, { encoding: 'utf8', highWaterMark: 1 << 22 });
  st.on('data', ch => {
    for (let i = 0; i < ch.length; i++) {
      const c = ch[i];
      if (q) { if (c === '"') { if (ch[i + 1] === '"') { cur += '"'; i++; } else q = false; } else cur += c; }
      else if (c === '"') q = true;
      else if (c === ',') { f.push(cur); cur = ''; }
      else if (c === '\n') { f.push(cur.replace(/\r$/, '')); row(f); f = []; cur = ''; }
      else cur += c;
    }
  });
  st.on('end', () => { if (cur || f.length) { f.push(cur); row(f); } done(); });
} else {
  const ssXml = execFileSync('unzip', ['-p', file, 'xl/sharedStrings.xml'], { maxBuffer: 1 << 30 }).toString('utf8');
  const dec = s => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
  const ss = [];
  for (const m of ssXml.matchAll(/<si>([\s\S]*?)<\/si>/g)) ss.push(dec([...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map(x => x[1]).join('')));
  const colIdx = l => { let n = 0; for (const ch of l) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
  const p = spawn('unzip', ['-p', file, 'xl/worksheets/sheet1.xml']);
  p.stdout.setEncoding('utf8'); let buf = '';
  const cellRe = /<c r="([A-Z]+)\d+"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
  p.stdout.on('data', ch => {
    buf += ch; let k;
    while ((k = buf.indexOf('</row>')) >= 0) {
      const rx = buf.slice(0, k); buf = buf.slice(k + 6);
      const r = [];
      for (const m of rx.matchAll(cellRe)) {
        const body = m[3]; if (body == null) continue;
        const t = /t="(\w+)"/.exec(m[2]); const v = /<v>([\s\S]*?)<\/v>/.exec(body);
        let val;
        if (t && t[1] === 's') val = ss[+v[1]];
        else if (t && t[1] === 'inlineStr') val = dec((/<t[^>]*>([\s\S]*?)<\/t>/.exec(body) || [, ''])[1]);
        else val = v ? dec(v[1]) : '';
        r[colIdx(m[1])] = val;
      }
      row(r);
    }
  });
  p.on('close', done);
}
