/* ============================================================================
   render-tax-slides.js — build the taxation deck headlessly, so slides can be
   LOOKED AT rather than only measured.

   Every PPTX geometry bug in this project so far survived a numeric check and
   was caught by a human opening the file. The browser pane blocks downloads,
   so there was no way to see a slide from here. This runs the shipped engine
   and exporters in Node against a real base file + IRR sheet and writes a
   .pptx; PowerPoint COM then exports PNGs from it.

   It evaluates the very script blocks index.html ships — not a copy — so what
   it renders is what the browser renders. Blocks are located by their own
   marker strings, never by index, because block order shifts.

   Usage:
     node tools/render-tax-slides.js <base.xlsx> <irr.xlsx> <out.pptx> [client]
   ========================================================================== */
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const [, , BASE, IRR, OUT, CLIENT] = process.argv;
if (!BASE || !IRR || !OUT) {
  console.error('usage: node tools/render-tax-slides.js <base.xlsx> <irr.xlsx> <out.pptx> [client]');
  process.exit(2);
}

const HTML = path.join(__dirname, '..', 'index.html');
const text = fs.readFileSync(HTML, 'utf8');

/* ---- carve out every inline script body, in source order ---- */
const blocks = [];
{
  const re = /<script([^>]*)>/g;
  let m;
  while ((m = re.exec(text))) {
    const end = text.indexOf('</script>', m.index);
    if (end < 0) break;
    blocks.push(text.slice(m.index + m[0].length, end));
    re.lastIndex = end;
  }
}

/* Blocks are chosen by what they contain. The PptxGenJS bundle embeds the
   literal string "</script>", which splits it across two blocks — both halves
   are needed and must run back to back. */
function findBlock(pred, label) {
  const i = blocks.findIndex(pred);
  if (i < 0) throw new Error('script block not found: ' + label);
  return i;
}
const iXlsx = findBlock(b => /xlsx\.js \(C\) 2013-present SheetJS/.test(b.slice(0, 200)), 'xlsx');
const iPptx = findBlock(b => /PptxGenJS 3\.\d/.test(b.slice(0, 200)), 'pptxgenjs');
const iRef = findBlock(b => /^\s*window\.REFDATA\s*=/.test(b.slice(0, 60)), 'REFDATA');
const iDiv = findBlock(b => /^\s*var DIVIDERS=/.test(b.slice(0, 40)), 'DIVIDERS');
const iLogo = findBlock(b => /^\s*var BRAND_LOGOS=/.test(b.slice(0, 40)), 'BRAND_LOGOS');
const iEng = findBlock(b => /Pure compute \+ PPTX\/XLSX builders/.test(b.slice(0, 900)), 'engine');
const iExp = findBlock(b => /EXPORTERS \(PPTX via PptxGenJS/.test(b.slice(0, 900)), 'exporters');
const iPar = findBlock(b => /RUNTIME REFERENCE PARSERS/.test(b.slice(0, 900)), 'parsers');
const iEr = findBlock(b => /^\s*window\.MF_ER_TABLE=/.test(b.slice(0, 40)), 'MF_ER_TABLE');

/* ---- the browser globals these blocks expect ---- */
global.window = global;
global.self = global;
/* node >=21 defines navigator as a getter-only global */
try { global.navigator = { userAgent: 'node' }; } catch (e) { /* keep node's own */ }
global.location = { href: 'file:///index.html', protocol: 'file:' };
const elStub = () => ({
  style: {}, classList: { add() {}, remove() {}, contains: () => false },
  appendChild() {}, removeChild() {}, setAttribute() {}, getAttribute: () => null,
  addEventListener() {}, removeEventListener() {}, querySelector: () => null,
  querySelectorAll: () => [], getContext: () => null, children: [], innerHTML: '', textContent: ''
});
global.document = {
  createElement: elStub, createElementNS: elStub, getElementById: () => null,
  querySelector: () => null, querySelectorAll: () => [], addEventListener() {},
  body: elStub(), head: elStub(), documentElement: elStub()
};
global.filterPerfBuckets = b => b || [];

const order = [iXlsx, iPptx, iPptx + 1, iRef, iDiv, iLogo, iEng, iExp, iPar, iEr];
const names = ['xlsx', 'pptx(a)', 'pptx(b)', 'REFDATA', 'DIVIDERS', 'BRAND_LOGOS', 'engine', 'exporters', 'parsers', 'MF_ER'];
order.forEach((i, k) => {
  try { vm.runInThisContext(blocks[i], { filename: 'index.html#' + names[k] }); }
  catch (e) { throw new Error('block ' + names[k] + ' failed: ' + e.message); }
});

const XLSX = global.XLSX;
const Ppt = global.PptxGenJS;
const E = global.ENGINE;
if (!XLSX || !Ppt || !E || !E.buildPPTX) throw new Error('globals missing after eval');

/* ---- parse the two workbooks exactly as the dashboard does ---- */
function aoaOf(wb, sheet) { return XLSX.utils.sheet_to_json(wb.Sheets[sheet], { header: 1, raw: true }); }

const wbB = XLSX.read(fs.readFileSync(BASE), { type: 'buffer', cellDates: true });
const shB = wbB.Sheets['Base File'] ? 'Base File' : wbB.SheetNames[0];
const rows = E.parseBaseRows(aoaOf(wbB, shB));
if (!rows.length) throw new Error('no holdings parsed from ' + BASE);

const wbI = XLSX.read(fs.readFileSync(IRR), { type: 'buffer', cellDates: true });
const irr = E.parsers.parseIRR(wbI.SheetNames, n => aoaOf(wbI, n));
if (!irr.fund.length) throw new Error('no IRR transactions recognised');

const REF = Object.assign({}, global.REFDATA, { IRR: irr });
const model = E.computeReview(rows, REF, { familyName: 'Render Check', asOnDate: '31 Jul 2026' });
const T = model.taxation;
if (!T || !T.hasData) throw new Error('taxation not computed: ' + (T && T.error));

global.__taxScope = CLIENT || null;

console.log('holdings parsed   : ' + rows.length);
console.log('positions         : ' + T.positions.length + '  unreconciled ' + T.unreconciled.length);
console.log('scope             : ' + (CLIENT || 'Family'));

const p = E.buildPPTX(model, Ppt, 'tax');
Promise.resolve(p.write({ outputType: 'nodebuffer' })).then(buf => {
  fs.writeFileSync(OUT, Buffer.from(buf));
  console.log('wrote             : ' + OUT + '  (' + Buffer.from(buf).length.toLocaleString() + ' bytes)');
}).catch(e => { console.error('write failed: ' + (e && e.message || e)); process.exit(1); });
