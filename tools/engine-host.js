/* ============================================================================
   engine-host.js — run the SHIPPED engine headlessly.
   Evaluates the script blocks straight out of index.html rather than a copy,
   so a test exercises the same code the browser runs. Blocks are located by
   their own marker strings, never by index: block order shifts.
   ========================================================================== */

'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const ROOT = 'C:/Users/IshitaMagazine/OneDrive - CENTRICITY FINANCIAL DISTRIBUTION PRIVATE LIMITED/Documents/Claude/Projects/Portfolio Review';

function load() {
  const text = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');
  const blocks = [];
  const re = /<script([^>]*)>/g; let m;
  while ((m = re.exec(text))) { const e = text.indexOf('</script>', m.index); if (e < 0) break;
    blocks.push(text.slice(m.index + m[0].length, e)); re.lastIndex = e; }
  const at = p => blocks.findIndex(p);
  const iX = at(b => /xlsx\.js \(C\) 2013-present SheetJS/.test(b.slice(0, 200)));
  const iP = at(b => /PptxGenJS 3\.\d/.test(b.slice(0, 200)));
  const iR = at(b => /^\s*window\.REFDATA\s*=/.test(b.slice(0, 60)));
  const iD = at(b => /^\s*var DIVIDERS=/.test(b.slice(0, 40)));
  const iL = at(b => /^\s*var BRAND_LOGOS=/.test(b.slice(0, 40)));
  const iE = at(b => /Pure compute \+ PPTX\/XLSX builders/.test(b.slice(0, 900)));
  const iXp = at(b => /EXPORTERS \(PPTX via PptxGenJS/.test(b.slice(0, 900)));
  const iPa = at(b => /RUNTIME REFERENCE PARSERS/.test(b.slice(0, 900)));
  const iEr = at(b => /^\s*window\.MF_ER_TABLE=/.test(b.slice(0, 40)));

  global.window = global; global.self = global;
  try { global.navigator = { userAgent: 'node' }; } catch (e) {}
  const el = () => ({ style: {}, classList: { add() {}, remove() {}, contains: () => false },
    appendChild() {}, removeChild() {}, setAttribute() {}, getAttribute: () => null,
    addEventListener() {}, removeEventListener() {}, querySelector: () => null,
    querySelectorAll: () => [], getContext: () => null, children: [], innerHTML: '', textContent: '' });
  global.document = { createElement: el, createElementNS: el, getElementById: () => null,
    querySelector: () => null, querySelectorAll: () => [], addEventListener() {},
    body: el(), head: el(), documentElement: el() };
  global.filterPerfBuckets = b => b || [];
  [iX, iP, iP + 1, iR, iD, iL, iE, iXp, iPa, iEr].forEach(i => vm.runInThisContext(blocks[i]));
  return { XLSX: global.XLSX, Ppt: global.PptxGenJS, E: global.ENGINE, REF: global.REFDATA };
}

function build(baseFile, meta) {
  const { XLSX, E, REF, Ppt } = load();
  const wb = XLSX.read(fs.readFileSync(baseFile), { type: 'buffer', cellDates: true });
  const sh = wb.Sheets['Base File'] ? 'Base File' : wb.SheetNames[0];
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sh], { header: 1, raw: true });
  const rows = E.parseBaseRows(aoa);
  if (!rows.length) throw new Error('no holdings parsed from ' + baseFile);
  const model = E.computeReview(rows, REF, meta || { familyName: 'Probe Family', asOnDate: '31 Jul 2026' });
  return { model, rows, E, REF, XLSX, Ppt };
}

module.exports = { load, build, ROOT };
