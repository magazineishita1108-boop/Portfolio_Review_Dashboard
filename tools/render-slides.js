/* ============================================================================
   render-slides.js — build any slide group headlessly, so slides can be
   LOOKED AT rather than only measured.

   Every PPTX geometry bug in this project so far survived a numeric check and
   was caught by a human opening the file. The browser pane blocks downloads,
   so there was no way to see a slide from here. This runs the shipped engine
   and exporters in Node against a real base file and writes a .pptx;
   PowerPoint COM then exports PNGs from it.

   It runs the very script blocks index.html ships — not a copy — so what it
   renders is what the browser renders.

   Usage:
     node tools/render-slides.js <base.xlsx> <irr.xlsx|-> <out.pptx> [group] [client]
   e.g.
     node tools/render-slides.js base.xlsx -        out.pptx aia
     node tools/render-slides.js base.xlsx irr.xlsx out.pptx tax "Pravin Jain"
   ========================================================================== */
'use strict';
const fs = require('fs');

const [, , BASE, IRR, OUT, GROUP, CLIENT] = process.argv;
if (!BASE || !IRR || !OUT) {
  console.error('usage: node tools/render-slides.js <base.xlsx> <irr.xlsx|-> <out.pptx> [group] [client]');
  process.exit(2);
}

/* One loader, shared with tools/test-ai-analyst.js. Two copies of a
   block-extraction routine drift apart, and this one is order-sensitive. */
const { load } = require('./engine-host.js');
const _H = load();
const XLSX = _H.XLSX;
const Ppt = _H.Ppt;
const E = _H.E;
if (!XLSX || !Ppt || !E || !E.buildPPTX) throw new Error('globals missing after eval');

/* ---- parse the two workbooks exactly as the dashboard does ---- */
function aoaOf(wb, sheet) { return XLSX.utils.sheet_to_json(wb.Sheets[sheet], { header: 1, raw: true }); }

const wbB = XLSX.read(fs.readFileSync(BASE), { type: 'buffer', cellDates: true });
const shB = wbB.Sheets['Base File'] ? 'Base File' : wbB.SheetNames[0];
const rows = E.parseBaseRows(aoaOf(wbB, shB));
if (!rows.length) throw new Error('no holdings parsed from ' + BASE);

/* the IRR sheet is optional: pass "-" for a group that does not need lot history */
let irr = null;
if (IRR && IRR !== '-') {
  const wbI = XLSX.read(fs.readFileSync(IRR), { type: 'buffer', cellDates: true });
  irr = E.parsers.parseIRR(wbI.SheetNames, n => aoaOf(wbI, n));
  if (!irr.fund.length) throw new Error('no IRR transactions recognised');
}

const REF = irr ? Object.assign({}, _H.REF, { IRR: irr }) : _H.REF;
const model = E.computeReview(rows, REF, { familyName: 'Render Check', asOnDate: '31 Jul 2026' });
const group = GROUP || 'tax';

console.log('holdings parsed   : ' + rows.length);
if (group === 'tax') {
  const T = model.taxation;
  if (!T || !T.hasData) throw new Error('taxation not computed: ' + (T && T.error));
  console.log('positions         : ' + T.positions.length + '  unreconciled ' + T.unreconciled.length);
  global.__taxScope = CLIENT || null;
  console.log('scope             : ' + (CLIENT || 'Family'));
}
if (group === 'aia') {
  /* the slide group reads the plan off window.__AIA_DATA, which the browser's
     renderAIAnalyst puts there. Headless there is no DOM, so the harness
     computes the same plan through the same engine entry point and publishes
     it the same way — otherwise the group silently renders nothing. */
  const plan = E.aiaCompute(model, REF, {});
  if (!plan.ok) throw new Error('AI Analyst did not produce a plan: ' + (plan.reason || '?'));
  global.__AIA_DATA = plan;
  console.log('recommendations   : ' + plan.recs.length + '  not assessed ' + plan.notAssessed.length);
  console.log('released          : ' + plan.summary.released + ' Cr across ' + plan.summary.sources + ' holdings');
}

const p = E.buildPPTX(model, Ppt, group);
Promise.resolve(p.write({ outputType: 'nodebuffer' })).then(buf => {
  fs.writeFileSync(OUT, Buffer.from(buf));
  console.log('wrote             : ' + OUT + '  (' + Buffer.from(buf).length.toLocaleString() + ' bytes)');
}).catch(e => { console.error('write failed: ' + (e && e.message || e)); process.exit(1); });
