/* ============================================================================
   parse-fund-mgr.js — build REFDATA.FUND_MGR from the monthly
   "FUND MANAGER HISTORY.xlsx" issue.

   Monthly refresh input. Columns resolved BY HEADER NAME and asserted present.

   Two things about this file decide the whole design:

   1. The body is GROUPED: a separator row reading "Scheme Name: <name>"
      precedes each scheme's stints. Those rows carry no data and would enter
      the block as phantom schemes if not skipped.

   2. The file's maximum "To Date" is the data as-of date, NOT a real end date.
      Every stint carrying it is a currently-serving manager. `cur` is therefore
      computed HERE, once, against this file's own maximum — never inferred in
      the browser. A partial issue would shift a browser-side inference and
      silently reclassify every manager in the block.

   Usage:
     node tools/parse-fund-mgr.js <FUND MANAGER HISTORY.xlsx> [--out mgr.json]
   ========================================================================== */
'use strict';
const fs = require('fs');
const path = require('path');

function loadEngine() {
  const vm = require('vm');
  const IDX = path.join(__dirname, '..', 'index.html');
  const text = fs.readFileSync(IDX, 'utf8');
  const blocks = [];
  const re = /<script([^>]*)>/g; let m;
  while ((m = re.exec(text))) { const e = text.indexOf('</script>', m.index); if (e < 0) break;
    blocks.push(text.slice(m.index + m[0].length, e)); re.lastIndex = e; }
  const pick = p => { const i = blocks.findIndex(p); if (i < 0) throw new Error('script block not found'); return blocks[i]; };
  global.window = global;
  vm.runInThisContext(pick(b => /xlsx\.js \(C\) 2013-present SheetJS/.test(b.slice(0, 200))));
  vm.runInThisContext(pick(b => /^\s*window\.REFDATA\s*=/.test(b.slice(0, 60))));
  vm.runInThisContext(pick(b => /Pure compute \+ PPTX\/XLSX builders/.test(b.slice(0, 900))));
  return { XLSX: global.XLSX, REF: global.REFDATA, E: global.ENGINE };
}

const COLS = { name: 'Scheme Name', inc: 'Inception Date', mgr: 'Fund Manager', from: 'From Date', to: 'To Date' };
const EXPECT = { stints: 539, schemes: 104, floorStints: 400, floorSchemes: 80 };

/* Empty value blocks a name from matching, as in the engine's ALIASES. The two
   IDCW entries are blocked for the same reason as in parse-fund-ranks.js: the
   fuzzy matcher resolves "…-Reg(IDCW)" to "…-Reg(G)" on a single token, and a
   manager history is per-scheme but the record would land on the wrong plan's
   key and displace the Growth plan's own history if it later arrives. */
const MGR_ALIASES = {
  'Aditya Birla SL ELSS Tax Saver Fund(IDCW)': '',
  'Aditya Birla SL Large & Mid Cap Fund-Reg(IDCW)': ''
};
function planOf(s) {
  const m = /\((idcw|g|d|dp|bonus)[^)]*\)\s*$/i.exec(String(s || ''));
  return m ? m[1].toLowerCase() : '';
}
function colIndex(header, label) {
  const want = String(label).trim().toLowerCase();
  for (let i = 0; i < header.length; i++) {
    if (String(header[i] == null ? '' : header[i]).trim().toLowerCase() === want) return i;
  }
  return -1;
}
const norm = s => String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]+/g, '');
/* Excel serial -> ISO. Serials only; a string date in these columns is a
   format change and must be caught, not coerced. */
function serToIso(v) {
  if (typeof v !== 'number' || !isFinite(v)) return null;
  return new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000).toISOString().slice(0, 10);
}

function findHeaderRow(aoa) {
  for (let i = 0; i < Math.min(aoa.length, 10); i++) {
    const r = aoa[i] || [];
    if (colIndex(r, COLS.name) >= 0 && colIndex(r, COLS.mgr) >= 0 && colIndex(r, COLS.from) >= 0) return i;
  }
  return -1;
}

function build(file, opts) {
  opts = opts || {};
  const { XLSX, REF, E } = loadEngine();
  const wb = XLSX.read(fs.readFileSync(file), { type: 'buffer', raw: true });
  const sn = wb.SheetNames[0];
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true });

  const hr = findHeaderRow(aoa);
  if (hr < 0) throw new Error('header row not found in sheet "' + sn + '" (looked at the first 10 rows)');
  const header = aoa[hr];
  const ix = {}; const missing = [];
  Object.keys(COLS).forEach(k => { ix[k] = colIndex(header, COLS[k]); if (ix[k] < 0) missing.push(COLS[k]); });
  if (missing.length) throw new Error('columns missing: ' + missing.join(' | ') + '\nheader was: ' + JSON.stringify(header));

  /* ---- read the stints, skipping the grouping separators ---- */
  const raw = []; let seps = 0, badDate = 0;
  aoa.slice(hr + 1).forEach(r => {
    if (!r) return;
    const a = r[ix.name];
    if (a == null || String(a).trim() === '') return;
    if (/^\s*Scheme Name\s*:/i.test(String(a))) { seps++; return; }      /* grouping row, no data */
    const mgr = r[ix.mgr];
    if (mgr == null || String(mgr).trim() === '') return;
    if (typeof r[ix.from] !== 'number' || typeof r[ix.to] !== 'number') { badDate++; return; }
    raw.push({ name: String(a).trim(), mgr: String(mgr).trim(), inc: r[ix.inc], from: r[ix.from], to: r[ix.to] });
  });
  if (badDate) throw new Error(badDate + ' stint row(s) carry a non-serial date — the issue format has changed, refusing to guess');

  /* ---- the file's own maximum To Date IS the as-of date ---- */
  let maxTo = -Infinity;
  raw.forEach(s => { if (s.to > maxTo) maxTo = s.to; });
  const asOf = serToIso(maxTo);

  const bySrc = {};
  raw.forEach(s => (bySrc[s.name] = bySrc[s.name] || []).push(s));
  const srcNames = Object.keys(bySrc);

  if (raw.length < EXPECT.floorStints) throw new Error('only ' + raw.length + ' stints (prior issue ' + EXPECT.stints + ', floor ' + EXPECT.floorStints + ')');
  if (srcNames.length < EXPECT.floorSchemes) throw new Error('only ' + srcNames.length + ' schemes (prior issue ' + EXPECT.schemes + ', floor ' + EXPECT.floorSchemes + ')');

  /* ---- resolve names to the dashboard's spelling ---- */
  const cands = Object.keys(REF.MF_PERF || {});
  const candNorm = {}; cands.forEach(c => candNorm[c] = E.normalize(c));
  const byNorm = {}; cands.forEach(c => { const n = norm(c); if (!(n in byNorm)) byNorm[n] = c; });

  const out = {};
  let exact = 0, aliased = 0, fuzzy = 0, blocked = 0, unmatched = 0, curTotal = 0;
  const fuzzyPairs = [], planRejects = [], unmatchedNames = [];

  srcNames.forEach(src => {
    let key = null, how = '';
    const isBlocked = Object.prototype.hasOwnProperty.call(MGR_ALIASES, src) && MGR_ALIASES[src] === '';
    if (byNorm[norm(src)]) { key = byNorm[norm(src)]; how = 'exact'; exact++; }
    else if (isBlocked) { blocked++; }
    else if (MGR_ALIASES[src] && byNorm[norm(MGR_ALIASES[src])]) { key = byNorm[norm(MGR_ALIASES[src])]; how = 'alias'; aliased++; }
    else {
      const amcTok = src.split(/[\s-]+/)[0] || '';
      const hit = E.matchInstrument(src, amcTok, cands, candNorm);
      if (hit && planOf(src) && planOf(hit) && planOf(src) !== planOf(hit)) planRejects.push([src, hit]);
      else if (hit) { key = hit; how = 'fuzzy'; fuzzy++; fuzzyPairs.push([src, hit]); }
    }
    if (!key) { key = src; unmatched++; unmatchedNames.push(src); how = 'unmatched'; }
    if (Object.prototype.hasOwnProperty.call(out, key)) {
      if (!opts.quiet) console.warn('  ! collision on "' + key + '" — keeping first, ignoring "' + src + '"');
      return;
    }

    const list = bySrc[src];
    let inc = null;
    for (const s of list) { const v = serToIso(s.inc); if (v) { inc = v; break; } }
    const stints = list.map(s => {
      const rec = { m: s.mgr, from: serToIso(s.from), to: serToIso(s.to), cur: s.to === maxTo };
      if (rec.cur) curTotal++;
      return rec;
    })
      /* newest first: the tab reads stints[0] constantly */
      .sort((a, b) => (a.from < b.from ? 1 : a.from > b.from ? -1 : (a.m < b.m ? -1 : 1)));

    const rec = { stints: stints };
    if (inc) rec.inc = inc;
    if (how !== 'exact' && how !== 'unmatched') rec.src = src;
    out[key] = rec;
  });

  const meta = { asOf: asOf, schemes: Object.keys(out).length, stints: raw.length, serving: curTotal };
  if (!opts.quiet) {
    console.log('sheet          : ' + sn + '   header row ' + (hr + 1));
    console.log('separators     : ' + seps + ' skipped');
    console.log('stints         : ' + raw.length + '   schemes ' + srcNames.length);
    console.log('as-of (max To) : ' + asOf + '   currently-serving stints ' + curTotal);
    console.log('name resolution: exact ' + exact + '  alias ' + aliased + '  fuzzy ' + fuzzy + '  blocked ' + blocked + '  unmatched ' + unmatched);
    if (fuzzyPairs.length) { console.log('fuzzy matches:'); fuzzyPairs.forEach(p => console.log('   ' + p[0] + '\n     -> ' + p[1])); }
    if (planRejects.length) { console.log('rejected (would change the plan):'); planRejects.forEach(p => console.log('   ' + p[0] + '\n     x-> ' + p[1])); }
    if (unmatchedNames.length) { console.log('unmatched:'); unmatchedNames.forEach(n => console.log('   ' + n)); }
  }
  return { FUND_MGR: out, FUND_MGR_META: meta,
    stats: { exact, aliased, fuzzy, blocked, unmatched, unmatchedNames, fuzzyPairs, planRejects } };
}

module.exports = { build };

if (require.main === module) {
  const args = process.argv.slice(2);
  const file = args.find(a => !a.startsWith('--'));
  if (!file) { console.error('usage: node tools/parse-fund-mgr.js <FUND MANAGER HISTORY.xlsx> [--out mgr.json]'); process.exit(2); }
  const oi = args.indexOf('--out');
  const r = build(file, { quiet: args.indexOf('--quiet') >= 0 });
  if (oi >= 0 && args[oi + 1]) {
    fs.writeFileSync(args[oi + 1], JSON.stringify({ FUND_MGR: r.FUND_MGR, FUND_MGR_META: r.FUND_MGR_META }));
    console.log('wrote ' + args[oi + 1]);
  }
}
