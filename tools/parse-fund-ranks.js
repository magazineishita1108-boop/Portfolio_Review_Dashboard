/* ============================================================================
   parse-fund-ranks.js — build REFDATA.FUND_RANKS from the monthly
   "Fund_Ranks_<date>.xlsx" issue.

   Monthly refresh input. Every column is resolved BY HEADER NAME and asserted
   present: tools/parse-performance.js resolved by name too and still dropped
   889 debt sub-blocks for a full cycle by simply never asking for three
   columns. A silent partial parse is the documented failure mode of this repo,
   so this one fails loudly instead.

   Row counts are asserted against the previous issue and a material drop is an
   error, not a warning — the PMS analytics workbook was re-issued mid-refresh
   at 10 schemes and then 15, and "check its size or hash before assuming a
   repeat request is a no-op".

   Usage:
     node tools/parse-fund-ranks.js <Fund_Ranks.xlsx> [--out ranks.json] [--quiet]
   ========================================================================== */
'use strict';
const fs = require('fs');
const path = require('path');

/* ---------- the engine, lifted out of the shipped file ---------- */
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

/* ---------- expected shape of the issue ---------- */
const COLS = {
  name: 'Fund Name', code: 'Scheme Code', cat: 'Category',
  rank: 'Rank in Category', of: 'Funds in Category', score: 'Score (/100)', status: 'Status'
};
/* Prior issue (15 Aug 2026). A drop below the floor is an error. */
const EXPECT = { rows: 1262, ranked: 524, floorRows: 1100, floorRanked: 450 };

/* Names the ranks file spells differently from the dashboard. Each was looked
   at individually; none lowers the matcher's threshold.

   An empty value BLOCKS a name from matching at all — the same convention
   ALIASES uses in the engine. The two IDCW entries are blocked deliberately:
   the ranks file carries the IDCW plan of those schemes and no Growth plan,
   and the fuzzy matcher happily resolves "…-Reg(IDCW)" to "…-Reg(G)" because
   the plan suffix is one token out of many. IDCW and Growth are different
   plans with different NAVs and different trailing returns, so that join would
   have attached one plan's category rank to the other plan's record. A rank
   that is absent is correct; a rank borrowed from a sibling plan is not. */
const RANK_ALIASES = {
  'SBI Healthcare Opportunities Fund-Reg(G)': 'SBI Healthcare Opp Fund-Reg(G)',
  'SBI Technology Opportunities Fund-Reg(G)': 'SBI Technology Opp Fund-Reg(G)',
  'SBI Consumption Opportunities Fund-Reg(G)': 'SBI Consumption Opp Fund-Reg(G)',
  'ICICI Pru Equity-Arbitrage Fund(G)': 'ICICI Pru Arbitrage Fund(G)',
  'SBI Equity Hybrid Fund-Reg(IDCW)': '',
  'Kotak Aggressive Hybrid Fund-Reg(IDCW)': ''
};
/* a resolved name must not change the plan: (G) never becomes (IDCW) or (D) */
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
const num = v => (typeof v === 'number' && isFinite(v)) ? v : null;

function build(file, opts) {
  opts = opts || {};
  const { XLSX, REF, E } = loadEngine();
  const wb = XLSX.read(fs.readFileSync(file), { type: 'buffer', raw: true });
  if (wb.SheetNames.indexOf('Fund Ranks') < 0) throw new Error('sheet "Fund Ranks" missing; found: ' + wb.SheetNames.join(', '));
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets['Fund Ranks'], { header: 1, raw: true });
  const header = aoa[0] || [];

  /* resolve every column by header name, and fail loudly on any that is absent */
  const ix = {}; const missing = [];
  Object.keys(COLS).forEach(k => { ix[k] = colIndex(header, COLS[k]); if (ix[k] < 0) missing.push(COLS[k]); });
  if (missing.length) throw new Error('columns missing from "Fund Ranks": ' + missing.join(' | ') + '\nheader was: ' + JSON.stringify(header));

  const body = aoa.slice(1).filter(r => r && r[ix.name] != null && String(r[ix.name]).trim() !== '');

  /* ---- candidate universe: the dashboard's own scheme names ---- */
  const cands = Object.keys(REF.MF_PERF || {});
  const candNorm = {}; cands.forEach(c => candNorm[c] = E.normalize(c));
  const byNorm = {}; cands.forEach(c => { const n = norm(c); if (!(n in byNorm)) byNorm[n] = c; });

  const out = {}, stat = {}, cat = {};
  const codes = new Set();
  let ranked = 0, exact = 0, fuzzy = 0, aliased = 0, unmatched = 0, blockedN = 0;
  const unmatchedNames = [], fuzzyPairs = [], planRejects = [];

  body.forEach(r => {
    const src = String(r[ix.name]).trim();
    const st = (r[ix.status] == null || String(r[ix.status]).trim() === '') ? '' : String(r[ix.status]).trim();
    const rk = num(r[ix.rank]), of = num(r[ix.of]), sc = num(r[ix.score]);
    if (rk != null) ranked++;
    stat[st || '(blank)'] = (stat[st || '(blank)'] || 0) + 1;
    if (r[ix.cat] != null) cat[String(r[ix.cat]).trim()] = 1;
    if (r[ix.code] != null) codes.add(String(r[ix.code]).trim());

    /* resolve to the dashboard's spelling: exact, then alias, then the shipped
       fuzzy matcher with the brand anchor supplied. The threshold is never
       lowered — CLAUDE.md documents what happens at 2. */
    let key = null, how = '';
    const blocked = Object.prototype.hasOwnProperty.call(RANK_ALIASES, src) && RANK_ALIASES[src] === '';
    if (byNorm[norm(src)]) { key = byNorm[norm(src)]; how = 'exact'; exact++; }
    else if (blocked) { blockedN++; }
    else if (RANK_ALIASES[src] && byNorm[norm(RANK_ALIASES[src])]) { key = byNorm[norm(RANK_ALIASES[src])]; how = 'alias'; aliased++; }
    else {
      const amcTok = src.split(/[\s-]+/)[0] || '';
      const hit = E.matchInstrument(src, amcTok, cands, candNorm);
      /* reject a match that silently changes the plan */
      if (hit && planOf(src) && planOf(hit) && planOf(src) !== planOf(hit)) {
        planRejects.push([src, hit]);
      } else if (hit) { key = hit; how = 'fuzzy'; fuzzy++; fuzzyPairs.push([src, hit]); }
    }
    if (!key) { key = src; unmatched++; unmatchedNames.push(src); how = 'unmatched'; }

    /* first writer wins: two ranks rows resolving to one dashboard name would
       otherwise silently overwrite each other */
    if (Object.prototype.hasOwnProperty.call(out, key)) {
      if (!opts.quiet) console.warn('  ! collision on "' + key + '" — keeping first, ignoring "' + src + '"');
      return;
    }
    const rec = { code: num(r[ix.code]) != null ? num(r[ix.code]) : String(r[ix.code] || ''), cat: String(r[ix.cat] || '').trim() };
    /* unranked rows carry code + cat + status only: no rank/of/score keys at all,
       so "absent" is distinguishable from "zero" downstream */
    if (rk != null) rec.rank = rk;
    if (of != null) rec.of = of;
    if (sc != null) rec.score = Math.round(sc * 100) / 100;
    if (st) rec.status = st;
    if (how !== 'exact' && how !== 'unmatched') rec.src = src;   /* audit trail for a resolved name */
    out[key] = rec;
  });

  /* ---- row-count assertions ---- */
  const errs = [];
  if (body.length < EXPECT.floorRows) errs.push('only ' + body.length + ' data rows (prior issue ' + EXPECT.rows + ', floor ' + EXPECT.floorRows + ')');
  if (ranked < EXPECT.floorRanked) errs.push('only ' + ranked + ' ranked rows (prior issue ' + EXPECT.ranked + ', floor ' + EXPECT.floorRanked + ')');
  if (codes.size !== body.length) errs.push('scheme codes not unique: ' + codes.size + ' distinct over ' + body.length + ' rows');
  if (errs.length) throw new Error('issue looks wrong, refusing to build:\n  - ' + errs.join('\n  - '));

  /* cycle date from the Notes sheet, never guessed from the filename */
  let cycle = '';
  if (wb.SheetNames.indexOf('Notes') >= 0) {
    const nt = XLSX.utils.sheet_to_json(wb.Sheets['Notes'], { header: 1, raw: true });
    nt.forEach(r => { if (r && String(r[0] || '').trim().toLowerCase() === 'cycle' && r[1] != null) cycle = String(r[1]).trim().slice(0, 10); });
  }
  if (!cycle) throw new Error('Notes!Cycle not found — the vintage must come from the file, not the filename');

  const meta = { cycle: cycle, listed: body.length, ranked: ranked, unranked: body.length - ranked };

  if (!opts.quiet) {
    console.log('rows           : ' + body.length + '   ranked ' + ranked + '   unranked ' + (body.length - ranked));
    console.log('cycle          : ' + cycle);
    console.log('categories     : ' + Object.keys(cat).length);
    console.log('status values  : ' + Object.keys(stat).map(k => k + ' ' + stat[k]).join('  |  '));
    console.log('name resolution: exact ' + exact + '  alias ' + aliased + '  fuzzy ' + fuzzy + '  blocked ' + blockedN + '  unmatched ' + unmatched);
    if (planRejects.length) { console.log('rejected (would change the plan):'); planRejects.forEach(pr => console.log('   ' + pr[0] + '\n     x-> ' + pr[1])); }
    if (fuzzyPairs.length) { console.log('fuzzy matches:'); fuzzyPairs.forEach(p => console.log('   ' + p[0] + '\n     -> ' + p[1])); }
    if (unmatchedNames.length) { console.log('unmatched (kept under their own name, joinable to nothing):');
      unmatchedNames.forEach(n => console.log('   ' + n)); }
  }
  return { FUND_RANKS: out, FUND_RANKS_META: meta,
    stats: { exact, aliased, fuzzy, blocked: blockedN, unmatched, unmatchedNames, fuzzyPairs, planRejects, status: stat } };
}

module.exports = { build };

if (require.main === module) {
  const args = process.argv.slice(2);
  const file = args.find(a => !a.startsWith('--'));
  if (!file) { console.error('usage: node tools/parse-fund-ranks.js <Fund_Ranks.xlsx> [--out ranks.json]'); process.exit(2); }
  const oi = args.indexOf('--out');
  const r = build(file, { quiet: args.indexOf('--quiet') >= 0 });
  if (oi >= 0 && args[oi + 1]) {
    fs.writeFileSync(args[oi + 1], JSON.stringify({ FUND_RANKS: r.FUND_RANKS, FUND_RANKS_META: r.FUND_RANKS_META }));
    console.log('wrote ' + args[oi + 1]);
  }
}
