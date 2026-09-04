/* ============================================================================
   test-ai-analyst.js — behavioural tests for the Centricity AI Analyst.
   Runs against the REAL embedded reference data, so it exercises the actual
   score coverage and the actual name joins rather than a fixture.
     node tools/test-ai-analyst.js
   ========================================================================== */
'use strict';
const { load } = require('./engine-host.js');
const { E, REF } = (function () { const o = load(); return { E: o.E, REF: o.REF }; })();

function row(o) {
  return Object.assign({
    client: 'Test Client', name: '', amc: '', product: 'Mutual Funds', asset: 'Equity',
    category: 'Flexi Cap', distributor: 'Centricity', folio: '1', units: 1000,
    invR: 10000000, mvR: 12000000, glR: 2000000, invCr: 1, cmvCr: 1.2, glCr: 0.2,
    xirrFrac: 0.12, bmxirrFrac: 0.10, days: 900
  }, o);
}

let fail = 0;
const ok = (c, m) => { console.log((c ? '  PASS  ' : '  FAIL  ') + m); if (!c) fail++; };

/* ---------- TEST 1: an unscored fund must not be penalised ---------- */
console.log('\n=== TEST 1 — unscored holding is assessed on the rest, never penalised ===');
{
  /* pick a real MF_PERF scheme that carries NO rank this cycle */
  const perf = Object.keys(REF.MF_PERF);
  const unscored = perf.filter(k => !REF.FUND_RANKS[k] || REF.FUND_RANKS[k].rank == null);
  const scored = perf.filter(k => REF.FUND_RANKS[k] && REF.FUND_RANKS[k].rank != null);
  console.log('  MF_PERF ' + perf.length + ' schemes: ' + scored.length + ' scored, ' + unscored.length + ' unscored');
  const target = unscored.filter(k => REF.MF_RISK[k])[0];
  console.log('  using unscored scheme: ' + target);

  const rows = [
    row({ name: target, amc: target.split(' ')[0], cmvCr: 1.0, invCr: 0.9 }),
    row({ name: scored[0], amc: scored[0].split(' ')[0], cmvCr: 1.0, invCr: 0.9 }),
    row({ name: scored[1], amc: scored[1].split(' ')[0], cmvCr: 1.0, invCr: 0.9 }),
    row({ name: scored[2], amc: scored[2].split(' ')[0], cmvCr: 1.0, invCr: 0.9 })
  ];
  const m = E.computeReview(rows, REF, { familyName: 'T1', asOnDate: '31 Jul 2026' });
  const P = E.aiaCompute(m, REF, {});
  const rec = P.recs.filter(r => r.name === target)[0];
  ok(!!rec, 'the unscored holding produced a recommendation rather than vanishing');
  if (rec) {
    const sc = rec.signals.filter(s => s.id === 'score')[0];
    ok(sc.verdict === 'na', 'its Fund Score signal is "na", not "red" (got ' + sc.verdict + ')');
    ok(rec.assessed < rec.of, 'it is assessed on fewer signals: ' + rec.assessed + ' of ' + rec.of);
    /* the composite must be the mean over AVAILABLE signals: recompute by hand */
    const W = P.weights;
    let num = 0, den = 0;
    rec.signals.forEach(s => { if (s.verdict === 'na') return;
      const v = s.verdict === 'green' ? 100 : s.verdict === 'amber' ? 50 : 0;
      num += v * W[s.id]; den += W[s.id]; });
    const expect = Math.round((num / den) * 100) / 100;
    ok(Math.abs(expect - rec.composite) < 0.02,
       'composite ' + rec.composite + ' equals the renormalised mean over available signals (' + expect + ')');
    ok(!(rec.action === 'Exit'), 'it is not exited merely for lacking a score (action ' + rec.action + ')');
    ok(rec.conviction === rec.signals.filter(s => s.verdict === 'red' &&
        ['score','perf','risk','overlap','mgr'].indexOf(s.id) >= 0).length,
       'conviction counts fault signals only');
  }
}

/* ---------- TEST 2: an AMC over the house cap ---------- */
console.log('\n=== TEST 2 — an AMC at ~22% is a breach on both tabs, at the 20% house cap ===');
{
  const perf = Object.keys(REF.MF_PERF);
  const hdfc = perf.filter(k => /^HDFC /.test(k)).slice(0, 2);
  const others = perf.filter(k => !/^HDFC /.test(k)).slice(0, 8);
  /* 2.2 of 10.0 Cr into one AMC = 22% */
  const rows = [
    row({ name: hdfc[0], amc: 'HDFC', cmvCr: 1.1, invCr: 1.0 }),
    row({ name: hdfc[1], amc: 'HDFC', cmvCr: 1.1, invCr: 1.0 })
  ].concat(others.map(k => row({ name: k, amc: k.split(' ')[0], cmvCr: 0.975, invCr: 0.9 })));
  const m = E.computeReview(rows, REF, { familyName: 'T2', asOnDate: '31 Jul 2026' });
  const hd = (m.amc.rows || []).filter(r => r.amc === 'HDFC')[0];
  console.log('  total CMV ' + m.totalCMV.toFixed(4) + '   HDFC ' + (hd ? hd.pct.toFixed(2) : '?') + '%');
  ok(!!hd && hd.pct > 20, 'HDFC is above 20% (' + (hd ? hd.pct.toFixed(2) : '?') + '%)');
  ok(!!hd && hd.breach === true, 'the AMC Allocation tab marks it a breach');
  ok((m.amc.breaches || []).indexOf('HDFC') >= 0 ||
     (m.amc.breaches || []).some(b => (b && (b.amc === 'HDFC' || b === 'HDFC'))),
     'it appears in model.amc.breaches');

  const P = E.aiaCompute(m, REF, {});
  ok(P.amcCap === 20, 'the AI Analyst applies a 20% cap (got ' + P.amcCap + ')');
  /* the analyst must see the same concentration */
  const hdRecs = P.recs.filter(r => r.amc === 'HDFC');
  const sawIt = hdRecs.some(r => r.signals.some(s => s.id === 'impact' &&
      s.evidence && s.evidence.amcBefore != null && s.evidence.amcBefore > 20));
  ok(sawIt, 'the analyst reports the same >20% concentration on the HDFC holdings');
  /* and no destination may push any AMC past the cap */
  const after = {}; (m.amc.rows || []).forEach(r => after[r.amc] = r.cmv);
  P.recs.forEach(r => { if (r.release > 0) after[r.amc] = (after[r.amc] || 0) - r.release; });
  P.recs.forEach(r => (r.destinations || []).forEach(d => {
    const a = d.amc || ''; after[a] = (after[a] || 0) + d.amount; }));
  const over = Object.keys(after).filter(a => a && (after[a] / m.totalCMV * 100) > 20.01);
  ok(over.length === 0, 'no destination pushes an AMC past the cap' +
     (over.length ? (' — ' + over.map(a => a + ' ' + (after[a] / m.totalCMV * 100).toFixed(2) + '%').join(', ')) : ''));
}

/* ---------- TEST 3: determinism and conservation on both portfolios ---------- */
console.log('\n=== TEST 3 — determinism and conservation ===');
{
  const perf = Object.keys(REF.MF_PERF).slice(0, 12);
  const rows = perf.map((k, i) => row({ name: k, amc: k.split(' ')[0], cmvCr: 0.5 + i * 0.1, invCr: 0.4 + i * 0.1 }));
  const m = E.computeReview(rows, REF, { familyName: 'T3', asOnDate: '31 Jul 2026' });
  const a = E.aiaCompute(m, REF, {}), b = E.aiaCompute(m, REF, {});
  ok(JSON.stringify(a) === JSON.stringify(b), 'the same portfolio produces a byte-identical plan');
  /* released = invested + unfunded.  is a legitimate outcome: the
     10% fund cap and the 20% AMC cap can both refuse the tail of a release, and
     it is reported rather than forced past a limit. */
  let bad = [];
  a.recs.forEach(r => {
    if (!r.destinations.length) return;
    const s = r.destinations.reduce((x, d) => x + d.amount, 0);
    if (Math.abs(s + (r.unfunded || 0) - r.release) > 0.0051)
      bad.push(r.name + ' dest ' + s.toFixed(4) + ' + unfunded ' + (r.unfunded || 0).toFixed(4) + ' vs release ' + r.release.toFixed(4));
  });
  ok(bad.length === 0, 'destinations plus unfunded equal the release on every source' + (bad.length ? (': ' + bad.join('; ')) : ''));
  /* the caps themselves */
  const capViol = [];
  const fin = {};
  a.recs.forEach(r => r.destinations.forEach(d => {
    const h = m.holdings.filter(x => x.name === d.name || x.perfKey === d.name)[0];
    if (!(d.name in fin)) fin[d.name] = h ? h.cmvCr : 0;
    fin[d.name] += d.amount; }));
  Object.keys(fin).forEach(k => { const w = fin[k] / m.totalCMV * 100;
    if (w > 10 + 1e-9) capViol.push('fund ' + k + ' at ' + w.toFixed(3) + '%'); });
  ok(capViol.length === 0, 'no destination exceeds the 10% per-fund cap' + (capViol.length ? (': ' + capViol.join('; ')) : ''));
  const exitsNoDest = a.recs.filter(r => r.action === 'Exit' && !r.destinations.length);
  ok(exitsNoDest.length === 0, 'no exit without a funded destination');
  const thin = a.recs.filter(r => r.action === 'Exit' &&
    r.signals.filter(s => s.verdict === 'red' && ['score','perf','risk','overlap','mgr'].indexOf(s.id) >= 0).length < 2);
  ok(thin.length === 0, 'no exit on fewer than two independent faults');
}

console.log('\n' + (fail ? (fail + ' FAILURE(S)') : 'all acceptance checks passed'));
process.exit(fail ? 1 : 0);
