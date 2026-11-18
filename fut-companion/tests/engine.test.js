import test from 'node:test';
import assert from 'node:assert/strict';
import { teamRating, chemistry, evaluate, solve, roundToMarket, afterTax } from '../js/engine.js';
import { formation, sampleClub, DEFAULT_PRICES, TEMPLATES } from '../js/data.js';

test('team rating: flat squad', () => {
  assert.equal(teamRating(Array(11).fill(84)), 84);
});

test('team rating: excess above average counts', () => {
  // 10x83 + 1x93: avg 83.909, excess 9.09 -> (923+9.09)/11 = 84.73 -> 84
  assert.equal(teamRating([...Array(10).fill(83), 93]), 84);
});

test('team rating never rounds up', () => {
  // 10x84 + 1x85 -> raw ~84.17 -> 84
  assert.equal(teamRating([...Array(10).fill(84), 85]), 84);
});

const mk = (o) => ({ rating: 80, positions: ['CM'], nation: 'A', league: 'L', club: 'C', rarity: 'rare', ...o });

test('chemistry: full same-club squad in position = 33', () => {
  const f = formation('4-3-3');
  const squad = f.slots.map((s, i) => mk({ id: i, positions: [s.pos] }));
  assert.equal(chemistry(squad, f).total, 33);
});

test('chemistry: out of position player gets 0', () => {
  const f = formation('4-3-3');
  const squad = f.slots.map((s, i) => mk({ id: i, positions: [s.pos] }));
  squad[0] = mk({ id: 'x', positions: ['ST'] }); // ST in GK slot
  const c = chemistry(squad, f);
  assert.equal(c.per[0], 0);
});

test('requirements: same/distinct', () => {
  const f = formation('4-3-3');
  const squad = f.slots.map((s, i) => mk({ id: i, positions: [s.pos], league: i < 4 ? 'L1' : 'L2' }));
  const ev = evaluate(squad, f, [
    { t: 'distinct', attr: 'league', op: 'min', n: 2 },
    { t: 'same', attr: 'league', op: 'max', n: 5 },
  ]);
  assert.equal(ev.results[1].ok, true);
  assert.equal(ev.results[2].ok, false); // 7 from L2 > 5
});

test('market price rounding', () => {
  assert.equal(roundToMarket(1234), 1200);
  assert.equal(roundToMarket(12380), 12500);
  assert.equal(roundToMarket(10), 150);
  assert.equal(afterTax(10000), 9500);
});

test('solver solves every template with the sample club + market fodder', () => {
  const club = sampleClub();
  for (const t of TEMPLATES) {
    const res = solve({ club, formation: formation(t.formation), reqs: t.reqs, priceTable: DEFAULT_PRICES, timeMs: 1500, seed: 7 });
    assert.ok(res.ok, `${t.name} failed: ${res.eval?.results.filter(r => !r.ok).map(r => `${r.label} (${r.actual})`).join('; ')}`);
    assert.equal(new Set(res.squad.map(p => p.uid)).size, 11);
    assert.ok(res.squad.every(p => !p.protected));
  }
});

test('solver respects protected and impossible requests report a reason', () => {
  const res = solve({ club: [], formation: formation('4-3-3'), reqs: [{ t: 'rating', n: 80 }], useMarket: false });
  assert.equal(res.ok, false);
  assert.match(res.reason, /need 11/);
});
