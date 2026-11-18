// Pure game-logic engine: team rating, chemistry, SBC requirement checks,
// and a local SBC solver. No DOM access so it can be unit-tested in Node.

export const ATTRS = ['nation', 'league', 'club'];

export function quality(rating) {
  return rating >= 75 ? 'gold' : rating >= 65 ? 'silver' : 'bronze';
}

// ---------------------------------------------------------------------------
// Team rating
// FUT divides by 11, then adds "excess" for every player above the average.
// We round DOWN (never up) so the app never claims a squad passes when the
// game would reject it — a top complaint about the original app was failed
// SBC submissions.
export function teamRatingRaw(ratings) {
  const total = ratings.reduce((a, b) => a + (b || 0), 0);
  const avg = total / 11;
  let excess = 0;
  for (const r of ratings) if (r && r > avg) excess += r - avg;
  return (total + excess) / 11;
}

export function teamRating(ratings) {
  return Math.floor(Math.round(teamRatingRaw(ratings) * 1000) / 1000);
}

// ---------------------------------------------------------------------------
// Chemistry (current system: 0–3 per player, 33 max)
const CLUB_STEPS = [2, 5, 8];
const NATION_STEPS = [2, 5, 8];
const LEAGUE_STEPS = [3, 5, 8];

function steps(count, thresholds) {
  let p = 0;
  for (const t of thresholds) if (count >= t) p++;
  return p;
}

export function inPosition(player, slotPos) {
  return !!player && Array.isArray(player.positions) && player.positions.includes(slotPos);
}

// squad: array aligned with formation.slots, entries may be null.
export function chemistry(squad, formation) {
  const club = new Map(), nation = new Map(), league = new Map();
  let iconLeagueBonus = 0;
  const add = (m, k, n) => { if (k) m.set(k, (m.get(k) || 0) + n); };

  squad.forEach((p, i) => {
    if (!p || p.generic || !inPosition(p, formation.slots[i].pos)) return;
    const r = p.rarity;
    if (r === 'icon') {
      add(nation, p.nation, 2);
      iconLeagueBonus += 1;
    } else if (r === 'hero') {
      add(nation, p.nation, 1);
      add(league, p.league, 2);
      add(club, p.club, 1);
    } else {
      add(nation, p.nation, 1);
      add(league, p.league, 1);
      add(club, p.club, 1);
    }
  });

  const per = squad.map((p, i) => {
    if (!p || p.generic) return 0;
    if (!inPosition(p, formation.slots[i].pos)) return 0;
    if (p.rarity === 'icon' || p.rarity === 'hero') return 3;
    const lg = (league.get(p.league) || 0) + iconLeagueBonus;
    const pts = steps(club.get(p.club) || 0, CLUB_STEPS)
      + steps(nation.get(p.nation) || 0, NATION_STEPS)
      + steps(lg, LEAGUE_STEPS);
    return Math.min(3, pts);
  });
  return { total: per.reduce((a, b) => a + b, 0), per };
}

// ---------------------------------------------------------------------------
// Requirements
// Types:
//   {t:'rating', n}                      min team rating
//   {t:'chem', n}                        min total chemistry
//   {t:'playerChem', n}                  every player has at least n chem
//   {t:'count', attr, value, op, n}      attr in nation|league|club|rarity|quality|rare
//   {t:'same', attr, op, n}              min/max players from the SAME attr value
//   {t:'distinct', attr, op, n}          min/max DIFFERENT attr values
//   {t:'ovr', min, max}                  every player's rating within bounds
const OPS = { min: 'Min', max: 'Max', exact: 'Exactly' };
const ATTR_LABEL = { nation: 'nation', league: 'league', club: 'club', rarity: 'rarity', quality: 'quality', rare: 'rare' };

export function describeReq(r) {
  switch (r.t) {
    case 'rating': return `Team rating: min ${r.n}`;
    case 'chem': return `Squad chemistry: min ${r.n}`;
    case 'playerChem': return `Each player chemistry: min ${r.n}`;
    case 'count':
      if (r.attr === 'rare') return `${OPS[r.op]} ${r.n} rare player${r.n === 1 ? '' : 's'}`;
      return `${OPS[r.op]} ${r.n} player${r.n === 1 ? '' : 's'} with ${ATTR_LABEL[r.attr]}: ${r.value}`;
    case 'same': return `${OPS[r.op]} ${r.n} players from the same ${ATTR_LABEL[r.attr]}`;
    case 'distinct': return `${OPS[r.op]} ${r.n} different ${ATTR_LABEL[r.attr]}s`;
    case 'ovr': {
      const parts = [];
      if (r.min != null) parts.push(`min ${r.min}`);
      if (r.max != null) parts.push(`max ${r.max}`);
      return `Player rating: ${parts.join(', ')}`;
    }
    default: return 'Unknown requirement';
  }
}

function attrValue(p, attr) {
  if (!p) return null;
  if (attr === 'quality') return quality(p.rating);
  if (attr === 'rare') return p.rare ? true : null;
  if (p.generic) return null; // "any card" fodder has no nation/league/club
  return p[attr] ?? null;
}

function counts(squad, attr) {
  const m = new Map();
  for (const p of squad) {
    const v = attrValue(p, attr);
    if (v == null) continue;
    m.set(v, (m.get(v) || 0) + 1);
  }
  return m;
}

function cmp(op, actual, n) {
  if (op === 'min') return { ok: actual >= n, deficit: Math.max(0, n - actual) };
  if (op === 'max') return { ok: actual <= n, deficit: Math.max(0, actual - n) };
  return { ok: actual === n, deficit: Math.abs(actual - n) };
}

// Evaluates a full squad. Returns per-requirement results with a numeric
// deficit (0 = satisfied) that the solver uses as a gradient.
export function evaluate(squad, formation, reqs) {
  const ratings = squad.map(p => (p ? p.rating : 0));
  const rawRating = teamRatingRaw(ratings);
  const rating = teamRating(ratings);
  const chem = chemistry(squad, formation);
  const filled = squad.filter(Boolean).length;

  const results = reqs.map(r => {
    let actual, res;
    switch (r.t) {
      case 'rating':
        actual = rating;
        res = { ok: rating >= r.n, deficit: rating >= r.n ? 0 : Math.max(0.05, r.n - rawRating) };
        break;
      case 'chem':
        actual = chem.total;
        res = cmp('min', chem.total, r.n);
        break;
      case 'playerChem': {
        const short = chem.per.reduce((a, c, i) => a + (squad[i] ? Math.max(0, r.n - c) : 0), 0);
        actual = Math.min(...chem.per.filter((_, i) => squad[i]), 3);
        res = { ok: short === 0, deficit: short };
        break;
      }
      case 'count': {
        const c = r.attr === 'rare'
          ? squad.filter(p => p && p.rare).length
          : (counts(squad, r.attr).get(r.value) || 0);
        actual = c;
        res = cmp(r.op, c, r.n);
        break;
      }
      case 'same': {
        const m = counts(squad, r.attr);
        const top = m.size ? Math.max(...m.values()) : 0;
        actual = top;
        if (r.op === 'max') {
          let over = 0;
          for (const v of m.values()) over += Math.max(0, v - r.n);
          res = { ok: over === 0, deficit: over };
        } else res = cmp(r.op, top, r.n);
        break;
      }
      case 'distinct': {
        const d = counts(squad, r.attr).size;
        actual = d;
        res = cmp(r.op, d, r.n);
        break;
      }
      case 'ovr': {
        let bad = 0;
        for (const p of squad) {
          if (!p) continue;
          if (r.min != null && p.rating < r.min) bad += r.min - p.rating;
          if (r.max != null && p.rating > r.max) bad += p.rating - r.max;
        }
        actual = bad ? 'out of range' : 'ok';
        res = { ok: bad === 0, deficit: bad };
        break;
      }
      default:
        res = { ok: true, deficit: 0 };
    }
    return { req: r, label: describeReq(r), actual, ...res };
  });

  const missing = 11 - filled;
  results.unshift({
    req: { t: 'size' }, label: 'Squad has 11 players', actual: filled,
    ok: missing === 0, deficit: missing * 5,
  });

  return {
    rating, rawRating, chem,
    results,
    ok: results.every(r => r.ok),
    violation: results.reduce((a, r) => a + r.deficit, 0),
  };
}

// ---------------------------------------------------------------------------
// Market price helpers

// Transfer-market price steps (prices must land on these increments).
export function priceStep(price) {
  if (price < 1000) return 50;
  if (price < 10000) return 100;
  if (price < 50000) return 250;
  if (price < 100000) return 500;
  return 1000;
}

export function roundToMarket(price, dir = 'nearest') {
  price = Math.max(150, price);
  const s = priceStep(price);
  const f = dir === 'down' ? Math.floor : dir === 'up' ? Math.ceil : Math.round;
  return Math.max(150, f(price / s) * s);
}

export const EA_TAX = 0.05;
export const afterTax = p => Math.floor(p * (1 - EA_TAX));

// ---------------------------------------------------------------------------
// Solver

function mulberry32(seed) {
  let a = seed >>> 0;
  return function () {
    a |= 0; a = (a + 0x6D2B79F5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

// What a card "costs" to put into an SBC.
//  - Protected cards are never used.
//  - Untradeables are free (they can't be sold anyway) — used first.
//  - Duplicates are discounted so they get cleared first.
//  - Tradeable club cards cost what you'd get selling them (after 5% tax).
//  - Market buys cost the full price.
export function cardCost(p, priceTable) {
  const est = p.price || priceTable[p.rating] || 0;
  if (p.generic) return est;
  if (p.untradeable) return p.rating * 0.01 + (p.duplicate ? 0 : 0.5);
  const v = afterTax(est);
  return p.duplicate ? v * 0.6 : v;
}

export function buildPool(club, opts = {}) {
  const { priceTable = {}, useMarket = true, ovrReq = null, excludeIds = [] } = opts;
  const ex = new Set(excludeIds);
  const pool = [];
  for (const p of club) {
    if (p.protected || ex.has(p.id)) continue;
    const copies = Math.max(1, p.count || 1);
    for (let k = 0; k < copies; k++) {
      pool.push({ ...p, uid: `${p.id}#${k}`, duplicate: p.duplicate || k > 0, source: 'club' });
    }
  }
  if (useMarket) {
    for (const [rStr, price] of Object.entries(priceTable)) {
      const r = +rStr;
      if (!price) continue;
      for (let k = 0; k < 11; k++) {
        pool.push({
          id: `any-${r}`, uid: `any-${r}#${k}`, name: `Any ${r}-rated card`, rating: r,
          positions: [], generic: true, rare: false, price, source: 'market',
        });
      }
    }
  }
  const inRange = p => !ovrReq
    || ((ovrReq.min == null || p.rating >= ovrReq.min) && (ovrReq.max == null || p.rating <= ovrReq.max));
  const filtered = pool.filter(inRange);
  for (const p of filtered) p.cost = cardCost(p, priceTable);
  return filtered;
}

export function solve({ club, formation, reqs, priceTable = {}, useMarket = true, locked = [],
  timeMs = 1200, seed = 12345, excludeIds = [] }) {
  const rng = mulberry32(seed);
  const ovrReq = reqs.find(r => r.t === 'ovr') || null;
  const pool = buildPool(club, { priceTable, useMarket, ovrReq, excludeIds });
  const N = formation.slots.length;
  if (pool.length < N) {
    return { ok: false, reason: `Only ${pool.length} eligible cards available — need ${N}. Add players to your club or enable market fodder.` };
  }

  // Locked slots: index -> pool uid that must stay.
  const lockedIdx = new Map();
  for (const { slot, id } of locked) {
    const i = pool.findIndex(p => p.id === id && ![...lockedIdx.values()].includes(pool.indexOf(p)));
    if (i >= 0) lockedIdx.set(slot, i);
  }

  // Indexes for "neighbour" moves that build links (chem / same-league etc.).
  const byAttr = {};
  for (const a of ATTRS) {
    byAttr[a] = new Map();
    pool.forEach((p, i) => {
      if (p.generic || !p[a]) return;
      if (!byAttr[a].has(p[a])) byAttr[a].set(p[a], []);
      byAttr[a].get(p[a]).push(i);
    });
  }
  const byPos = new Map();
  pool.forEach((p, i) => (p.positions || []).forEach(pos => {
    if (!byPos.has(pos)) byPos.set(pos, []);
    byPos.get(pos).push(i);
  }));
  const byRating = new Map();
  pool.forEach((p, i) => {
    if (!byRating.has(p.rating)) byRating.set(p.rating, []);
    byRating.get(p.rating).push(i);
  });
  const byCost = pool.map((_, i) => i).sort((a, b) => pool[a].cost - pool[b].cost);

  const costScale = Math.max(100, median(pool.map(p => p.cost).filter(c => c > 1)) || 1000);
  const W = 400; // requirement violations dominate cost

  const energy = (assign) => {
    const squad = assign.map(i => pool[i]);
    const ev = evaluate(squad, formation, reqs);
    const cost = squad.reduce((a, p) => a + p.cost, 0);
    return { e: cost / costScale + ev.violation * W, cost, ev };
  };

  const randomInit = () => {
    const used = new Set(lockedIdx.values());
    const a = new Array(N);
    for (const [s, i] of lockedIdx) a[s] = i;
    for (let s = 0; s < N; s++) {
      if (a[s] != null) continue;
      // Prefer an in-position cheap card, else any cheap card.
      const cands = (byPos.get(formation.slots[s].pos) || []).filter(i => !used.has(i));
      let pick;
      if (cands.length && rng() < 0.7) pick = cands[Math.floor(rng() * Math.min(cands.length, 40))];
      if (pick == null) {
        const free = byCost.filter(i => !used.has(i));
        pick = free[Math.floor(rng() * Math.min(free.length, 60))];
      }
      a[s] = pick; used.add(pick);
    }
    return a;
  };

  const freeSlots = [...Array(N).keys()].filter(s => !lockedIdx.has(s));
  const deadline = Date.now() + timeMs;
  let best = null;
  let restarts = 0;

  while (Date.now() < deadline && freeSlots.length) {
    restarts++;
    let cur = randomInit();
    let used = new Set(cur);
    let curE = energy(cur);
    if (!best || curE.e < best.e) best = { ...curE, assign: cur.slice() };
    let T = 3;
    const iters = 6000;
    for (let it = 0; it < iters; it++) {
      if ((it & 255) === 0 && Date.now() > deadline) break;
      const next = cur.slice();
      const r = rng();
      const s = freeSlots[Math.floor(rng() * freeSlots.length)];
      if (r < 0.25 && freeSlots.length > 1) {
        // swap two slots (fixes positions -> chem)
        const s2 = freeSlots[Math.floor(rng() * freeSlots.length)];
        [next[s], next[s2]] = [next[s2], next[s]];
      } else {
        let cand;
        if (r < 0.55) {
          // link move: share nation/league/club with another squad member
          const other = pool[cur[Math.floor(rng() * N)]];
          const a = ATTRS[Math.floor(rng() * 3)];
          const list = other && !other.generic ? byAttr[a].get(other[a]) : null;
          if (list && list.length) cand = list[Math.floor(rng() * list.length)];
        } else if (r < 0.7) {
          const list = byPos.get(formation.slots[s].pos);
          if (list && list.length) cand = list[Math.floor(rng() * list.length)];
        } else if (r < 0.8) {
          cand = byCost[Math.floor(rng() * Math.min(byCost.length, 80))];
        } else if (r < 0.92) {
          // rating nudge: similar rating, one up or one down
          const target = pool[cur[s]].rating + (rng() < 0.6 ? 1 : -1) * (1 + Math.floor(rng() * 2));
          const list = byRating.get(target);
          if (list && list.length) cand = list[Math.floor(rng() * list.length)];
        }
        if (cand == null) cand = Math.floor(rng() * pool.length);
        if (used.has(cand)) continue;
        next[s] = cand;
      }
      const nE = energy(next);
      if (nE.e <= curE.e || rng() < Math.exp((curE.e - nE.e) / T)) {
        cur = next; curE = nE; used = new Set(cur);
        if (curE.e < best.e) best = { ...curE, assign: cur.slice() };
      }
      T *= 0.9993;
    }
  }
  if (!best) best = { ...energy(randomInit()), assign: null };
  if (!best.assign) best.assign = randomInit();

  // Polish: greedy single-slot replacements with cheaper cards that keep validity.
  polish(best, pool, freeSlots, energy);

  const squad = best.assign.map(i => pool[i]);
  const ev = evaluate(squad, formation, reqs);
  return {
    ok: ev.ok,
    squad,
    eval: ev,
    cost: squad.reduce((a, p) => a + p.cost, 0),
    buyCost: squad.filter(p => p.source === 'market').reduce((a, p) => a + p.price, 0),
    restarts,
    reason: ev.ok ? null : 'No valid squad found with the cards available. See which requirements still fail below.',
  };
}

function polish(best, pool, freeSlots, energy) {
  let improved = true, rounds = 0;
  while (improved && rounds++ < 4) {
    improved = false;
    const used = new Set(best.assign);
    for (const s of freeSlots) {
      for (let i = 0; i < pool.length; i++) {
        if (used.has(i) || pool[i].cost >= pool[best.assign[s]].cost) continue;
        const next = best.assign.slice(); next[s] = i;
        const e = energy(next);
        if (e.e < best.e - 1e-9 && e.ev.violation <= best.ev.violation) {
          used.delete(best.assign[s]); used.add(i);
          Object.assign(best, e, { assign: next });
          improved = true;
        }
      }
    }
  }
}

function median(arr) {
  if (!arr.length) return 0;
  const s = arr.slice().sort((a, b) => a - b);
  return s[Math.floor(s.length / 2)];
}
