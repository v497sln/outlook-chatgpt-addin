// Static data: formations, SBC templates, default fodder prices, sample club.
// The sample club is FICTIONAL (made-up names and clubs) and exists only so the
// app can be tried before importing your own club.

const S = (pos, x, y) => ({ pos, x, y });

export const FORMATIONS = {
  '4-3-3': [S('GK', 50, 90), S('LB', 14, 70), S('CB', 38, 74), S('CB', 62, 74), S('RB', 86, 70),
    S('CM', 28, 50), S('CM', 50, 54), S('CM', 72, 50), S('LW', 18, 24), S('ST', 50, 16), S('RW', 82, 24)],
  '4-4-2': [S('GK', 50, 90), S('LB', 14, 70), S('CB', 38, 74), S('CB', 62, 74), S('RB', 86, 70),
    S('LM', 14, 44), S('CM', 38, 50), S('CM', 62, 50), S('RM', 86, 44), S('ST', 36, 18), S('ST', 64, 18)],
  '4-2-3-1': [S('GK', 50, 90), S('LB', 14, 70), S('CB', 38, 74), S('CB', 62, 74), S('RB', 86, 70),
    S('CDM', 36, 56), S('CDM', 64, 56), S('LM', 16, 34), S('CAM', 50, 36), S('RM', 84, 34), S('ST', 50, 14)],
  '4-1-2-1-2': [S('GK', 50, 90), S('LB', 14, 70), S('CB', 38, 74), S('CB', 62, 74), S('RB', 86, 70),
    S('CDM', 50, 58), S('CM', 26, 46), S('CM', 74, 46), S('CAM', 50, 34), S('ST', 36, 16), S('ST', 64, 16)],
  '3-5-2': [S('GK', 50, 90), S('CB', 26, 74), S('CB', 50, 76), S('CB', 74, 74),
    S('LM', 10, 46), S('CDM', 38, 56), S('CDM', 62, 56), S('RM', 90, 46), S('CAM', 50, 36),
    S('ST', 36, 16), S('ST', 64, 16)],
  '5-2-1-2': [S('GK', 50, 90), S('LWB', 10, 62), S('CB', 30, 74), S('CB', 50, 76), S('CB', 70, 74), S('RWB', 90, 62),
    S('CM', 34, 50), S('CM', 66, 50), S('CAM', 50, 34), S('ST', 36, 16), S('ST', 64, 16)],
};

export function formation(name) {
  return { name, slots: FORMATIONS[name] || FORMATIONS['4-3-3'] };
}

// Default "any card" fodder prices by rating. Users should edit these to the
// current market — the app does NOT scrape prices or log into any account.
export const DEFAULT_PRICES = {
  75: 400, 76: 450, 77: 500, 78: 600, 79: 700, 80: 900, 81: 1200, 82: 1600, 83: 2400,
  84: 4200, 85: 7500, 86: 12000, 87: 17000, 88: 26000, 89: 38000, 90: 55000,
};

// Generic challenge templates (pattern-based, not copies of specific in-game SBCs).
export const TEMPLATES = [
  { id: 't-rating-83', name: 'Rating squad · 83', formation: '4-3-3', reqs: [{ t: 'rating', n: 83 }] },
  { id: 't-rating-85', name: 'Rating squad · 85', formation: '4-3-3', reqs: [{ t: 'rating', n: 85 }] },
  { id: 't-rating-87', name: 'Rating squad · 87', formation: '4-4-2', reqs: [{ t: 'rating', n: 87 }] },
  { id: 't-league-nation', name: 'League & nation mix', formation: '4-4-2', reqs: [
    { t: 'distinct', attr: 'league', op: 'min', n: 3 },
    { t: 'distinct', attr: 'nation', op: 'min', n: 3 },
    { t: 'same', attr: 'club', op: 'max', n: 3 },
    { t: 'count', attr: 'quality', value: 'gold', op: 'min', n: 11 },
    { t: 'chem', n: 18 },
  ] },
  { id: 't-hybrid', name: 'Hybrid leagues', formation: '4-2-3-1', reqs: [
    { t: 'distinct', attr: 'league', op: 'min', n: 5 },
    { t: 'same', attr: 'league', op: 'max', n: 3 },
    { t: 'rating', n: 80 },
    { t: 'chem', n: 20 },
  ] },
  { id: 't-first-xi', name: 'First XI', formation: '4-3-3', reqs: [
    { t: 'same', attr: 'nation', op: 'min', n: 4 },
    { t: 'same', attr: 'league', op: 'min', n: 5 },
    { t: 'count', attr: 'rare', op: 'min', n: 3 },
    { t: 'rating', n: 82 },
    { t: 'chem', n: 25 },
  ] },
  { id: 't-silver', name: 'Silver stars', formation: '4-4-2', reqs: [
    { t: 'ovr', min: 65, max: 74 },
    { t: 'count', attr: 'quality', value: 'silver', op: 'exact', n: 11 },
    { t: 'chem', n: 15 },
  ] },
  { id: 't-full-chem', name: 'Perfect links (33 chem)', formation: '4-3-3', reqs: [
    { t: 'chem', n: 33 },
    { t: 'rating', n: 78 },
  ] },
];

// --- sample club (fictional) ------------------------------------------------
const FIRST = ['Leo', 'Mateo', 'Jonah', 'Ilias', 'Tomas', 'Kofi', 'Rui', 'Anders', 'Luca', 'Yann', 'Diego', 'Emil',
  'Sami', 'Noah', 'Felix', 'Hugo', 'Marco', 'Ivan', 'Theo', 'Kenji', 'Omar', 'Pablo', 'Bastian', 'Nico', 'Adrien'];
const LAST = ['Varela', 'Brandt', 'Okafor', 'Lindqvist', 'Moreau', 'Castell', 'Novak', 'Ferro', 'Halden', 'Duarte',
  'Kessler', 'Mbeki', 'Rossetti', 'Sandoval', 'Aaltonen', 'Weir', 'Calder', 'Mendez', 'Petrov', 'Ishida',
  'Lambert', 'Quint', 'Barros', 'Hughes', 'Vidal', 'Strand', 'Ortega', 'Bianchi', 'Kovac', 'Dumont'];

const LEAGUES = {
  'Premier League': { nations: ['England', 'France', 'Brazil', 'Netherlands'], clubs: ['Northbridge FC', 'Harbour City', 'Ashford Rovers', 'Kingsmere United'] },
  'LALIGA': { nations: ['Spain', 'Argentina', 'Portugal', 'France'], clubs: ['Real Costa', 'Atlético Sierra', 'CD Marina'] },
  'Serie A': { nations: ['Italy', 'Argentina', 'Brazil', 'Croatia'], clubs: ['AC Lagoon', 'Sporting Verona', 'Torre FC'] },
  'Bundesliga': { nations: ['Germany', 'Netherlands', 'Croatia', 'England'], clubs: ['FC Rheinwald', 'Borussia Nordtal', 'SV Alpen'] },
  'Ligue 1': { nations: ['France', 'Portugal', 'Morocco', 'Brazil'], clubs: ['Olympique Azur', 'Stade Lumière', 'AS Côte'] },
  'Eredivisie': { nations: ['Netherlands', 'Morocco', 'Germany'], clubs: ['Polder FC', 'Haven Utd'] },
  'Liga Portugal': { nations: ['Portugal', 'Brazil', 'Spain'], clubs: ['SC Atlântico', 'FC Douro'] },
  'MLS': { nations: ['United States', 'Argentina', 'England'], clubs: ['Pacific Coast SC', 'Lakeshore FC'] },
};

const POS_GROUPS = [
  ['GK'], ['GK'], ['CB'], ['CB'], ['CB', 'CDM'], ['LB', 'LWB', 'LM'], ['RB', 'RWB', 'RM'],
  ['CDM', 'CM'], ['CM', 'CAM'], ['CM', 'CDM'], ['CAM', 'CM'], ['LW', 'LM'], ['RW', 'RM'],
  ['ST'], ['ST', 'CF'], ['LM', 'LW', 'LWB'], ['RM', 'RW', 'RWB'], ['CB', 'RB'],
];

function rng(seed) {
  let a = seed;
  return () => { a = (a * 1664525 + 1013904223) % 4294967296; return a / 4294967296; };
}

export function sampleClub() {
  const r = rng(27);
  const pick = arr => arr[Math.floor(r() * arr.length)];
  const leagues = Object.keys(LEAGUES);
  const out = [];
  for (let i = 0; i < 160; i++) {
    const league = pick(leagues);
    const L = LEAGUES[league];
    // rating distribution weighted towards 70–84
    const base = 62 + Math.floor(r() * 18) + Math.floor(r() * 8);
    const rating = Math.min(91, base + (r() < 0.12 ? 4 : 0));
    const rare = rating >= 75 ? r() < 0.7 : r() < 0.35;
    const untradeable = r() < 0.4;
    out.push({
      id: `s${i + 1}`,
      name: `${pick(FIRST)} ${pick(LAST)}`,
      rating,
      positions: pick(POS_GROUPS),
      nation: pick(L.nations),
      league,
      club: pick(L.clubs),
      rarity: rare ? 'rare' : 'common',
      rare,
      untradeable,
      count: r() < 0.08 ? 2 : 1,
      price: 0,
      protected: false,
    });
  }
  // a couple of special cards
  out.push({ id: 's-icon1', name: 'Henrik Vale', rating: 89, positions: ['CM', 'CDM'], nation: 'Netherlands', league: 'Icons', club: 'ICON', rarity: 'icon', rare: true, untradeable: true, count: 1, price: 0, protected: true });
  out.push({ id: 's-hero1', name: 'Dario Lentini', rating: 87, positions: ['ST'], nation: 'Italy', league: 'Serie A', club: 'Hero', rarity: 'hero', rare: true, untradeable: false, count: 1, price: 95000, protected: true });
  return out;
}
