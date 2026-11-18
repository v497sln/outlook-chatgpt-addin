// Local persistence. Everything stays on the device — no account, no server.
import { DEFAULT_PRICES } from './data.js';

const KEY = 'squadsage:v1';

export function defaultState() {
  return {
    club: [],
    trades: [],
    savedSbcs: [],
    history: [],
    squad: { formation: '4-3-3', ids: Array(11).fill(null) },
    prices: { ...DEFAULT_PRICES },
    settings: { useMarket: true, thorough: false },
  };
}

export function load() {
  try {
    const raw = localStorage.getItem(KEY);
    if (!raw) return defaultState();
    return { ...defaultState(), ...JSON.parse(raw) };
  } catch {
    return defaultState();
  }
}

export function save(state) {
  try { localStorage.setItem(KEY, JSON.stringify(state)); return true; } catch { return false; }
}

// --- CSV --------------------------------------------------------------------
export const CSV_HEADER = 'name,rating,positions,nation,league,club,rarity,untradeable,count,price,protected';

function splitCsvLine(line) {
  const out = []; let cur = ''; let q = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (q) {
      if (c === '"' && line[i + 1] === '"') { cur += '"'; i++; }
      else if (c === '"') q = false;
      else cur += c;
    } else if (c === '"') q = true;
    else if (c === ',') { out.push(cur); cur = ''; }
    else cur += c;
  }
  out.push(cur);
  return out.map(s => s.trim());
}

const truthy = v => /^(1|y|yes|true|x)$/i.test(String(v || '').trim());

// Returns { players, errors } — bad rows are reported, not silently dropped.
export function parseCsv(text) {
  const lines = text.split(/\r?\n/).filter(l => l.trim());
  if (!lines.length) return { players: [], errors: ['File is empty'] };
  let header = splitCsvLine(lines[0]).map(h => h.toLowerCase());
  let start = 1;
  if (!header.includes('name') || !header.includes('rating')) {
    header = CSV_HEADER.split(',');
    start = 0;
  }
  const idx = k => header.indexOf(k);
  const players = []; const errors = [];
  for (let i = start; i < lines.length; i++) {
    const c = splitCsvLine(lines[i]);
    const get = k => (idx(k) >= 0 ? c[idx(k)] : '');
    const rating = parseInt(get('rating'), 10);
    const name = get('name');
    if (!name || !(rating >= 1 && rating <= 99)) {
      errors.push(`Line ${i + 1}: needs a name and a rating 1–99`);
      continue;
    }
    const rarity = (get('rarity') || 'common').toLowerCase();
    players.push({
      id: `p${Date.now().toString(36)}${i}${Math.random().toString(36).slice(2, 6)}`,
      name,
      rating,
      positions: (get('positions') || '').split(/[|/ ;]+/).map(s => s.toUpperCase()).filter(Boolean),
      nation: get('nation'),
      league: get('league'),
      club: get('club'),
      rarity,
      rare: rarity !== 'common',
      untradeable: truthy(get('untradeable')),
      count: Math.max(1, parseInt(get('count'), 10) || 1),
      price: Math.max(0, parseInt(get('price'), 10) || 0),
      protected: truthy(get('protected')),
    });
  }
  return { players, errors };
}

const esc = v => {
  const s = String(v ?? '');
  return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
};

export function toCsv(club) {
  return [CSV_HEADER, ...club.map(p => [
    p.name, p.rating, (p.positions || []).join('|'), p.nation, p.league, p.club, p.rarity,
    p.untradeable ? 'yes' : 'no', p.count || 1, p.price || 0, p.protected ? 'yes' : 'no',
  ].map(esc).join(','))].join('\n');
}

export function download(filename, text, type = 'text/plain') {
  const blob = new Blob([text], { type });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
}
