// SquadSage overlay — runs INSIDE the EA FC Web App page (injected by the
// Android shell). EA's own app stays fully usable; SquadSage adds a floating
// button that opens its panel (SBC solver, club locks, trade stats, limits).
//
// The panel lives in a Shadow DOM, so EA's CSS can't break it and it can't
// break EA's page. EA's internals are undocumented: every call is wrapped and
// failures are shown in plain words and in the log.
//
// Bundled with engine.js by build-www.mjs; `__m.engine` is available here.

(function () {
  if (window.__ssOverlay) return;
  window.__ssOverlay = true;

  const E = __m.engine;
  const KEY = 'squadsage:overlay:v1';

  // ------------------------------------------------------------ storage
  const S = (() => {
    const base = { v: 2, bgPrices: false, bgPerHour: 40, locked: {}, subs: [], searches: [], prices: {}, limits: { hour: 90, day: 300, search: 150 }, club: null, clubAt: 0, onlyUT: false, undercut: true, floorBE: true, log: [] };
    try { return Object.assign(base, JSON.parse(localStorage.getItem(KEY) || '{}')); } catch { return base; }
  })();
  if ((S.v || 1) < 2) { S.bgPrices = false; S.v = 2; }
  S.limits.search = S.limits.search || 150; S.prices = S.prices || {}; S.searches = S.searches || [];
  const save = () => { try { localStorage.setItem(KEY, JSON.stringify(S)); } catch { /* storage full */ } };
  function log(msg, kind = 'info') {
    S.log.unshift({ at: Date.now(), msg, kind });
    S.log = S.log.slice(0, 80);
    save();
    if (ui.view === 'settings') render();
  }

  // ------------------------------------------------------------ EA helpers
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const obs = o => new Promise((resolve, reject) => {
    if (!o || !o.observe) return reject(new Error('EA service not available'));
    o.observe(window, (x, r) => {
      try { x.unobserve(window); } catch { /* ignore */ }
      if (r && r.success) resolve(r);
      else reject(new Error(`EA request failed (${(r && (r.status || (r.error && r.error.code))) || 'unknown'})`));
    });
  });
  const body = r => (r && (r.response || r.data)) || {};
  const loc = (keys, fb) => {
    try { for (const k of keys) { const t = services.Localization.localize(k); if (t && t !== k && !/^\*/.test(t) && !/global\.|search\./.test(t)) return t; } } catch { /* ignore */ }
    return fb;
  };
  const nation = id => loc([`search.nationName.nation${id}`, `global.nationName.nation${id}`], `Nation #${id}`);
  const league = id => loc([`global.leagueFull.2027.league${id}`, `global.leagueFull.2026.league${id}`, `global.leagueFull.${id}`, `global.leagueabbr15.2027.league${id}`, `global.leagueabbr15.${id}`], `League #${id}`);
  const club = id => loc([`global.teamFull.2027.team${id}`, `global.teamFull.2026.team${id}`, `global.teamFull.${id}`, `global.teamabbr15.2027.team${id}`, `global.teamabbr15.${id}`], `Club #${id}`);
  const POS = ['GK', 'SW', 'RWB', 'RB', 'CB', 'CB', 'CB', 'LB', 'LWB', 'CDM', 'CDM', 'CDM', 'RM', 'CM', 'CM', 'CM', 'LM', 'CAM', 'CAM', 'CAM', 'CF', 'CF', 'CF', 'RW', 'ST', 'ST', 'ST', 'LW'];
  const pos = p => (typeof p === 'number' ? POS[p] || null : typeof p === 'string' ? p.toUpperCase() : null);
  const nameOf = i => { try { const s = i.getStaticData ? i.getStaticData() : i._staticData; if (s && s.name) return s.name; } catch { /* ignore */ } return `Player ${i.definitionId}`; };
  const items = {}; // eaId -> live EA item (needed to submit)

  function isUntradeable(i) {
    try {
      if (typeof i.isUntradeable === 'function') return !!i.isUntradeable();
      if (typeof i.untradeable === 'boolean') return i.untradeable;
      if (i.untradeableCount > 0 || i._untradeable === true) return true;
      if (typeof i.isTradeable === 'function') return !i.isTradeable();
    } catch { /* ignore */ }
    return !!i.untradeable;
  }

  function mapItem(i) {
    let ps = [pos(i.preferredPosition), ...(i.possiblePositions || []).map(pos)].filter(Boolean);
    ps = ps.filter((p, k) => ps.indexOf(p) === k);
    const rf = i.rareflag || 0;
    return {
      id: `ea${i.id}`, eaId: i.id, definitionId: i.definitionId, name: nameOf(i), rating: i.rating, positions: ps,
      nation: nation(i.nationId), league: league(i.leagueId), club: club(i.teamId),
      rarity: rf === 12 ? 'icon' : rf > 0 ? 'rare' : 'common', rare: rf > 0,
      untradeable: isUntradeable(i),
      loan: !!((i.isLimitedUse && i.isLimitedUse()) || i.loans > 0),
    };
  }

  const ready = () => { try { return !!(window.services && services.User && services.User.getUser() && services.Club && services.SBC); } catch { return false; } };

  let clubLoading = null;
  let clubProgress = null;
  function loadClub(onProgress) {
    if (onProgress) clubProgress = onProgress;
    if (!clubLoading) clubLoading = loadClubNow(n => clubProgress && clubProgress(n)).finally(() => { clubLoading = null; clubProgress = null; });
    return clubLoading;
  }
  async function loadClubNow(onProgress) {
    const seen = {}; const out = []; let offset = 0; const count = 91; let repeats = 0; const sizes = []; let loans = 0;
    for (let page = 0; page < 120; page++) {
      const c = new UTSearchCriteriaDTO();
      c.type = SearchType.PLAYER; c.count = count; c.offset = offset;
      const r = await obs(services.Club.search(c));
      const list = body(r).items || [];
      let fresh = 0;
      for (const i of list) {
        if (seen[i.id]) continue;
        seen[i.id] = 1; fresh++; items[i.id] = i;
        const m = mapItem(i);
        if (!m.loan) out.push(m); else loans++;
      }
      sizes.push(list.length);
      onProgress && onProgress(out.length);
      if (!list.length) break;
      if (!fresh && ++repeats >= 2) break;
      offset += list.length;
      await sleep(60);
    }
    S.club = out; S.clubAt = Date.now(); save();
    log(`Club loaded: ${out.length} players (${loans} loans skipped, ${sizes.length} pages: ${sizes.join(',')}).`, 'good');
    return out;
  }

  // ------------------------------------------------------------ SBC requirement mapping
  const KEYN = k => { const K = window.SBCEligibilityKey || {}; for (const n in K) if (K[n] === k) return n; return String(k); };
  const SCOPEN = s => { const K = window.SBCEligibilityScope || {}; for (const n in K) if (K[n] === s) return n; return String(s); };

  function readReqs(ch) {
    return (ch.eligibilityRequirements || []).map(e => {
      const pairs = [];
      try {
        const col = e.kvPairs && (e.kvPairs._collection || e.kvPairs);
        if (col && typeof col.forEach === 'function' && !Array.isArray(col)) col.forEach((v, k) => pairs.push([KEYN(+k), v]));
        else for (const k in col) pairs.push([KEYN(+k), col[k]]);
      } catch { /* ignore */ }
      const key = pairs.length ? pairs[0][0] : KEYN(e.key);
      const vals = pairs.length ? [].concat(pairs[0][1]) : [];
      const values = vals.map(v => (/NATION_ID/.test(key) ? nation(v) : /LEAGUE_ID/.test(key) ? league(v) : /CLUB_ID/.test(key) ? club(v) : v));
      return { key, scope: SCOPEN(e.scope), count: e.count, values };
    });
  }

  function translate(r) {
    const op = /GREATER/.test(r.scope) ? 'min' : /LOWER/.test(r.scope) ? 'max' : 'exact';
    const v = r.values[0];
    const n = Number(r.count > 0 ? r.count : v);
    switch (r.key) {
      case 'TEAM_RATING': case 'TEAM_STAR_RATING': return { t: 'rating', n: Number(v) };
      case 'CHEMISTRY_POINTS': case 'TEAM_CHEMISTRY': return { t: 'chem', n: Number(v) };
      case 'ALL_PLAYERS_CHEMISTRY_POINTS': return { t: 'playerChem', n: Number(v) };
      case 'NATION_ID': return { t: 'count', attr: 'nation', value: v, op, n };
      case 'LEAGUE_ID': return { t: 'count', attr: 'league', value: v, op, n };
      case 'CLUB_ID': return { t: 'count', attr: 'club', value: v, op, n };
      case 'SAME_NATION_COUNT': return { t: 'same', attr: 'nation', op, n: Number(v) };
      case 'SAME_LEAGUE_COUNT': return { t: 'same', attr: 'league', op, n: Number(v) };
      case 'SAME_CLUB_COUNT': return { t: 'same', attr: 'club', op, n: Number(v) };
      case 'NATION_COUNT': return { t: 'distinct', attr: 'nation', op, n: Number(v) };
      case 'LEAGUE_COUNT': return { t: 'distinct', attr: 'league', op, n: Number(v) };
      case 'CLUB_COUNT': return { t: 'distinct', attr: 'club', op, n: Number(v) };
      case 'PLAYER_RARITY': case 'PLAYER_RARITY_GROUP': return Number(v) >= 1 ? { t: 'count', attr: 'rare', value: true, op, n } : null;
      case 'PLAYER_QUALITY': case 'PLAYER_LEVEL': {
        const q = { 1: 'bronze', 2: 'silver', 3: 'gold' }[Number(v)];
        if (!q) return null;
        if (op === 'exact') return { t: 'count', attr: 'quality', value: q, op: 'exact', n: 11 };
        return { t: 'ovr', min: op === 'min' ? { bronze: 0, silver: 65, gold: 75 }[q] : null, max: op === 'max' ? { bronze: 64, silver: 74, gold: 99 }[q] : null };
      }
      case 'PLAYER_MIN_OVR': return { t: 'ovr', min: Number(v), max: null };
      case 'PLAYER_MAX_OVR': return { t: 'ovr', min: null, max: Number(v) };
      default: return null;
    }
  }

  function slotsOf(ch) {
    const out = []; let bricks = 0;
    try {
      const field = (ch.squad.getFieldPlayers && ch.squad.getFieldPlayers()) || [];
      for (const s of field) {
        const p = s.getPosition ? s.getPosition() : s.position;
        const id = p && (p.typeName || p.id);
        out.push(pos(typeof id === 'number' ? id : String(id || '').replace(/^[RL](?=(CB|CM|DM|AM|S|F)$)/, '').replace(/^DM$/, 'CDM').replace(/^AM$/, 'CAM').replace(/^S$/, 'ST').replace(/^F$/, 'CF')));
        if (s.isBrick && s.isBrick()) bricks++;
      }
    } catch { /* ignore */ }
    return { slots: out, bricks };
  }

  function formationFromSlots(positions) {
    const ROW = { GK: 90, LB: 70, RB: 70, CB: 74, LWB: 62, RWB: 62, CDM: 56, LM: 44, RM: 44, CM: 50, CAM: 36, LW: 24, RW: 24, CF: 22, ST: 15 };
    const SIDE = { LB: 12, LWB: 10, LM: 12, LW: 16, RB: 88, RWB: 90, RM: 88, RW: 84 };
    const groups = {};
    positions.forEach((p, i) => { (groups[p] = groups[p] || []).push(i); });
    const slots = positions.map(p => ({ pos: p, x: 50, y: ROW[p] ?? 50 }));
    for (const [p, idx] of Object.entries(groups)) {
      if (SIDE[p] != null) { idx.forEach(i => (slots[i].x = SIDE[p])); continue; }
      idx.forEach((i, k) => (slots[i].x = idx.length === 1 ? 50 : 22 + (56 * k) / (idx.length - 1)));
    }
    return { name: 'EA squad', slots };
  }

  // ------------------------------------------------------------ live market prices
  // Lowest Buy Now price on the transfer market. Searches, then searches again
  // below the cheapest price seen while the results page is full, until nothing
  // cheaper appears. Searches are spaced out and capped per hour, because EA
  // locks the market (or shows a captcha) when it's searched too fast.
  const PRICE_TTL = 6 * 3600000; // how long a price is shown
  const FRESH = 30 * 60000;       // how old a price may be when relisting
  let lastSearch = 0;
  let runTok = 0; // bumped by Stop; every long task checks it
  class Stopped extends Error {}
  const checkTok = tok => { if (tok !== runTok) throw new Stopped('Stopped'); };
  async function csleep(ms, tok) {
    const end = Date.now() + ms;
    while (Date.now() < end) { checkTok(tok); await sleep(Math.min(150, end - Date.now())); }
    checkTok(tok);
  }
  const priceOf = (defId, maxAge = PRICE_TTL) => { const c = S.prices[defId]; return c && Date.now() - c.at < maxAge ? c : null; };

  function searchCount() {
    const now = Date.now();
    S.searches = S.searches.filter(t => now - t < 3600000);
    return S.searches.length;
  }

  async function marketSearch(defId, maxBuy, tok = runTok) {
    checkTok(tok);
    if (searchCount() >= S.limits.search) throw new Error(`Market search limit reached (${S.limits.search} per hour). Try again later.`);
    const gap = 2000 + Math.random() * 2000;
    const wait = lastSearch + gap - Date.now();
    if (wait > 0) await csleep(wait, tok);
    lastSearch = Date.now();
    S.searches.push(lastSearch);
    const c = new UTSearchCriteriaDTO();
    c.type = SearchType.PLAYER;
    c.maskedDefId = defId;
    if (maxBuy) c.maxBuy = maxBuy;
    const r = await obs(services.Item.searchTransferMarket(c, 1));
    checkTok(tok);
    return (body(r).items || [])
      .map(i => ((i.getAuctionData && i.getAuctionData()) || i._auction || {}).buyNowPrice)
      .filter(v => v > 0);
  }

  async function marketPrice(defId, force, tok = runTok) {
    const cached = !force && priceOf(defId, typeof force === 'number' ? force : PRICE_TTL);
    if (cached) return cached.p;
    let best = null; let maxBuy = null; let seen = 0;
    for (let round = 0; round < 5; round++) {
      const list = await marketSearch(defId, maxBuy, tok);
      seen += list.length;
      if (!list.length) break;
      const m = Math.min(...list);
      best = best == null ? m : Math.min(best, m);
      if (list.length < 20) break;
      maxBuy = E.roundToMarket(best - E.priceStep(best), 'down');
      if (maxBuy < 200) break;
    }
    S.prices[defId] = { p: best, at: Date.now(), seen };
    save();
    paintPrices(defId);
    const who = (S.club || []).find(x => x.definitionId === defId);
    log(`Price ${who ? who.name : defId}: ${best ? fmt(best) : 'no listings found'} (${seen} listings seen).`, best ? 'info' : 'bad');
    return best;
  }

  const kfmt = n => (n == null ? '—' : n >= 1e6 ? `${(n / 1e6).toFixed(2)}M` : n >= 1e4 ? `${(n / 1e3).toFixed(1)}k` : fmt(n));

  let bgBusy = false; let bgNext = 0;
  async function backgroundPrice() {
    if (!S.bgPrices || bgBusy || !S.club || Date.now() < bgNext || ui.busy) return;
    const budget = Math.min(S.limits.search - 10, S.bgPerHour || 40);
    const bgCount = (S.bgSearches = (S.bgSearches || []).filter(t => Date.now() - t < 3600000)).length;
    if (bgCount >= budget || searchCount() >= S.limits.search - 10) return;
    const seenDef = new Set();
    const next = S.club.filter(x => !x.untradeable && !S.locked[x.definitionId] && !priceOf(x.definitionId) && !seenDef.has(x.definitionId) && seenDef.add(x.definitionId))
      .sort((a, b) => b.rating - a.rating)[0];
    if (!next) return;
    bgBusy = true;
    try {
      const before = S.searches.length;
      await marketPrice(next.definitionId);
      for (let k = before; k < S.searches.length; k++) S.bgSearches.push(Date.now());
      save(); render();
    } catch (e) { log(`Background price check paused: ${e.message}`, 'bad'); bgNext = Date.now() + 10 * 60000; }
    finally { bgBusy = false; bgNext = Math.max(bgNext, Date.now() + 30000 + Math.random() * 20000); }
  }

  // Prices inside EA's own lists: EA draws each player row with a view class;
  // we wrap its render so a price line is added under EA's own row. EA's card,
  // player details and actions stay EA's.
  const hooked = [];
  function rowItem(view, arg) {
    const cands = [arg, view && view.data, view && view._data, view && view.item, view && view._item, view && view.model];
    return cands.find(x => x && typeof x === 'object' && x.definitionId) || null;
  }
  function rowEl(view) {
    try { if (view.getRootElement) return view.getRootElement(); } catch { /* ignore */ }
    return [view.__root, view._root, view.rootElement, view.element].find(x => x instanceof Element) || null;
  }
  function priceText(def, item) {
    const pr = priceOf(def);
    if (!pr) return 'Tap for price';
    return pr.p ? `${kfmt(pr.p)}` : 'none listed';
  }
  function decorate(view, arg) {
    const item = rowItem(view, arg);
    if (!item || (item.type && item.type !== 'player')) return;
    const el = rowEl(view);
    if (!el) return;
    let tag = el.querySelector('.ss-price');
    if (!tag) {
      tag = document.createElement('div');
      tag.className = 'ss-price';
      tag.addEventListener('click', ev => {
        ev.preventDefault(); ev.stopPropagation();
        const d = +tag.dataset.def;
        tag.classList.add('ss-loading'); tag.querySelector('b').textContent = 'checking…';
        marketPrice(d, true).catch(e => { if (!(e instanceof Stopped)) toastAnywhere(e.message); }).finally(() => paintPrices(d));
      }, true);
      const anchor = el.querySelector('.rowContent, .entityContainer, .name') || el;
      (anchor === el ? el : anchor.parentElement).appendChild(tag);
    }
    tag.dataset.def = item.definitionId;
    tag.innerHTML = `<i></i><b>${priceText(item.definitionId, item)}</b>${isUntradeable(item) ? '<span class="ss-ut">untradeable</span>' : ''}`;
    tag.classList.toggle('ss-none', !priceOf(item.definitionId));
  }
  function paintPrices(def) {
    document.querySelectorAll(`.ss-price[data-def="${def}"]`).forEach(t => {
      t.classList.remove('ss-loading'); t.classList.toggle('ss-none', !priceOf(def));
      const b = t.querySelector('b'); if (b) b.textContent = priceText(def);
    });
    if (ui.open) render();
  }
  function toastAnywhere(msg) { if (ui.open) toast(msg, true); else log(msg, 'bad'); }
  function hookRows() {
    const names = Object.getOwnPropertyNames(window).filter(n => /^UT\w*(ItemTableCell|ItemCell|PlayerItem|ItemList|ListItem|ItemRow)\w*View$/.test(n));
    for (const n of names) {
      let proto;
      try { proto = window[n] && window[n].prototype; } catch { continue; }
      if (!proto) continue;
      for (const m of ['render', 'setData', 'setItem', 'update']) {
        const orig = proto[m];
        if (typeof orig !== 'function' || orig.__ss) continue;
        const wrapped = function (...a) { const r = orig.apply(this, a); try { decorate(this, a[0]); } catch { /* ignore */ } return r; };
        wrapped.__ss = true;
        proto[m] = wrapped;
        hooked.push(`${n}.${m}`);
      }
    }
    if (hooked.length) log(`Prices shown in EA's lists (${hooked.slice(0, 4).join(', ')}).`, 'good');
    else log('Could not find EA\'s player row views. Send Safety → Copy page report.', 'bad');
  }
  function rowStyle() {
    if (document.getElementById('ss-rows')) return;
    const st = document.createElement('style');
    st.id = 'ss-rows';
    st.textContent = `.ss-price{display:flex;align-items:center;gap:6px;margin-top:6px;font-weight:700;font-size:15px;color:#f7d75c;cursor:pointer;position:relative;z-index:2}
      .ss-price i{width:14px;height:14px;border-radius:50%;background:radial-gradient(circle at 35% 35%,#ffe9a3,#d9a930 60%,#a87912);display:inline-block}
      .ss-price.ss-none b{color:rgba(255,255,255,.6);font-weight:600}
      .ss-price.ss-loading b{color:rgba(255,255,255,.75)}
      .ss-price .ss-ut{margin-left:4px;font-size:12px;font-weight:700;color:#cfe0ff;background:rgba(80,110,160,.45);border-radius:6px;padding:1px 6px}`;
    document.head.appendChild(st);
  }

  // ------------------------------------------------------------ submission limits
  function subCounts() {
    const now = Date.now();
    S.subs = S.subs.filter(t => now - t < 86400000);
    return { hour: S.subs.filter(t => now - t < 3600000).length, day: S.subs.length };
  }

  // ------------------------------------------------------------ UI
  const ui = { open: false, view: 'hub', busy: '', sets: null, openSet: null, chs: null, pick: null, result: null, seed: 1, q: '', trades: null, confirm: null, report: '' };
  const h = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const fmt = n => Math.round(n || 0).toLocaleString();
  const ago = t => { const m = Math.round((Date.now() - t) / 60000); return m < 1 ? 'just now' : m < 60 ? `${m} min ago` : `${Math.round(m / 60)} h ago`; };
  const q = c => (c >= 75 ? 'gold' : c >= 65 ? 'silver' : 'bronze');

  const ICON = {
    logo: '<svg viewBox="0 0 24 24"><path d="M12 2.5l2.8 5.7 6.3.9-4.6 4.4 1.1 6.2L12 16.8l-5.6 2.9 1.1-6.2-4.6-4.4 6.3-.9z"/></svg>',
    sbc: '<svg viewBox="0 0 64 64"><path d="M32 4l7 6h9v9l6 7-6 7v9h-9l-7 6-7-6h-9v-9l-6-7 6-7v-9h9z" fill="none" stroke="#fff" stroke-width="4"/><path d="M24 33l6 6 11-13" fill="none" stroke="#fff" stroke-width="5" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    club: '<svg viewBox="0 0 64 64"><path d="M22 8l10 6 10-6 16 9-6 12-6-3v30H18V26l-6 3-6-12z" fill="#fff"/></svg>',
    trades: '<svg viewBox="0 0 64 64"><path d="M10 22h36l-8-8M54 42H18l8 8" fill="none" stroke="#fff" stroke-width="6" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    safety: '<svg viewBox="0 0 64 64"><path d="M32 6l22 8v16c0 14-9 24-22 29C19 54 10 44 10 30V14z" fill="#fff"/><path d="M23 32l7 7 12-14" fill="none" stroke="#2c3a4e" stroke-width="5" stroke-linecap="round" stroke-linejoin="round"/></svg>',
  };
  const TITLES = { hub: 'SquadSage', sbc: 'SBC Solver', club: 'My Club', trades: 'Trades', settings: 'Safety' };

  const host = document.createElement('div');
  host.id = 'squadsage-host';
  const root = host.attachShadow({ mode: 'open' });
  root.innerHTML = `<style>${CSS()}</style>
    <button class="fab" data-a="open" aria-label="Open SquadSage" hidden>${ICON.logo}</button>
    <div class="page" hidden>
      <div class="top"><button class="back" data-a="home" aria-label="Back">‹</button><h2 class="title">SquadSage</h2><button class="x" data-a="close" aria-label="Close" hidden>×</button></div>
      <div class="body"></div>
    </div>
    <div class="toast" hidden></div>`;
  const $ = s => root.querySelector(s);

  function toast(msg, bad) {
    const t = $('.toast');
    t.textContent = msg; t.className = `toast ${bad ? 'bad' : ''}`; t.hidden = false;
    clearTimeout(toast.t); toast.t = setTimeout(() => (t.hidden = true), 3500);
  }

  function render() {
    $('.page').hidden = !ui.open;
    $('.fab').hidden = ui.open || mode !== 'float';
    $('.x').hidden = mode !== 'float';
    $('.back').style.visibility = ui.view === 'hub' ? 'hidden' : 'visible';
    $('.title').textContent = TITLES[ui.view];
    if (!ui.open) return;
    const top = host.scrollTop;
    $('.body').innerHTML = (ui.busy ? `<div class="busy"><span class="spin"></span><span class="grow">${h(ui.busy)}</span>${ui.stoppable ? '<button class="btn stop" data-a="stop">Stop</button>' : ''}</div>` : '') + VIEW[ui.view]() + (ui.confirm ? confirmBox() : '');
    host.scrollTop = top;
  }

  // ---- placing SquadSage inside EA's own app ---------------------------------
  // mode 'tab':   a SquadSage tab in EA's bottom bar; our page fills EA's content area
  // mode 'float': fallback when EA's bar can't be found; a small button opens a full page
  let mode = null; let navBtn = null; let content = null; let tries = 0; let booted = false; let clubAuto = false;
  const NAV_SEL = ['.ut-tab-bar', 'nav.ut-tab-bar', '.ut-tab-bar-view', 'nav[class*="tab-bar"]', '[class*="tab-bar"]'];
  const CONTENT_SEL = ['.ut-content', '.ut-navigation-container-view--content', '.ut-content-container', 'main'];

  function findNav() {
    for (const sel of NAV_SEL) {
      for (const n of document.querySelectorAll(sel)) {
        if (n.closest('#squadsage-host')) continue;
        if (n.querySelectorAll('button, a').length >= 3) return n;
      }
    }
    return null;
  }
  function findContent(nav) {
    for (const sel of CONTENT_SEL) { const c = document.querySelector(sel); if (c && !c.contains(nav)) return c; }
    return nav && nav.parentElement && [...nav.parentElement.children].filter(c => c !== nav && c.offsetHeight > 200)[0];
  }
  function tabItems(nav) { return [...nav.querySelectorAll('button, a')].filter(b => b !== navBtn && !navBtn?.contains(b)); }

  function globalStyle() {
    if (document.getElementById('ss-global')) return;
    const st = document.createElement('style');
    st.id = 'ss-global';
    st.textContent = `.ss-tab::before,.ss-tab::after{content:none!important;display:none!important}
      .ss-tab .ss-ico{display:block;width:26px;height:26px;margin:0 auto 4px;fill:currentColor}
      .ss-tab .ss-ico svg{width:100%;height:100%;fill:currentColor}`;
    document.head.appendChild(st);
  }

  function integrate() {
    const nav = findNav();
    if (!nav) return false;
    const its = tabItems(nav);
    if (!its.length) return false;
    if (!navBtn || !nav.contains(navBtn)) {
      globalStyle();
      const tpl = its[0];
      navBtn = tpl.cloneNode(true);
      navBtn.removeAttribute('id');
      const strip = el => { if (typeof el.className === 'string') el.className = el.className.split(/\s+/).filter(c => !/^icon-|selected|active/.test(c)).join(' '); };
      strip(navBtn); navBtn.querySelectorAll('*').forEach(strip);
      navBtn.classList.add('ss-tab');
      // replace the label text, drop any EA icon image, add ours
      const walker = document.createTreeWalker(navBtn, NodeFilter.SHOW_TEXT);
      let labelled = false;
      for (let n = walker.nextNode(); n; n = walker.nextNode()) {
        if (!n.textContent.trim()) continue;
        n.textContent = labelled ? '' : 'SquadSage'; labelled = true;
      }
      if (!labelled) navBtn.append('SquadSage');
      navBtn.querySelectorAll('img, svg').forEach(x => x.remove());
      const ico = document.createElement('span');
      ico.className = 'ss-ico'; ico.innerHTML = ICON.logo;
      navBtn.prepend(ico);
      tpl.after(navBtn);
      navBtn.addEventListener('click', e => { e.preventDefault(); e.stopPropagation(); openPage(); }, true);
      nav.addEventListener('click', e => { if (!navBtn.contains(e.target)) closePage(); }, true);
    }
    content = findContent(nav);
    return !!content;
  }

  function placeHost() {
    if (mode === 'tab') {
      if (getComputedStyle(content).position === 'static') content.style.position = 'relative';
      if (host.parentNode !== content) content.appendChild(host);
      host.style.cssText = 'position:absolute;inset:0;z-index:60;overflow:auto;-webkit-overflow-scrolling:touch';
    } else {
      if (host.parentNode !== document.body) document.body.appendChild(host);
      host.style.cssText = `position:fixed;inset:0;z-index:2147483600;overflow:${ui.open ? 'auto' : 'visible'};pointer-events:${ui.open ? 'auto' : 'none'}`;
    }
    host.style.display = ui.open || mode === 'float' ? '' : 'none';
  }

  function openPage() {
    ui.open = true;
    if (mode === 'tab') {
      const nav = findNav();
      if (nav) tabItems(nav).forEach(b => b.classList.remove('selected'));
      navBtn && navBtn.classList.add('selected');
    }
    placeHost(); render();
    if (!ui.sets) loadSets().catch(() => {});
    if (!ui.trades) A.trades(true);
  }
  function closePage() {
    if (!ui.open) return;
    ui.open = false; ui.confirm = null;
    navBtn && navBtn.classList.remove('selected');
    placeHost(); render();
  }

  // Short outline of EA's page, so a missing tab can be fixed quickly.
  function pageReport() {
    const out = [];
    const walk = (el, d) => {
      if (d > 6 || out.length > 300) return;
      for (const c of el.children) {
        if (c.id === 'squadsage-host' || /^(SCRIPT|STYLE|LINK|META)$/.test(c.tagName)) continue;
        const cls = typeof c.className === 'string' && c.className.trim() ? `.${c.className.trim().split(/\s+/).slice(0, 4).join('.')}` : '';
        const txt = c.children.length ? '' : ` "${(c.textContent || '').trim().slice(0, 24)}"`;
        out.push(`${'  '.repeat(d)}${c.tagName.toLowerCase()}${c.id ? `#${c.id}` : ''}${cls}${txt}`);
        walk(c, d + 1);
      }
    };
    walk(document.body, 0);
    let sample = '';
    try {
      const it = Object.values(items)[0];
      if (it) {
        const own = Object.keys(it).map(k => `${k}=${typeof it[k] === 'object' ? (it[k] && it[k].constructor ? it[k].constructor.name : 'obj') : String(it[k]).slice(0, 20)}`);
        const fns = []; let pr = Object.getPrototypeOf(it);
        for (let d = 0; pr && d < 3; d++, pr = Object.getPrototypeOf(pr)) fns.push(...Object.getOwnPropertyNames(pr).filter(n => /trad|price|name|untrad|league|nation|team|rare|loan/i.test(n)));
        sample = `\nITEM ${own.join(' ')}\nFNS ${[...new Set(fns)].join(' ')}`;
      }
      sample += `\nHOOKED ${hooked.join(' ') || 'none'}\nVIEWS ${Object.getOwnPropertyNames(window).filter(n => /^UT\w*View$/.test(n)).slice(0, 120).join(' ')}`;
      const lk = services.Localization.localize(`global.leagueFull.2027.league13`);
      sample += `\nLOC ${lk}`;
    } catch (e) { sample += `\nsample error ${e.message}`; }
    return `mode=${mode} nav=${!!findNav()} content=${content ? content.className : 'none'} url=${location.pathname}${sample}\n${out.join('\n')}`;
  }

  function confirmBox() {
    return `<div class="dim"><div class="dlg"><p>${h(ui.confirm.msg)}</p><div class="row"><button class="btn primary" data-a="yes">${h(ui.confirm.yes)}</button><button class="btn" data-a="no">Cancel</button></div></div></div>`;
  }
  const ask = (msg, yes, fn) => { ui.confirm = { msg, yes, fn }; render(); };

  const stat = (l, v, cls = '') => `<div class="stat ${cls}"><span>${h(l)}</span><b>${v}</b></div>`;

  const tile = (v, title, icon, badge) => `<button class="tile" data-a="go" data-v="${v}"><span class="tt">${title}</span><span class="ti">${icon}</span>${badge ? `<span class="tb">${h(badge)}</span>` : ''}</button>`;

  const VIEW = {
    hub() {
      const c = subCounts();
      const exp = ui.trades ? ui.trades.filter(i => i.state === 'expired').length : null;
      const locked = Object.keys(S.locked).length;
      return `<div class="tiles">
        ${tile('sbc', 'SBC Solver', ICON.sbc, ui.sets ? `${ui.sets.length} open` : '')}
        ${tile('club', 'My Club', ICON.club, S.club ? `${fmt(S.club.length)} players${locked ? ` · ${locked} 🔒` : ''}` : '')}
        ${tile('trades', 'Trades', ICON.trades, exp != null ? `${exp} expired` : '')}
        ${tile('settings', 'Safety', ICON.safety, `${c.hour}/${S.limits.hour} SBCs this hour`)}
      </div>`;
    },
    sbc() {
      const c = subCounts();
      const lim = `<div class="limits ${c.hour >= S.limits.hour || c.day >= S.limits.day ? 'bad' : ''}">Submitted by SquadSage: <b>${c.hour}/${S.limits.hour}</b> this hour · <b>${c.day}/${S.limits.day}</b> today</div>`;
      if (ui.result) return lim + resultView();
      const clubLine = S.club ? `<p class="hint">${fmt(S.club.length)} club players · loaded ${ago(S.clubAt)} · <a data-a="loadClub">reload</a></p>` : '';
      if (!ui.sets) return lim + clubLine + '<p class="hint">Loading your SBCs…</p>';
      return lim + clubLine + `<ul class="sets">${ui.sets.map(s => `<li class="${ui.openSet === s ? 'open' : ''}">
          <button class="row-btn" data-a="set" data-i="${ui.sets.indexOf(s)}"><b>${h(s.name)}</b><small>${s.challengesCompletedCount || 0}/${s.challengesCount || '?'} done</small></button>
          ${ui.openSet === s ? (ui.chs ? `<ul class="chs">${ui.chs.map((ch, i) => `<li><button class="row-btn" data-a="pick" data-i="${i}"><b>${h(ch.name)}</b><small>tap to solve</small></button></li>`).join('') || '<li class="hint">Everything here is done.</li>'}</ul>` : '<p class="hint">Loading challenges…</p>') : ''}
        </li>`).join('') || '<li class="hint">No open SBCs right now.</li>'}</ul>`;
    },
    club() {
      if (!S.club) return '<p class="hint">Your club hasn\'t been loaded yet.</p><button class="btn primary" data-a="loadClub">Load club</button>';
      const locked = S.club.filter(p => S.locked[p.definitionId]).length;
      const ut = S.club.filter(p => p.untradeable).length;
      const qq = ui.q.toLowerCase();
      const list = S.club.filter(p => !qq || `${p.name} ${p.nation} ${p.league} ${p.club} ${p.positions.join(' ')}`.toLowerCase().includes(qq))
        .sort((a, b) => b.rating - a.rating).slice(0, 150);
      return `<div class="stats">${stat('Players', fmt(S.club.length))}${stat('Untradeable', fmt(ut))}${stat('Locked', fmt(locked))}${stat('Loaded', ago(S.clubAt))}</div>
        <p class="hint">Locked players are never used by the SBC solver. Tap 🔒 to lock or unlock. Tap a tradeable player's price to check it live on the market.</p>

        <div class="row"><input type="search" placeholder="Search name, league, nation…" data-in="q" value="${h(ui.q)}"><button class="btn" data-a="loadClub">Reload</button></div>
        <ul class="players">${list.map(p => `<li><span class="badge ${q(p.rating)}">${p.rating}</span><div class="grow"><b>${h(p.name)}</b><small>${h(p.positions.join('/'))} · ${h(p.league)}${p.untradeable ? ' · UT' : ''}</small></div>
          ${p.untradeable ? '' : `<button class="price ${priceOf(p.definitionId) ? 'on' : ''}" data-a="price1" data-d="${p.definitionId}">${priceOf(p.definitionId) ? (priceOf(p.definitionId).p ? kfmt(priceOf(p.definitionId).p) : 'none') : 'price?'}</button>`}
          <button class="lock ${S.locked[p.definitionId] ? 'on' : ''}" data-a="lock" data-d="${p.definitionId}">${S.locked[p.definitionId] ? '🔒' : '🔓'}</button></li>`).join('')}</ul>`;
    },
    trades() {
      const t = ui.trades;
      if (!t) return '<p class="hint">Loading your transfer list…</p>';
      let potIncome = 0, potProfit = 0, soldIncome = 0, soldProfit = 0, expired = 0, active = 0, unpriced = 0;
      for (const i of t) {
        if (i.state === 'active') { active++; potIncome += E.afterTax(i.bin); if (i.paid) potProfit += E.afterTax(i.bin) - i.paid; else unpriced++; }
        if (i.state === 'expired') expired++;
        if (i.state === 'closed') { const v = E.afterTax(i.bid || i.bin); soldIncome += v; if (i.paid) soldProfit += v - i.paid; else unpriced++; }
      }
      const sign = n => `${n >= 0 ? '+' : ''}${fmt(n)}`;
      return `<div class="stats">
          ${stat('Selling now', active)}${stat('Expired', expired, expired ? 'warn' : '')}
          ${stat('Income if they sell', fmt(potIncome))}${stat('Profit if they sell', sign(potProfit), potProfit >= 0 ? 'good' : 'bad')}
          ${stat('Sold income', fmt(soldIncome))}${stat('Sold profit', sign(soldProfit), soldProfit >= 0 ? 'good' : 'bad')}
        </div>
        <p class="hint">All figures are after EA's 5% tax.${unpriced ? ` ${unpriced} items have no known buy price (packs or SBC rewards) and are left out of profit.` : ''}</p>
        <div class="row"><button class="btn primary" data-a="smartRelist" ${expired && !ui.busy ? '' : 'disabled'}>Relist all at market price (${expired})</button><button class="btn" data-a="relist" ${expired && !ui.busy ? '' : 'disabled'}>Relist at same price</button><button class="btn" data-a="trades">Refresh</button></div>
        <label class="chk"><input type="checkbox" data-in="undercut" ${S.undercut ? 'checked' : ''}> Smart relist: list one price step under the cheapest</label>
        <label class="chk"><input type="checkbox" data-in="floorBE" ${S.floorBE ? 'checked' : ''}> Never relist below what I paid (after tax)</label>
        ${expired ? `<h3>Expired</h3><ul class="players">${t.filter(i => i.state === 'expired').map(i => `<li><span class="badge ${q(i.rating)}">${i.rating}</span><div class="grow"><b>${h(i.name)}</b><small>listed ${fmt(i.bin)}${i.paid ? ` · paid ${fmt(i.paid)}` : ''}</small></div>
          <button class="price ${priceOf(i.defId) ? 'on' : ''}" data-a="price1" data-d="${i.defId}">${priceOf(i.defId) ? (priceOf(i.defId).p ? kfmt(priceOf(i.defId).p) : 'none') : 'price?'}</button></li>`).join('')}</ul>` : ''}`;
    },
    settings() {
      return `<h3>SBC submission limits</h3>
        <p class="hint">EA soft-bans accounts that submit SBCs too fast. SquadSage stops submitting when either limit is reached. It only counts submissions it made itself.</p>
        <div class="row"><label>Per hour<input type="number" min="1" data-in="lh" value="${S.limits.hour}"></label><label>Per day<input type="number" min="1" data-in="ld" value="${S.limits.day}"></label></div>
        <h3>Market price checks</h3>
        <p class="hint">Searches this hour: <b>${searchCount()}/${S.limits.search}</b>. Searches are spaced 2–4 seconds apart; prices are remembered for 30 minutes.</p>
        <div class="row"><label>Searches per hour<input type="number" min="10" data-in="ls" value="${S.limits.search}"></label></div>
        <label class="chk"><input type="checkbox" data-in="bg" ${S.bgPrices ? 'checked' : ''}> Check prices in the background (one search every 30–50 s, up to ${S.bgPerHour || 40} per hour, best players first)</label>
        <p class="hint">${S.club ? `${S.club.filter(x => !x.untradeable && priceOf(x.definitionId)).length} of ${S.club.filter(x => !x.untradeable).length} tradeable players priced.` : ''}</p>
        <label class="chk"><input type="checkbox" data-in="ut" ${S.onlyUT ? 'checked' : ''}> Solve with untradeable players only</label>
        <p class="warn">Automating the EA Web App is against EA's terms. EA can restrict or ban accounts for it.</p>
        <h3>Activity</h3>
        <ul class="log">${S.log.slice(0, 30).map(l => `<li class="${l.kind}"><time>${new Date(l.at).toLocaleTimeString()}</time> ${h(l.msg)}</li>`).join('') || '<li class="hint">Nothing yet.</li>'}</ul>
        <div class="row"><button class="btn" data-a="copylog">Copy log</button><button class="btn" data-a="report">Copy page report</button><button class="btn" data-a="reset">Clear SquadSage data</button></div>
        ${ui.report ? `<p class="hint">If copying didn't work, long-press the box below, Select all, then Copy.</p><textarea readonly rows="8">${h(ui.report)}</textarea>` : ''}`;
    },
  };

  function resultView() {
    const r = ui.result; const p = ui.pick;
    if (r.error) return `<div class="callout bad">${h(r.error)}</div><button class="btn" data-a="back">Back</button>`;
    const ev = r.eval;
    const block = !r.ok ? 'Not every requirement can be met with your club.'
      : p.unsupported.length ? `Some requirements couldn't be read (${p.unsupported.join(', ')}), so check this squad yourself before submitting in EA's screen.`
      : r.squad.some(x => x.source !== 'club') ? 'This squad needs cards you would have to buy.' : '';
    const c = subCounts();
    const limitHit = c.hour >= S.limits.hour || c.day >= S.limits.day;
    return `<button class="link" data-a="back">‹ All SBCs</button>
      <h3 class="ch">${h(p.ch.name)}</h3>
      <div class="callout ${r.ok ? 'good' : 'bad'}">${r.ok ? 'All requirements met.' : 'No valid squad found with your club.'}</div>
      <div class="stats">${stat('Rating', ev.rating)}${stat('Chemistry', `${ev.chem.total}/33`)}</div>
      <ul class="checks">${ev.results.map(x => `<li class="${x.ok ? 'ok' : 'no'}"><span>${x.ok ? '✔' : '✖'}</span>${h(x.label)}<em>${h(x.actual)}</em></li>`).join('')}</ul>
      <ul class="players">${r.squad.map((x, i) => `<li><span class="slot">${h(p.formation.slots[i].pos)}</span><span class="badge ${q(x.rating)}">${x.rating}</span><div class="grow"><b>${h(x.name)}</b><small>${h(x.league)}${x.untradeable ? ' · UT' : ''}</small></div><span class="chem c${ev.chem.per[i]}">${'◆'.repeat(ev.chem.per[i])}</span></li>`).join('')}</ul>
      ${block ? `<p class="hint">${h(block)}</p>` : ''}
      ${limitHit ? '<p class="warn">Submission limit reached. Wait before submitting more.</p>' : ''}
      <div class="row"><button class="btn primary" data-a="submit" ${!block && !limitHit && !ui.busy ? '' : 'disabled'}>Submit to EA</button><button class="btn" data-a="again">Another solution</button></div>`;
  }

  // ------------------------------------------------------------ actions
  let busyId = 0;
  async function busy(label, fn, stoppable = false) {
    const id = ++busyId;
    ui.busy = label; ui.stoppable = stoppable; render();
    try { return await fn(); } catch (e) {
      if (e instanceof Stopped) { log(`${label.replace('…', '')} stopped.`); return null; }
      log(`${label.replace('…', '')} failed: ${e.message}`, 'bad'); toast(e.message, true); return null;
    } finally { if (id === busyId) { ui.busy = ''; ui.stoppable = false; render(); } }
  }

  async function loadSets() {
    const r = await obs(services.SBC.requestSets());
    const sets = body(r).sets || (body(r).categories || []).reduce((a, c) => a.concat(c.sets || []), []);
    ui.sets = sets.filter(s => !(s.isComplete && s.isComplete()));
    render();
  }

  async function ensureClub() {
    if (S.club && Date.now() - S.clubAt < 10 * 60000 && Object.keys(items).length) return;
    await busy('Loading your club…', () => loadClub(n => { ui.busy = `Loading your club… ${n} players`; render(); }));
  }

  function solveFor(p) {
    const exclude = S.onlyUT ? S.club.filter(x => !x.untradeable).map(x => x.id) : [];
    const clubForSolve = S.club.map(x => {
      const pr = priceOf(x.definitionId);
      const y = pr && pr.p ? { ...x, price: pr.p } : x;
      return S.locked[x.definitionId] ? { ...y, protected: true } : y;
    });
    return E.solve({ club: clubForSolve, formation: p.formation, reqs: p.reqs, useMarket: false, timeMs: 1400, seed: ui.seed * 7919, excludeIds: exclude });
  }

  function mapTrades(r) {
    return (body(r).items || []).map(i => {
      const a = (i.getAuctionData && i.getAuctionData()) || i._auction || {};
      return { id: i.id, defId: i.definitionId, name: nameOf(i), rating: i.rating, state: a.tradeState || null, bin: a.buyNowPrice || 0, bid: a.currentBid || 0, paid: i.lastSalePrice || 0, raw: i };
    });
  }

  const A = {
    open() { openPage(); },
    close() { closePage(); },
    home() { ui.view = 'hub'; ui.result = null; ui.pick = null; ui.confirm = null; render(); host.scrollTop = 0; },
    go(el) {
      ui.view = el.dataset.v; ui.confirm = null; render(); host.scrollTop = 0;
      if (ui.view === 'trades') A.trades();
      if (ui.view === 'club' && !S.club) A.loadClub();
      if (ui.view === 'sbc' && !ui.sets) busy('Loading your SBCs…', loadSets);
    },
    async report() {
      ui.report = pageReport(); render();
      try { await navigator.clipboard.writeText(ui.report); toast('Page report copied'); } catch { toast('Select the text in the box to copy it', true); }
    },
    loadClub() { busy('Loading your club…', () => loadClub(n => { ui.busy = `Loading your club… ${n} players`; render(); })); },
    async set(el) {
      const s = ui.sets[+el.dataset.i];
      if (ui.openSet === s) { ui.openSet = null; render(); return; }
      ui.openSet = s; ui.chs = null; render();
      await busy('Loading challenges…', async () => {
        const r = await obs(services.SBC.requestChallengesForSet(s));
        ui.chs = (body(r).challenges || []).filter(c => c.status !== 'COMPLETED' && !(c.isCompleted && c.isCompleted()));
      });
    },
    async pick(el) {
      const ch = ui.chs[+el.dataset.i];
      await ensureClub();
      if (!S.club) return;
      await busy('Solving…', async () => {
        try { await obs(services.SBC.loadChallenge(ch)); } catch { /* may already be loaded */ }
        const reqs = []; const unsupported = [];
        for (const r of readReqs(ch)) { const t = translate(r); if (t) reqs.push(t); else unsupported.push(`${r.key} ${r.values.join('/')}`); }
        const { slots, bricks } = slotsOf(ch);
        if (bricks) unsupported.push('locked slots');
        const sl = slots.filter(Boolean);
        if (sl.length !== 11) unsupported.push('squad positions');
        ui.pick = { ch, set: ui.openSet, reqs, unsupported, formation: sl.length === 11 ? formationFromSlots(sl) : { name: '4-3-3', slots: ['GK', 'LB', 'CB', 'CB', 'RB', 'CM', 'CM', 'CM', 'LW', 'ST', 'RW'].map((x, i) => ({ pos: x, x: 0, y: i })) } };
        await sleep(30);
        ui.result = solveFor(ui.pick);
        if (!ui.result.squad) ui.result = { error: ui.result.reason };
        log(`Solved "${ch.name}": ${ui.result.ok ? 'valid' : 'not valid'}.`, ui.result.ok ? 'good' : 'bad');
      });
    },
    again() { ui.seed++; busy('Solving…', async () => { await sleep(30); ui.result = solveFor(ui.pick); }); },
    back() { ui.result = null; ui.pick = null; render(); },
    submit() {
      const p = ui.pick; const r = ui.result;
      ask(`Submit "${p.ch.name}" with these 11 players? They will be used up.`, 'Submit', () => busy('Submitting…', async () => {
        const its = r.squad.map(x => items[x.eaId]);
        if (its.some(x => !x)) throw new Error('Some players are no longer in your club. Reload your club and solve again.');
        p.ch.squad.setPlayers(its, true);
        await obs(services.SBC.saveChallenge(p.ch));
        await sleep(300 + Math.random() * 400);
        const chemOn = services.Chemistry && services.Chemistry.isFeatureEnabled ? services.Chemistry.isFeatureEnabled() : true;
        await obs(services.SBC.submitChallenge(p.ch, p.set, true, chemOn));
        S.subs.push(Date.now());
        const used = new Set(r.squad.map(x => x.eaId));
        S.club = S.club.filter(x => !used.has(x.eaId));
        used.forEach(id => delete items[id]);
        save();
        log(`Submitted "${p.ch.name}".`, 'good');
        toast('Submitted to EA');
        ui.result = null; ui.pick = null; ui.chs = null; ui.openSet = null;
        loadSets().catch(() => {});
      }));
    },
    yes() { const f = ui.confirm.fn; ui.confirm = null; render(); f(); },
    no() { ui.confirm = null; render(); },
    lock(el) {
      const d = el.dataset.d;
      if (S.locked[d]) delete S.locked[d]; else S.locked[d] = 1;
      save(); render();
    },
    trades(quiet) {
      if (quiet === true) {
        obs(services.Item.requestTransferItems()).then(r => { ui.trades = mapTrades(r); render(); }).catch(() => {});
        return;
      }
      busy('Loading your transfer list…', async () => {
        const r = await obs(services.Item.requestTransferItems());
        ui.trades = mapTrades(r);
      });
    },

    relist() {
      busy('Relisting expired items…', async () => {
        await obs(services.Item.relistExpiredAuctions());
        log('Relisted expired items.', 'good'); toast('Relisted');
        A.trades();
      });
    },
    async price1(el) {
      const d = +el.dataset.d;
      const who = (S.club || []).find(x => x.definitionId === d);
      await busy(`Checking price${who ? ` of ${who.name}` : ''}…`, async () => {
        const p = await marketPrice(d, true);
        toast(p ? `Lowest Buy Now: ${fmt(p)}` : 'None on the market right now');
      }, true);
    },
    stop() {
      runTok++; busyId++;
      ui.busy = ''; ui.stoppable = false; render();
      toast('Stopped');
    },
    smartRelist() {
      const ex = (ui.trades || []).filter(i => i.state === 'expired');
      ask(`Check live prices for all ${ex.length} expired items and relist each at the market price? This uses up to ${ex.length * 3} market searches. You can stop at any time.`, 'Relist all', () => busy('Relisting at market price…', async () => {
        const tok = runTok;
        let done = 0, kept = 0, none = 0;
        for (const i of ex) {
          checkTok(tok);
          ui.busy = `Relisting at market price… ${done + kept + none + 1}/${ex.length} (${i.name})`; render();
          const m = await marketPrice(i.defId, FRESH, tok);
          if (!m) { none++; continue; }
          let bin = S.undercut ? E.roundToMarket(m - E.priceStep(m), 'down') : m;
          if (S.floorBE && i.paid) bin = Math.max(bin, E.roundToMarket(Math.ceil(i.paid / 0.95), 'up'));
          bin = Math.max(200, bin);
          const start = E.roundToMarket(bin - E.priceStep(bin), 'down');
          checkTok(tok);
          await obs(services.Item.list(i.raw, Math.max(150, start), bin, 3600));
          done++;
          await csleep(500 + Math.random() * 600, tok);
        }
        log(`Smart relist: ${done} relisted${none ? `, ${none} had no market price (left as is)` : ''}.`, 'good');
        toast(`Relisted ${done}`);
        A.trades();
      }, true));
    },
    async copylog() {
      const text = S.log.map(l => `${new Date(l.at).toISOString()} [${l.kind}] ${l.msg}`).join('\n');
      try { await navigator.clipboard.writeText(text); toast('Log copied'); } catch { toast('Could not copy. Long-press the log to select it.', true); }
    },
    reset() {
      ask('Clear SquadSage data (locks, cached club, counters, log)? Your EA account is not affected.', 'Clear', () => {
        localStorage.removeItem(KEY); Object.assign(S, { locked: {}, subs: [], club: null, clubAt: 0, log: [] }); render();
      });
    },
  };

  root.addEventListener('click', e => {
    const el = e.target.closest('[data-a]');
    if (!el || el.disabled) return;
    const f = A[el.dataset.a];
    if (f) { e.preventDefault(); f(el); }
  });
  root.addEventListener('input', e => {
    const k = e.target.dataset.in;
    if (k === 'q') { ui.q = e.target.value; const pos2 = e.target.selectionStart; render(); const i = root.querySelector('[data-in="q"]'); i.focus(); i.setSelectionRange(pos2, pos2); }
  });
  root.addEventListener('change', e => {
    const k = e.target.dataset.in;
    if (k === 'lh') S.limits.hour = Math.max(1, +e.target.value || 90);
    if (k === 'ld') S.limits.day = Math.max(1, +e.target.value || 300);
    if (k === 'ut') S.onlyUT = e.target.checked;
    if (k === 'bg') S.bgPrices = e.target.checked;
    if (k === 'ls') S.limits.search = Math.max(10, +e.target.value || 150);
    if (k === 'undercut') S.undercut = e.target.checked;
    if (k === 'floorBE') S.floorBE = e.target.checked;
    save();
  });

  // Mount once EA's app is up (the login page has no `services`), then keep
  // our tab in EA's bar even when EA re-renders it.
  (function mount() {
    if (!document.body) return setTimeout(mount, 300);
    if (!ready()) { host.remove(); return setTimeout(mount, 1000); }
    if (!mode) log('Connected to the EA Web App.', 'good');
    if (integrate()) {
      if (mode !== 'tab') log('SquadSage tab added to EA\'s bottom bar.', 'good');
      mode = 'tab';
    } else if (++tries >= 8 && mode !== 'tab') {
      if (mode !== 'float') log('Could not find EA\'s bottom bar; using the floating button. Send Safety → Copy page report.', 'bad');
      mode = 'float';
    }
    if (mode) { placeHost(); render(); }
    if (!booted) { booted = true; rowStyle(); hookRows(); setTimeout(() => { loadSets().catch(() => {}); }, 1500); }
    if (mode && ready()) {
      if (!S.club && !ui.busy && !clubAuto) { clubAuto = true; loadClub().catch(() => {}); }
      backgroundPrice();
    }
    setTimeout(mount, mode === 'tab' ? 2000 : 1000);
  })();

  function CSS() {
    return `
    :host { display: block; color: #fff; font-family: inherit; -webkit-tap-highlight-color: transparent; }
    * { box-sizing: border-box; font-family: inherit; }
    [hidden] { display: none !important; }
    .page { min-height: 100%; padding: 14px 16px calc(28px + env(safe-area-inset-bottom, 0px));
      background: radial-gradient(120% 70% at 85% 10%, rgba(32, 104, 104, .55), transparent 60%), linear-gradient(180deg, #17243a 0%, #123b45 100%); }
    .top { display: flex; align-items: center; gap: 6px; margin: 0 0 14px; }
    .title { margin: 0; font-size: 22px; font-weight: 700; flex: 1; }
    .back, .x { background: none; border: 0; color: #fff; font-size: 34px; line-height: 1; width: 34px; height: 40px; padding: 0; cursor: pointer; }
    .fab { pointer-events: auto; position: fixed; right: 14px; bottom: calc(96px + env(safe-area-inset-bottom, 0px)); width: 52px; height: 52px; border-radius: 50%;
      border: 0; background: linear-gradient(180deg, #f7d75c, #d9a92c); box-shadow: 0 6px 18px #0009; cursor: pointer; padding: 13px; }
    .fab svg { width: 100%; height: 100%; fill: #1b1406; }
    .tiles { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }
    .tile { display: flex; flex-direction: column; justify-content: flex-start; align-items: flex-start; position: relative; aspect-ratio: 1 / 1.12; border-radius: 22px; border: 1px solid rgba(255, 255, 255, .14);
      background: linear-gradient(160deg, rgba(66, 84, 106, .95), rgba(40, 52, 70, .95)); color: #fff; text-align: left; padding: 18px; overflow: hidden; cursor: pointer; }
    .tile:active { transform: scale(.98); }
    .tile .tt { display: block; font-size: 19px; font-weight: 700; line-height: 1.2; }
    .tile .ti { position: absolute; right: 16px; bottom: 26px; width: 64px; height: 64px; }
    .tile .ti svg { width: 100%; height: 100%; fill: #fff; }
    .tile .tb { position: absolute; left: 0; bottom: 0; max-width: 85%; background: linear-gradient(180deg, #f7d75c, #d6a72a); color: #1b1406;
      font-weight: 700; font-size: 13px; padding: 8px 18px 8px 14px; border-top-right-radius: 18px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .hint { color: rgba(255, 255, 255, .7); font-size: 13.5px; margin: 8px 0; }
    .hint a, .link { color: #f7d75c; cursor: pointer; background: none; border: 0; font-size: 14px; padding: 0; font-weight: 700; }
    .warn { color: #ffc857; font-size: 13px; }
    .limits { font-size: 13px; color: rgba(255, 255, 255, .75); background: rgba(13, 20, 32, .55); border: 1px solid rgba(255, 255, 255, .18); border-radius: 12px; padding: 9px 12px; margin-bottom: 10px; }
    .limits b { color: #fff; } .limits.bad { border-color: #ff6b6b; } .limits.bad b { color: #ff6b6b; }
    .busy { display: flex; gap: 8px; align-items: center; color: #f7d75c; font-weight: 700; margin-bottom: 10px; }
    .busy .grow { flex: 1; min-width: 0; } .btn.stop { min-height: 36px; padding: 6px 14px; border-color: #ff6b6b; color: #ff8a8a; }
    .spin { width: 14px; height: 14px; border: 2px solid currentColor; border-right-color: transparent; border-radius: 50%; animation: s .7s linear infinite; }
    @keyframes s { to { transform: rotate(360deg); } }
    ul { list-style: none; padding: 0; margin: 8px 0; display: grid; gap: 8px; }
    .row-btn { all: unset; display: block; width: 100%; box-sizing: border-box; background: linear-gradient(160deg, rgba(66, 84, 106, .9), rgba(40, 52, 70, .9));
      border: 1px solid rgba(255, 255, 255, .14); border-radius: 16px; padding: 12px 14px; cursor: pointer; }
    .row-btn b { display: block; color: #fff; font-size: 16px; } .row-btn small { color: rgba(255, 255, 255, .65); font-size: 13px; }
    .sets li.open > .row-btn { border-color: #f7d75c; }
    .chs { padding-left: 12px; }
    .stats { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin: 8px 0; }
    .stat { background: rgba(13, 20, 32, .55); border: 1px solid rgba(255, 255, 255, .18); border-radius: 12px; padding: 10px 12px; }
    .stat span { display: block; color: rgba(255, 255, 255, .7); font-size: 11.5px; text-transform: uppercase; letter-spacing: .06em; }
    .stat b { font-size: 20px; font-variant-numeric: tabular-nums; }
    .stat.good b { color: #3ddc84; } .stat.bad b { color: #ff6b6b; } .stat.warn b { color: #f7d75c; }
    .btn { background: rgba(13, 20, 32, .6); color: #fff; border: 1px solid rgba(255, 255, 255, .35); border-radius: 12px; padding: 11px 16px; font-size: 15px; font-weight: 700; min-height: 44px; cursor: pointer; }
    .btn.primary { background: linear-gradient(180deg, #f7d75c, #d9a92c); border-color: #e9c14a; color: #1b1406; } .btn:disabled { opacity: .45; }
    .row { display: flex; gap: 8px; flex-wrap: wrap; align-items: end; margin: 10px 0; }
    input[type=search], input[type=number], textarea { flex: 1; min-width: 0; width: 100%; background: rgba(13, 20, 32, .7); color: #fff; border: 1px solid rgba(255, 255, 255, .25); border-radius: 12px; padding: 11px; font-size: 15px; }
    textarea { font-family: ui-monospace, monospace; font-size: 11px; }
    label { display: flex; flex-direction: column; gap: 4px; color: rgba(255, 255, 255, .7); font-size: 13px; flex: 1; }
    label.chk { flex-direction: row; align-items: center; gap: 8px; color: #fff; margin: 8px 0; }
    .players li { display: flex; align-items: center; gap: 10px; background: linear-gradient(160deg, rgba(66, 84, 106, .9), rgba(40, 52, 70, .9)); border: 1px solid rgba(255, 255, 255, .12); border-radius: 14px; padding: 8px 10px; }
    .players .grow { flex: 1; min-width: 0; } .players b { display: block; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .players small { color: rgba(255, 255, 255, .65); font-size: 12px; }
    .badge { width: 36px; height: 40px; border-radius: 8px 8px 12px 12px; display: grid; place-items: center; font-weight: 800; color: #3a2a05; background: linear-gradient(160deg, #f9e7a0, #d9b24a); flex: none; }
    .badge.silver { background: linear-gradient(160deg, #f1f4f8, #aeb8c6); color: #2b3440; } .badge.bronze { background: linear-gradient(160deg, #f0c7a3, #b9784a); color: #3a1f0a; }
    .slot { width: 36px; color: rgba(255, 255, 255, .7); font-size: 12px; font-weight: 700; }
    .chem { color: #3ddc84; font-size: 10px; } .chem.c0 { color: #ff6b6b; }
    .price { background: rgba(13, 20, 32, .6); color: rgba(255, 255, 255, .75); border: 1px solid rgba(255, 255, 255, .25); border-radius: 9px; padding: 6px 9px; font-size: 13px; font-weight: 700; cursor: pointer; font-variant-numeric: tabular-nums; }
    .price.on { color: #f7d75c; border-color: #b8923a; }
    .lock { background: none; border: 0; font-size: 20px; cursor: pointer; opacity: .55; } .lock.on { opacity: 1; }
    .callout { border-radius: 12px; padding: 10px 12px; margin: 8px 0; border: 1px solid rgba(255, 255, 255, .2); background: rgba(13, 20, 32, .5); }
    .callout.good { color: #3ddc84; border-color: #3ddc8466; } .callout.bad { color: #ff6b6b; border-color: #ff6b6b66; }
    .checks li { display: flex; gap: 8px; font-size: 14px; } .checks em { margin-left: auto; font-style: normal; color: rgba(255, 255, 255, .7); }
    .checks li.ok span { color: #3ddc84; } .checks li.no { color: #ff6b6b; }
    h3 { font-size: 13px; text-transform: uppercase; letter-spacing: .06em; color: rgba(255, 255, 255, .7); margin: 18px 0 6px; } h3.ch { font-size: 19px; text-transform: none; letter-spacing: 0; color: #fff; }
    .log li { font-size: 12.5px; color: rgba(255, 255, 255, .7); } .log .good { color: #3ddc84; } .log .bad { color: #ff6b6b; } .log time { opacity: .7; margin-right: 6px; }
    .toast { position: fixed; left: 50%; bottom: calc(110px + env(safe-area-inset-bottom, 0px)); transform: translateX(-50%); background: #fff; color: #17243a; padding: 10px 16px; border-radius: 12px; font-weight: 700; max-width: 90vw; z-index: 5; pointer-events: none; }
    .toast.bad { background: #ff6b6b; color: #2b0000; }
    .dim { position: fixed; inset: 0; background: #000a; display: grid; place-items: center; padding: 16px; z-index: 4; }
    .dlg { background: #22324a; border: 1px solid rgba(255, 255, 255, .2); border-radius: 18px; padding: 18px; max-width: 420px; }
    `;
  }
})();
