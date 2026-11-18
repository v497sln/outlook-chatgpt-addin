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
    const base = { locked: {}, subs: [], searches: [], prices: {}, limits: { hour: 90, day: 300, search: 150 }, club: null, clubAt: 0, onlyUT: false, undercut: true, floorBE: true, log: [] };
    try { return Object.assign(base, JSON.parse(localStorage.getItem(KEY) || '{}')); } catch { return base; }
  })();
  S.limits.search = S.limits.search || 150; S.prices = S.prices || {}; S.searches = S.searches || [];
  const save = () => { try { localStorage.setItem(KEY, JSON.stringify(S)); } catch { /* storage full */ } };
  function log(msg, kind = 'info') {
    S.log.unshift({ at: Date.now(), msg, kind });
    S.log = S.log.slice(0, 80);
    save();
    if (ui.tab === 'settings') render();
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
  const loc = (keys, fb) => {
    try { for (const k of keys) { const t = services.Localization.localize(k); if (t && t !== k) return t; } } catch { /* ignore */ }
    return fb;
  };
  const nation = id => loc([`search.nationName.nation${id}`], `Nation #${id}`);
  const league = id => loc([`global.leagueFull.${id}`, `global.leagueFull.2027.league${id}`, `global.leagueFull.2026.league${id}`], `League #${id}`);
  const club = id => loc([`global.teamFull.${id}`, `global.teamFull.2027.team${id}`, `global.teamFull.2026.team${id}`], `Club #${id}`);
  const POS = ['GK', 'SW', 'RWB', 'RB', 'CB', 'CB', 'CB', 'LB', 'LWB', 'CDM', 'CDM', 'CDM', 'RM', 'CM', 'CM', 'CM', 'LM', 'CAM', 'CAM', 'CAM', 'CF', 'CF', 'CF', 'RW', 'ST', 'ST', 'ST', 'LW'];
  const pos = p => (typeof p === 'number' ? POS[p] || null : typeof p === 'string' ? p.toUpperCase() : null);
  const nameOf = i => { try { const s = i.getStaticData ? i.getStaticData() : i._staticData; if (s && s.name) return s.name; } catch { /* ignore */ } return `Player ${i.definitionId}`; };
  const items = {}; // eaId -> live EA item (needed to submit)

  function mapItem(i) {
    let ps = [pos(i.preferredPosition), ...(i.possiblePositions || []).map(pos)].filter(Boolean);
    ps = ps.filter((p, k) => ps.indexOf(p) === k);
    const rf = i.rareflag || 0;
    return {
      id: `ea${i.id}`, eaId: i.id, definitionId: i.definitionId, name: nameOf(i), rating: i.rating, positions: ps,
      nation: nation(i.nationId), league: league(i.leagueId), club: club(i.teamId),
      rarity: rf === 12 ? 'icon' : rf > 0 ? 'rare' : 'common', rare: rf > 0,
      untradeable: !!i.untradeable,
      loan: !!((i.isLimitedUse && i.isLimitedUse()) || i.loans > 0),
    };
  }

  const ready = () => { try { return !!(window.services && services.User && services.User.getUser() && services.Club && services.SBC); } catch { return false; } };

  async function loadClub(onProgress) {
    const seen = {}; const out = []; let offset = 0; const count = 91; let repeats = 0;
    for (let page = 0; page < 120; page++) {
      const c = new UTSearchCriteriaDTO();
      c.type = SearchType.PLAYER; c.count = count; c.offset = offset;
      const r = await obs(services.Club.search(c));
      const list = (r.response && r.response.items) || [];
      let fresh = 0;
      for (const i of list) {
        if (seen[i.id]) continue;
        seen[i.id] = 1; fresh++; items[i.id] = i;
        const m = mapItem(i);
        if (!m.loan) out.push(m);
      }
      onProgress && onProgress(out.length);
      if (!list.length || r.response.retrievedAll === true || list.length < count) break;
      if (!fresh && ++repeats >= 2) break;
      offset += list.length;
      await sleep(60);
    }
    S.club = out; S.clubAt = Date.now(); save();
    log(`Club loaded: ${out.length} players.`, 'good');
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
  const PRICE_TTL = 30 * 60000;
  let lastSearch = 0;
  let stopPricing = false;
  const priceOf = defId => { const c = S.prices[defId]; return c && Date.now() - c.at < PRICE_TTL ? c : null; };

  function searchCount() {
    const now = Date.now();
    S.searches = S.searches.filter(t => now - t < 3600000);
    return S.searches.length;
  }

  async function marketSearch(defId, maxBuy) {
    if (searchCount() >= S.limits.search) throw new Error(`Market search limit reached (${S.limits.search} per hour). Try again later.`);
    const gap = 2000 + Math.random() * 2000;
    const wait = lastSearch + gap - Date.now();
    if (wait > 0) await sleep(wait);
    lastSearch = Date.now();
    S.searches.push(lastSearch);
    const c = new UTSearchCriteriaDTO();
    c.type = SearchType.PLAYER;
    c.maskedDefId = defId;
    if (maxBuy) c.maxBuy = maxBuy;
    const r = await obs(services.Item.searchTransferMarket(c, 1));
    return ((r.response && r.response.items) || [])
      .map(i => ((i.getAuctionData && i.getAuctionData()) || i._auction || {}).buyNowPrice)
      .filter(v => v > 0);
  }

  async function marketPrice(defId, force) {
    const cached = !force && priceOf(defId);
    if (cached) return cached.p;
    let best = null; let maxBuy = null; let seen = 0;
    for (let round = 0; round < 5; round++) {
      const list = await marketSearch(defId, maxBuy);
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
    return best;
  }

  const kfmt = n => (n == null ? '—' : n >= 1e6 ? `${(n / 1e6).toFixed(2)}M` : n >= 1e4 ? `${(n / 1e3).toFixed(1)}k` : fmt(n));

  // ------------------------------------------------------------ submission limits
  function subCounts() {
    const now = Date.now();
    S.subs = S.subs.filter(t => now - t < 86400000);
    return { hour: S.subs.filter(t => now - t < 3600000).length, day: S.subs.length };
  }

  // ------------------------------------------------------------ UI
  const ui = { open: false, tab: 'sbc', busy: '', sets: null, openSet: null, chs: null, pick: null, result: null, seed: 1, q: '', trades: null, confirm: null };
  const h = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const fmt = n => Math.round(n || 0).toLocaleString();
  const ago = t => { const m = Math.round((Date.now() - t) / 60000); return m < 1 ? 'just now' : m < 60 ? `${m} min ago` : `${Math.round(m / 60)} h ago`; };
  const q = c => (c >= 75 ? 'gold' : c >= 65 ? 'silver' : 'bronze');

  const host = document.createElement('div');
  host.id = 'squadsage-host';
  host.style.cssText = 'position:fixed;inset:0;pointer-events:none;z-index:2147483600';
  const root = host.attachShadow({ mode: 'open' });
  root.innerHTML = `<style>${CSS()}</style>
    <button class="fab" data-a="toggle" aria-label="Open SquadSage"><span>★</span></button>
    <section class="sheet" hidden>
      <header><b>Squad<i>Sage</i></b><button class="x" data-a="toggle" aria-label="Close">×</button></header>
      <nav>${[['sbc', 'SBC'], ['club', 'Club'], ['trades', 'Trades'], ['settings', 'Safety']].map(([k, l]) => `<button data-a="tab" data-t="${k}">${l}</button>`).join('')}</nav>
      <div class="body"></div>
      <div class="toast" hidden></div>
    </section>`;
  const $ = s => root.querySelector(s);

  function toast(msg, bad) {
    const t = $('.toast');
    t.textContent = msg; t.className = `toast ${bad ? 'bad' : ''}`; t.hidden = false;
    clearTimeout(toast.t); toast.t = setTimeout(() => (t.hidden = true), 3500);
  }

  function render() {
    $('.sheet').hidden = !ui.open;
    $('.fab').hidden = ui.open;
    root.querySelectorAll('nav button').forEach(b => b.classList.toggle('on', b.dataset.t === ui.tab));
    if (!ui.open) return;
    const body = $('.body');
    const top = body.scrollTop;
    body.innerHTML = (ui.busy ? `<div class="busy"><span class="spin"></span>${h(ui.busy)}</div>` : '') + VIEW[ui.tab]() + (ui.confirm ? confirmBox() : '');
    body.scrollTop = top;
  }

  function confirmBox() {
    return `<div class="dim"><div class="dlg"><p>${h(ui.confirm.msg)}</p><div class="row"><button class="btn primary" data-a="yes">${h(ui.confirm.yes)}</button><button class="btn" data-a="no">Cancel</button></div></div></div>`;
  }
  const ask = (msg, yes, fn) => { ui.confirm = { msg, yes, fn }; render(); };

  const stat = (l, v, cls = '') => `<div class="stat ${cls}"><span>${h(l)}</span><b>${v}</b></div>`;

  const VIEW = {
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
        <div class="row"><button class="btn primary" data-a="priceBatch" ${ui.busy ? 'disabled' : ''}>Price my top tradeables</button>${ui.busy ? '<button class="btn" data-a="stopPrice">Stop</button>' : ''}</div>
        <div class="row"><input type="search" placeholder="Search name, league, nation…" data-in="q" value="${h(ui.q)}"><button class="btn" data-a="loadClub">Reload</button></div>
        <ul class="players">${list.map(p => `<li><span class="badge ${q(p.rating)}">${p.rating}</span><div class="grow"><b>${h(p.name)}</b><small>${h(p.positions.join('/'))} · ${h(p.league)}${p.untradeable ? ' · UT' : ''}</small></div>
          ${p.untradeable ? '' : `<button class="price ${priceOf(p.definitionId) ? 'on' : ''}" data-a="price1" data-d="${p.definitionId}">${priceOf(p.definitionId) ? kfmt(priceOf(p.definitionId).p) : 'price?'}</button>`}
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
        <div class="row"><button class="btn primary" data-a="smartRelist" ${expired && !ui.busy ? '' : 'disabled'}>Smart relist ${expired}</button><button class="btn" data-a="relist" ${expired && !ui.busy ? '' : 'disabled'}>Relist at same price</button><button class="btn" data-a="trades">Refresh</button></div>
        <label class="chk"><input type="checkbox" data-in="undercut" ${S.undercut ? 'checked' : ''}> Smart relist: list one price step under the cheapest</label>
        <label class="chk"><input type="checkbox" data-in="floorBE" ${S.floorBE ? 'checked' : ''}> Never relist below what I paid (after tax)</label>
        ${expired ? `<h3>Expired</h3><ul class="players">${t.filter(i => i.state === 'expired').map(i => `<li><span class="badge ${q(i.rating)}">${i.rating}</span><div class="grow"><b>${h(i.name)}</b><small>listed ${fmt(i.bin)}${i.paid ? ` · paid ${fmt(i.paid)}` : ''}</small></div>
          <button class="price ${priceOf(i.defId) ? 'on' : ''}" data-a="price1" data-d="${i.defId}">${priceOf(i.defId) ? kfmt(priceOf(i.defId).p) : 'price?'}</button></li>`).join('')}</ul>` : ''}`;
    },
    settings() {
      return `<h3>SBC submission limits</h3>
        <p class="hint">EA soft-bans accounts that submit SBCs too fast. SquadSage stops submitting when either limit is reached. It only counts submissions it made itself.</p>
        <div class="row"><label>Per hour<input type="number" min="1" data-in="lh" value="${S.limits.hour}"></label><label>Per day<input type="number" min="1" data-in="ld" value="${S.limits.day}"></label></div>
        <h3>Market price checks</h3>
        <p class="hint">Searches this hour: <b>${searchCount()}/${S.limits.search}</b>. Searches are spaced 2–4 seconds apart; prices are remembered for 30 minutes.</p>
        <div class="row"><label>Searches per hour<input type="number" min="10" data-in="ls" value="${S.limits.search}"></label></div>
        <label class="chk"><input type="checkbox" data-in="ut" ${S.onlyUT ? 'checked' : ''}> Solve with untradeable players only</label>
        <p class="warn">Automating the EA Web App is against EA's terms. EA can restrict or ban accounts for it.</p>
        <h3>Activity</h3>
        <ul class="log">${S.log.slice(0, 30).map(l => `<li class="${l.kind}"><time>${new Date(l.at).toLocaleTimeString()}</time> ${h(l.msg)}</li>`).join('') || '<li class="hint">Nothing yet.</li>'}</ul>
        <div class="row"><button class="btn" data-a="copylog">Copy log</button><button class="btn" data-a="reset">Clear SquadSage data</button></div>`;
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
  async function busy(label, fn) {
    ui.busy = label; render();
    try { return await fn(); } catch (e) { log(`${label.replace('…', '')} failed: ${e.message}`, 'bad'); toast(e.message, true); return null; } finally { ui.busy = ''; render(); }
  }

  async function loadSets() {
    const r = await obs(services.SBC.requestSets());
    const sets = r.response.sets || (r.response.categories || []).reduce((a, c) => a.concat(c.sets || []), []);
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

  const A = {
    toggle() {
      ui.open = !ui.open;
      render();
      if (ui.open && !ui.sets && ready()) busy('Loading your SBCs…', loadSets);
    },
    tab(el) {
      ui.tab = el.dataset.t; ui.confirm = null; render();
      if (ui.tab === 'trades') A.trades();
      if (ui.tab === 'club' && !S.club) A.loadClub();
    },
    loadClub() { busy('Loading your club…', () => loadClub(n => { ui.busy = `Loading your club… ${n} players`; render(); })); },
    async set(el) {
      const s = ui.sets[+el.dataset.i];
      if (ui.openSet === s) { ui.openSet = null; render(); return; }
      ui.openSet = s; ui.chs = null; render();
      await busy('Loading challenges…', async () => {
        const r = await obs(services.SBC.requestChallengesForSet(s));
        ui.chs = (r.response.challenges || []).filter(c => c.status !== 'COMPLETED' && !(c.isCompleted && c.isCompleted()));
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
    trades() {
      busy('Loading your transfer list…', async () => {
        const r = await obs(services.Item.requestTransferItems());
        ui.trades = ((r.response && r.response.items) || []).map(i => {
          const a = (i.getAuctionData && i.getAuctionData()) || i._auction || {};
          return { id: i.id, defId: i.definitionId, name: nameOf(i), rating: i.rating, state: a.tradeState || null, bin: a.buyNowPrice || 0, bid: a.currentBid || 0, paid: i.lastSalePrice || 0, raw: i };
        });
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
      await busy('Checking market price…', async () => {
        const p = await marketPrice(d, true);
        toast(p ? `Lowest Buy Now: ${fmt(p)}` : 'None on the market right now');
      });
    },
    stopPrice() { stopPricing = true; },
    async priceBatch() {
      const todo = S.club.filter(p => !p.untradeable && !priceOf(p.definitionId)).sort((a, b) => b.rating - a.rating);
      const seen = new Set(); const list = todo.filter(p => !seen.has(p.definitionId) && seen.add(p.definitionId)).slice(0, 15);
      if (!list.length) return toast('All tradeables are priced');
      stopPricing = false;
      await busy('Checking prices…', async () => {
        for (let i = 0; i < list.length && !stopPricing; i++) {
          ui.busy = `Checking prices… ${i + 1}/${list.length} (${list[i].name})`; render();
          await marketPrice(list[i].definitionId);
        }
      });
      log(`Priced ${list.length} players.`, 'good');
    },
    smartRelist() {
      const ex = (ui.trades || []).filter(i => i.state === 'expired');
      ask(`Check live prices for ${ex.length} expired items and relist each at the market price? This uses up to ${ex.length * 3} market searches.`, 'Smart relist', () => busy('Smart relisting…', async () => {
        stopPricing = false;
        let done = 0, kept = 0, none = 0;
        for (const i of ex) {
          if (stopPricing) break;
          ui.busy = `Smart relisting… ${done + kept + none + 1}/${ex.length} (${i.name})`; render();
          const m = await marketPrice(i.defId);
          if (!m) { none++; continue; }
          let bin = S.undercut ? E.roundToMarket(m - E.priceStep(m), 'down') : m;
          if (S.floorBE && i.paid) bin = Math.max(bin, E.roundToMarket(Math.ceil(i.paid / 0.95), 'up'));
          bin = Math.max(200, bin);
          const start = E.roundToMarket(bin - E.priceStep(bin), 'down');
          await obs(services.Item.list(i.raw, Math.max(150, start), bin, 3600));
          done++;
          await sleep(500 + Math.random() * 600);
        }
        log(`Smart relist: ${done} relisted${none ? `, ${none} had no market price (left as is)` : ''}.`, 'good');
        toast(`Relisted ${done}`);
        A.trades();
      }));
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
    if (k === 'ls') S.limits.search = Math.max(10, +e.target.value || 150);
    if (k === 'undercut') S.undercut = e.target.checked;
    if (k === 'floorBE') S.floorBE = e.target.checked;
    save();
  });

  // Mount once EA's app is up (the login page has no `services`).
  (function mount() {
    if (!document.body) return setTimeout(mount, 300);
    if (!document.getElementById('squadsage-host')) document.body.appendChild(host);
    host.style.display = ready() ? '' : 'none';
    if (!ready()) return setTimeout(mount, 1000);
    log('Connected to the EA Web App.', 'good');
    // Warm up quietly so the panel opens instantly.
    setTimeout(() => { loadSets().catch(() => {}); }, 1500);
  })();

  function CSS() {
    return `
    :host { all: initial; }
    * { box-sizing: border-box; font-family: system-ui, -apple-system, Roboto, sans-serif; }
    [hidden] { display: none !important; }
    .fab { pointer-events: auto; position: fixed; right: 14px; bottom: calc(96px + env(safe-area-inset-bottom, 0px)); width: 54px; height: 54px; border-radius: 50%;
      border: 0; background: #12c2a0; color: #062a23; font-size: 24px; box-shadow: 0 6px 18px #0009; cursor: pointer; }
    .sheet { pointer-events: auto; position: fixed; inset: 0; background: #0d1420f2; color: #e9eef6; display: flex; flex-direction: column; backdrop-filter: blur(6px); }
    header { display: flex; justify-content: space-between; align-items: center; padding: calc(12px + env(safe-area-inset-top, 0px)) 16px 8px; }
    header b { font-size: 19px; font-weight: 600; } header i { font-style: normal; color: #12c2a0; }
    .x { background: none; border: 0; color: #e9eef6; font-size: 30px; width: 44px; height: 44px; cursor: pointer; }
    nav { display: flex; gap: 6px; padding: 0 12px 10px; border-bottom: 1px solid #26344a; }
    nav button { flex: 1; background: #141d2c; color: #8b9ab2; border: 1px solid #26344a; border-radius: 10px; padding: 9px 0; font-size: 14px; font-weight: 600; cursor: pointer; }
    nav button.on { background: #12c2a0; color: #062a23; border-color: #12c2a0; }
    .body { flex: 1; overflow: auto; padding: 12px 16px calc(24px + env(safe-area-inset-bottom, 0px)); position: relative; }
    .hint { color: #8b9ab2; font-size: 13px; margin: 8px 0; } .hint a, .link { color: #12c2a0; cursor: pointer; background: none; border: 0; font-size: 14px; padding: 0; font-weight: 600; }
    .warn { color: #ffc857; font-size: 13px; }
    .limits { font-size: 13px; color: #8b9ab2; background: #141d2c; border: 1px solid #26344a; border-radius: 10px; padding: 8px 10px; margin-bottom: 8px; }
    .limits b { color: #e9eef6; } .limits.bad { border-color: #ff6b6b; } .limits.bad b { color: #ff6b6b; }
    .busy { display: flex; gap: 8px; align-items: center; color: #12c2a0; font-weight: 600; margin-bottom: 10px; }
    .spin { width: 14px; height: 14px; border: 2px solid currentColor; border-right-color: transparent; border-radius: 50%; animation: s .7s linear infinite; }
    @keyframes s { to { transform: rotate(360deg); } }
    ul { list-style: none; padding: 0; margin: 8px 0; display: grid; gap: 6px; }
    .row-btn { all: unset; display: block; width: 100%; box-sizing: border-box; background: #141d2c; border: 1px solid #26344a; border-radius: 12px; padding: 10px 12px; cursor: pointer; }
    .row-btn b { display: block; color: #e9eef6; } .row-btn small { color: #8b9ab2; font-size: 12.5px; }
    .sets li.open > .row-btn { border-color: #12c2a0; }
    .chs { padding-left: 14px; }
    .stats { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin: 8px 0; }
    .stat { background: #141d2c; border: 1px solid #26344a; border-radius: 12px; padding: 9px 11px; }
    .stat span { display: block; color: #8b9ab2; font-size: 11px; text-transform: uppercase; letter-spacing: .05em; }
    .stat b { font-size: 19px; font-variant-numeric: tabular-nums; }
    .stat.good b { color: #3ddc84; } .stat.bad b { color: #ff6b6b; } .stat.warn b { color: #ffc857; }
    .btn { background: #141d2c; color: #e9eef6; border: 1px solid #26344a; border-radius: 12px; padding: 10px 14px; font-size: 14px; font-weight: 600; min-height: 42px; cursor: pointer; }
    .btn.primary { background: #12c2a0; border-color: #12c2a0; color: #062a23; } .btn:disabled { opacity: .45; }
    .row { display: flex; gap: 8px; flex-wrap: wrap; align-items: end; margin: 10px 0; }
    input[type=search], input[type=number] { flex: 1; min-width: 0; background: #0d1420; color: #e9eef6; border: 1px solid #26344a; border-radius: 10px; padding: 10px; font-size: 15px; }
    label { display: flex; flex-direction: column; gap: 4px; color: #8b9ab2; font-size: 13px; flex: 1; }
    label.chk { flex-direction: row; align-items: center; gap: 8px; color: #e9eef6; margin: 8px 0; }
    .players li { display: flex; align-items: center; gap: 10px; background: #141d2c; border: 1px solid #26344a; border-radius: 12px; padding: 7px 10px; }
    .players .grow { flex: 1; min-width: 0; } .players b { display: block; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .players small { color: #8b9ab2; font-size: 12px; }
    .badge { width: 34px; height: 34px; border-radius: 9px; display: grid; place-items: center; font-weight: 800; color: #1b1406; background: #e6c35c; flex: none; }
    .badge.silver { background: #b9c2cf; } .badge.bronze { background: #c98a5a; }
    .slot { width: 36px; color: #8b9ab2; font-size: 12px; font-weight: 700; }
    .chem { color: #3ddc84; font-size: 10px; } .chem.c0 { color: #ff6b6b; }
    .price { background: #0d1420; color: #8b9ab2; border: 1px solid #26344a; border-radius: 8px; padding: 5px 8px; font-size: 12.5px; font-weight: 700; cursor: pointer; font-variant-numeric: tabular-nums; }
    .price.on { color: #ffd76a; border-color: #5a4a1b; }
    .lock { background: none; border: 0; font-size: 20px; cursor: pointer; opacity: .55; } .lock.on { opacity: 1; }
    .callout { border-radius: 12px; padding: 10px 12px; margin: 8px 0; border: 1px solid #26344a; }
    .callout.good { color: #3ddc84; border-color: #3ddc8466; } .callout.bad { color: #ff6b6b; border-color: #ff6b6b66; }
    .checks li { display: flex; gap: 8px; font-size: 14px; } .checks em { margin-left: auto; font-style: normal; color: #8b9ab2; }
    .checks li.ok span { color: #3ddc84; } .checks li.no { color: #ff6b6b; }
    h3 { font-size: 13px; text-transform: uppercase; letter-spacing: .06em; color: #8b9ab2; margin: 16px 0 6px; } h3.ch { font-size: 18px; text-transform: none; letter-spacing: 0; color: #e9eef6; }
    .log li { font-size: 12.5px; color: #8b9ab2; } .log .good { color: #3ddc84; } .log .bad { color: #ff6b6b; } .log time { opacity: .7; margin-right: 6px; }
    .toast { position: fixed; left: 50%; bottom: calc(24px + env(safe-area-inset-bottom, 0px)); transform: translateX(-50%); background: #e9eef6; color: #0d1420; padding: 10px 16px; border-radius: 12px; font-weight: 600; max-width: 90vw; }
    .toast.bad { background: #ff6b6b; color: #2b0000; }
    .dim { position: fixed; inset: 0; background: #000a; display: grid; place-items: center; padding: 16px; }
    .dlg { background: #141d2c; border: 1px solid #26344a; border-radius: 16px; padding: 16px; max-width: 420px; }
    `;
  }
})();
