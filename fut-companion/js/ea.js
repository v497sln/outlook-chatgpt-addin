// EA Web App bridge (Android app only).
//
// How it works: the official EA FC Web App is opened inside the app in an
// in-app browser. YOU log in on EA's own page — SquadSage never sees or stores
// your password. Once you're logged in, SquadSage runs small scripts inside
// that page which call the Web App's own internal services (the same ones its
// buttons use) to read your club / transfer list and to submit SBCs.
//
// RISK: this is automation of the Web App and is against EA's terms. EA can
// restrict or ban accounts for it. Requests are spaced out, but there is no
// way to make this risk-free. The user opts in explicitly.
//
// The Web App's internals are undocumented and change between game years, so
// every call is wrapped and every failure is logged in plain words.

export const WEBAPP_URL = 'https://www.ea.com/ea-sports-fc/ultimate-team/web-app/';

export const state = { status: 'off', persona: '', coins: null, task: '', log: [] }; // off | opening | waiting | ready | busy
let ref = null;
let pollTimer = null;
let onChange = () => {};
let onReady = () => {};
let showTimer = null;
let jobSeq = 0;

export function setListener(fn) { onChange = fn; }
export function setReadyHandler(fn) { onReady = fn; }
export function available() { return !!(window.cordova && window.cordova.InAppBrowser); }

function log(msg, kind = 'info') {
  state.log.unshift({ at: Date.now(), msg, kind });
  state.log = state.log.slice(0, 60);
  onChange();
}
function setStatus(s) { state.status = s; onChange(); }

function exec(code) {
  return new Promise((resolve, reject) => {
    if (!ref) return reject(new Error('EA Web App is not open. Tap Connect first.'));
    try { ref.executeScript({ code }, r => resolve(r && r[0])); } catch (e) { reject(e); }
  });
}

const sleep = ms => new Promise(r => setTimeout(r, ms));

// Helper installed into the EA page once per page load.
const HELPER = `(function () {
  if (window.__ss) return 'ok';
  var ss = window.__ss = { res: {}, cmd: null, items: {} };
  ss.wait = function (ms) { return new Promise(function (r) { setTimeout(r, ms); }); };
  ss.human = function () { return ss.wait(350 + Math.random() * 450); };
  ss.obs = function (o) {
    return new Promise(function (res, rej) {
      if (!o || !o.observe) return rej(new Error('EA service returned nothing'));
      o.observe(window, function (_o, r) {
        try { _o.unobserve(window); } catch (e) {}
        if (r && r.success) res(r);
        else rej(new Error('EA request failed (' + ((r && (r.status || (r.error && r.error.code))) || 'unknown') + ')'));
      });
    });
  };
  ss.run = function (id, fn) {
    ss.res[id] = { done: false };
    Promise.resolve().then(fn).then(function (v) { ss.res[id] = { done: true, ok: true, value: v }; })
      .catch(function (e) { ss.res[id] = { done: true, ok: false, error: String((e && e.message) || e) }; });
  };
  ss.loc = function (keys, fallback) {
    try {
      for (var i = 0; i < keys.length; i++) {
        var t = services.Localization.localize(keys[i]);
        if (t && t !== keys[i]) return t;
      }
    } catch (e) {}
    return fallback;
  };
  ss.nation = function (id) { return ss.loc(['search.nationName.nation' + id], 'Nation #' + id); };
  ss.league = function (id) { return ss.loc(['global.leagueFull.' + id, 'global.leagueFull.2027.league' + id, 'global.leagueFull.2026.league' + id], 'League #' + id); };
  ss.club = function (id) { return ss.loc(['global.teamFull.' + id, 'global.teamFull.2027.team' + id, 'global.teamFull.2026.team' + id], 'Club #' + id); };
  var POS = ['GK','SW','RWB','RB','CB','CB','CB','LB','LWB','CDM','CDM','CDM','RM','CM','CM','CM','LM','CAM','CAM','CAM','CF','CF','CF','RW','ST','ST','ST','LW'];
  ss.pos = function (p) { if (typeof p === 'number') return POS[p] || null; if (typeof p === 'string') return p.toUpperCase(); return null; };
  ss.name = function (i) {
    try { var s = i.getStaticData ? i.getStaticData() : i._staticData; if (s && s.name) return s.name; } catch (e) {}
    return 'Player ' + i.definitionId;
  };
  ss.mapItem = function (i) {
    var pos = [ss.pos(i.preferredPosition)].concat((i.possiblePositions || []).map(ss.pos)).filter(Boolean);
    pos = pos.filter(function (p, k) { return pos.indexOf(p) === k; });
    var rf = i.rareflag || 0;
    return {
      eaId: i.id, definitionId: i.definitionId, name: ss.name(i), rating: i.rating, positions: pos,
      nation: ss.nation(i.nationId), league: ss.league(i.leagueId), club: ss.club(i.teamId),
      rarity: rf === 12 ? 'icon' : rf > 0 ? 'rare' : 'common', rare: rf > 0,
      untradeable: !!i.untradeable,
      loan: !!((i.isLimitedUse && i.isLimitedUse()) || i.loans > 0)
    };
  };
  // Floating bar on EA's page so you can always get back to SquadSage.
  try {
    var bar = document.createElement('div');
    bar.id = '__ss_bar';
    bar.style.cssText = 'position:fixed;left:50%;bottom:16px;transform:translateX(-50%);z-index:2147483647;background:#12c2a0;color:#062a23;font:600 15px system-ui;padding:12px 18px;border-radius:24px;box-shadow:0 4px 16px #0008;cursor:pointer';
    bar.textContent = '← Back to SquadSage';
    bar.onclick = function () { ss.cmd = 'back'; };
    document.body.appendChild(bar);
  } catch (e) {}
  return 'ok';
})()`;

// Runs an async function body inside the EA page and waits for its result.
async function job(body, timeoutMs = 120000, onProgress = null) {
  await exec(HELPER);
  const id = `j${++jobSeq}`;
  await exec(`window.__ss.run(${JSON.stringify(id)}, async function () { ${body} }); 'started'`);
  const end = Date.now() + timeoutMs;
  while (Date.now() < end) {
    await sleep(600);
    const raw = await exec(`JSON.stringify({ r: window.__ss && window.__ss.res[${JSON.stringify(id)}] || null, p: window.__ss && window.__ss.progress != null ? { n: window.__ss.progress, batch: (window.__ss.partial || []).splice(0) } : null })`);
    const box = raw ? JSON.parse(raw) : {};
    const r = box.r;
    if (onProgress && box.p != null) onProgress(box.p);
    if (r && r.done) {
      exec(`delete window.__ss.res[${JSON.stringify(id)}]`).catch(() => {});
      if (!r.ok) throw new Error(r.error);
      return r.value;
    }
  }
  throw new Error('EA Web App did not answer in time');
}

// ---------------------------------------------------------------- connection
// hidden=true: reconnect in the background using EA's saved login. The EA page
// only appears if EA asks you to log in again.
export function connect({ hidden = false } = {}) {
  if (!available()) { log('EA connection only works in the Android app.', 'bad'); return; }
  if (ref) { if (!hidden) ref.show(); return; }
  setStatus('opening');
  ref = window.cordova.InAppBrowser.open(WEBAPP_URL, '_blank',
    `location=no,toolbar=yes,closebuttoncaption=Back,hidenavigationbuttons=yes,zoom=no,clearcache=no,clearsessioncache=no,hidden=${hidden ? 'yes' : 'no'}`);
  clearTimeout(showTimer);
  if (hidden) {
    showTimer = setTimeout(() => {
      if (ref && state.status === 'waiting') { log('EA needs you to log in again.'); ref.show(); }
    }, 20000);
  }
  ref.addEventListener('loadstop', () => { exec(HELPER).catch(() => {}); });
  ref.addEventListener('exit', () => { ref = null; clearInterval(pollTimer); setStatus('off'); log('EA Web App closed.'); });
  ref.addEventListener('loaderror', e => log(`Could not load EA Web App: ${e.message || 'network error'}`, 'bad'));
  setStatus('waiting');
  log(hidden ? 'Reconnecting to EA…' : 'Log in on EA\'s page. SquadSage never sees your password.');
  clearInterval(pollTimer);
  pollTimer = setInterval(checkLogin, 2000);
}

async function checkLogin() {
  try {
    const raw = await exec(`(function(){ try {
      var cmd = window.__ss && window.__ss.cmd; if (window.__ss) window.__ss.cmd = null;
      var u = window.services && services.User && services.User.getUser && services.User.getUser();
      var name = u ? (u.personaName || (u.getSelectedPersona && u.getSelectedPersona() && u.getSelectedPersona().name) || 'connected') : '';
      var ready = !!(name && window.services && services.Club && services.SBC);
      return JSON.stringify({ name: name, ready: ready, cmd: cmd });
    } catch (e) { return JSON.stringify({ error: String(e) }); } })()`);
    const r = raw ? JSON.parse(raw) : {};
    if (r.cmd === 'back' && ref) ref.hide();
    if (r.ready && state.status === 'waiting') {
      state.persona = r.name;
      clearTimeout(showTimer);
      if (ref) ref.hide();
      setStatus('ready');
      log(`Connected as ${r.name}.`, 'good');
      Promise.resolve().then(onReady).catch(() => {});
    }
  } catch { /* page still loading */ }
}

export function show() { if (ref) ref.show(); else connect(); }
export function disconnect() { if (ref) ref.close(); ref = null; clearInterval(pollTimer); clearTimeout(showTimer); state.persona = ''; setStatus('off'); }

// EA requests run one at a time, in the order they were asked for.
let queue = Promise.resolve();
let pending = 0;
function guarded(label, fn) {
  if (state.status !== 'ready' && state.status !== 'busy') return Promise.reject(new Error('Log in with EA first.'));
  pending++;
  const run = queue.then(async () => {
    setStatus('busy');
    state.task = `${label}…`;
    log(`${label}…`);
    try {
      return await fn();
    } catch (e) {
      log(`${label} failed: ${e.message}`, 'bad');
      throw e;
    } finally {
      pending--;
      state.task = '';
      if (state.status === 'busy' && pending === 0) setStatus('ready'); else onChange();
    }
  });
  queue = run.catch(() => {});
  return run;
}

// ---------------------------------------------------------------- account
export function readAccount() {
  return guarded('Loading your account', async () => {
    const a = await job(`
      var u = services.User.getUser(), c = null, name = '';
      // The coin counter in EA's own header is the most reliable source.
      try { var el = document.querySelector('.view-navbar-currency-coins'); if (el) c = parseInt(el.textContent.replace(/[^0-9]/g, ''), 10); } catch (e) {}
      if (c == null || isNaN(c)) try { c = u.coins.amount; } catch (e) {}
      if (c == null || isNaN(c)) try { c = (u.currencies || []).filter(function (x) { return /coin/i.test(x.name); })[0].amount; } catch (e) {}
      try { name = u.personaName || ''; } catch (e) {}
      if (!name) try { name = u.getSelectedPersona().name; } catch (e) {}
      if (!name) try { name = u.selectedPersona.name; } catch (e) {}
      if (!name) try { name = document.querySelector('.view-navbar-clubinfo-name').textContent.trim(); } catch (e) {}
      return { coins: (c == null || isNaN(c)) ? null : c, name: name };`, 30000);
    if (a.coins != null) state.coins = a.coins;
    if (a.name) state.persona = a.name;
    return a;
  });
}

// ---------------------------------------------------------------- club
// Streams players back in batches (onBatch) while it pages through the club.
// Stops when EA starts returning players it has already sent, so it can't loop.
export function importClub(onBatch = () => {}) {
  return guarded('Loading your club', async () => {
    let total = 0;
    const res = await job(`
      var ss = window.__ss, seen = {}, offset = 0, count = 91, n = 0, pages = 0, repeats = 0;
      ss.partial = []; ss.progress = 0;
      for (var page = 0; page < 120; page++) {
        var c = new UTSearchCriteriaDTO();
        c.type = SearchType.PLAYER; c.count = count; c.offset = offset;
        var r = await ss.obs(services.Club.search(c));
        pages++;
        var items = (r.response && r.response.items) || [], fresh = 0;
        items.forEach(function (i) {
          if (seen[i.id]) return;
          seen[i.id] = 1; fresh++; n++;
          ss.items[i.id] = i;
          ss.partial.push(ss.mapItem(i));
        });
        ss.progress = n;
        if (!items.length || r.response.retrievedAll === true || items.length < count) break;
        if (!fresh && ++repeats >= 2) break;
        offset += items.length;
        await ss.wait(60);
      }
      await ss.wait(700); // let the last batch be collected
      ss.progress = null;
      return { n: n, pages: pages, repeats: repeats };`, 300000, p => {
      total = p.n;
      state.task = `Loading your club… ${p.n} players`;
      if (p.batch && p.batch.length) onBatch(p.batch.filter(x => !x.loan));
      onChange();
    });
    log(`Club loaded: ${res.n} players in ${res.pages} requests${res.repeats ? ' (EA repeated players; stopped)' : ''}.`, 'good');
    return res.n || total;
  });
}

// ---------------------------------------------------------------- transfer list
export function readTransferList() {
  return guarded('Reading your transfer list', async () => {
    const items = await job(`
      var ss = window.__ss;
      var r = await ss.obs(services.Item.requestTransferItems());
      var items = (r.response && r.response.items) || [];
      return items.map(function (i) {
        var a = (i.getAuctionData && i.getAuctionData()) || i._auction || {};
        return { eaId: i.id, name: ss.name(i), rating: i.rating, tradeState: a.tradeState || null,
          bin: a.buyNowPrice || 0, bid: a.currentBid || 0, paid: i.lastSalePrice || 0 };
      });`);
    log(`Transfer list: ${items.length} items.`, 'good');
    return items;
  });
}

export function relistExpired() {
  return guarded('Relisting expired items', async () => {
    await job(`await window.__ss.obs(services.Item.relistExpiredAuctions()); return true;`);
    log('Expired items relisted at their previous prices.', 'good');
    return true;
  });
}

// ---------------------------------------------------------------- SBCs
// Maps EA eligibility requirements to SquadSage requirements, inside the page
// so ids can be turned into the same names the club import uses.
export function loadSbcs() {
  return guarded('Loading SBCs', async () => {
    const sets = await job(`
      var ss = window.__ss;
      var r = await ss.obs(services.SBC.requestSets());
      var sets = (r.response && (r.response.sets || (r.response.categories || []).reduce(function (a, c) { return a.concat(c.sets || []); }, []))) || [];
      return sets.filter(function (s) { return !s.isComplete || !s.isComplete(); }).slice(0, 80).map(function (s) {
        return { id: s.id, name: s.name, challenges: s.challengesCount || 0, done: s.challengesCompletedCount || 0 };
      });`);
    log(`Found ${sets.length} open SBC sets.`, 'good');
    return sets;
  });
}

export function loadChallenges(setId) {
  return guarded('Loading challenges', async () => {
    const list = await job(`
      var ss = window.__ss;
      var r = await ss.obs(services.SBC.requestSets());
      var sets = r.response.sets || (r.response.categories || []).reduce(function (a, c) { return a.concat(c.sets || []); }, []);
      var set = sets.find(function (s) { return s.id === ${Number(setId)}; });
      if (!set) throw new Error('SBC set not found');
      var cr = await ss.obs(services.SBC.requestChallengesForSet(set));
      var chs = (cr.response && cr.response.challenges) || [];
      var KEYS = window.SBCEligibilityKey || {}, SCOPES = window.SBCEligibilityScope || {};
      function keyName(k) { for (var n in KEYS) if (KEYS[n] === k) return n; return String(k); }
      function scopeName(s) { for (var n in SCOPES) if (SCOPES[n] === s) return n; return String(s); }
      var out = [];
      for (var c = 0; c < chs.length; c++) {
        var ch = chs[c];
        if (ch.status === 'COMPLETED' || (ch.isCompleted && ch.isCompleted())) continue;
        try { await ss.obs(services.SBC.loadChallenge(ch)); } catch (e) {}
        var reqs = (ch.eligibilityRequirements || []).map(function (e) {
          var pairs = [];
          try {
            var col = e.kvPairs && (e.kvPairs._collection || e.kvPairs);
            if (col && typeof col.forEach === 'function' && !Array.isArray(col)) col.forEach(function (v, k) { pairs.push([keyName(+k), v]); });
            else for (var k in col) pairs.push([keyName(+k), col[k]]);
          } catch (x) {}
          var key = pairs.length ? pairs[0][0] : keyName(e.key);
          var vals = pairs.length ? [].concat(pairs[0][1]) : [];
          var names = vals.map(function (v) {
            if (/NATION_ID/.test(key)) return ss.nation(v);
            if (/LEAGUE_ID/.test(key)) return ss.league(v);
            if (/CLUB_ID/.test(key)) return ss.club(v);
            return v;
          });
          return { key: key, scope: scopeName(e.scope), count: e.count, values: names, raw: vals };
        });
        var slots = [], bricks = 0;
        try {
          var sq = ch.squad;
          var field = (sq.getFieldPlayers && sq.getFieldPlayers()) || [];
          field.forEach(function (s) {
            var p = s.getPosition ? s.getPosition() : s.position;
            var id = p && (p.typeName || p.id);
            slots.push(ss.pos(typeof id === 'number' ? id : String(id || '').replace(/^[RL](?=(CB|CM|DM|AM|S|F)$)/, '').replace(/^DM$/, 'CDM').replace(/^AM$/, 'CAM').replace(/^S$/, 'ST').replace(/^F$/, 'CF')));
            if (s.isBrick && s.isBrick()) bricks++;
          });
        } catch (x) {}
        out.push({ id: ch.id, name: ch.name, setId: ${Number(setId)}, reqs: reqs, slots: slots, bricks: bricks });
        await ss.human();
      }
      return out;`, 180000);
    log(`Loaded ${list.length} challenges.`, 'good');
    return list;
  });
}

// EA requirement -> SquadSage requirement (null if unsupported).
export function translateReq(r) {
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
    case 'PLAYER_RARITY': case 'PLAYER_RARITY_GROUP':
      return Number(v) >= 1 ? { t: 'count', attr: 'rare', value: true, op, n } : null;
    case 'PLAYER_QUALITY': case 'PLAYER_LEVEL': {
      const q = { 1: 'bronze', 2: 'silver', 3: 'gold' }[Number(v)];
      if (!q) return null;
      if (r.scope === 'EXACT' || op === 'exact') return { t: 'count', attr: 'quality', value: q, op: 'exact', n: 11 };
      return { t: 'ovr', min: op === 'min' ? { bronze: 0, silver: 65, gold: 75 }[q] : null, max: op === 'max' ? { bronze: 64, silver: 74, gold: 99 }[q] : null };
    }
    case 'PLAYER_MIN_OVR': return { t: 'ovr', min: Number(v), max: null };
    case 'PLAYER_MAX_OVR': return { t: 'ovr', min: null, max: Number(v) };
    default: return null;
  }
}

export function submitSbc({ setId, challengeId, eaIds }) {
  return guarded('Submitting SBC', async () => {
    await job(`
      var ss = window.__ss, ids = ${JSON.stringify(eaIds)};
      var r = await ss.obs(services.SBC.requestSets());
      var sets = r.response.sets || (r.response.categories || []).reduce(function (a, c) { return a.concat(c.sets || []); }, []);
      var set = sets.find(function (s) { return s.id === ${Number(setId)}; });
      if (!set) throw new Error('SBC set not found — it may have expired');
      var cr = await ss.obs(services.SBC.requestChallengesForSet(set));
      var ch = cr.response.challenges.find(function (c) { return c.id === ${Number(challengeId)}; });
      if (!ch) throw new Error('Challenge not found — it may already be completed');
      await ss.obs(services.SBC.loadChallenge(ch));
      await ss.human();
      var missing = ids.filter(function (id) { return !ss.items[id]; });
      if (missing.length) {
        var offset = 0;
        for (var page = 0; page < 60 && missing.length; page++) {
          var c = new UTSearchCriteriaDTO(); c.type = SearchType.PLAYER; c.count = 91; c.offset = offset;
          var sr = await ss.obs(services.Club.search(c));
          var items = sr.response.items || [];
          items.forEach(function (i) { ss.items[i.id] = i; });
          missing = ids.filter(function (id) { return !ss.items[id]; });
          if (!items.length || sr.response.retrievedAll || items.length < 91) break;
          offset += items.length; await ss.human();
        }
      }
      if (missing.length) throw new Error(missing.length + ' players are no longer in your club. Re-import your club and solve again.');
      var players = ids.map(function (id) { return ss.items[id]; });
      ch.squad.setPlayers(players, true);
      await ss.obs(services.SBC.saveChallenge(ch));
      await ss.human();
      var chemOn = services.Chemistry && services.Chemistry.isFeatureEnabled ? services.Chemistry.isFeatureEnabled() : true;
      await ss.obs(services.SBC.submitChallenge(ch, set, true, chemOn));
      ids.forEach(function (id) { delete ss.items[id]; });
      return true;`, 120000);
    log('SBC submitted.', 'good');
    return true;
  });
}

// Builds a pitch layout for an arbitrary list of EA slot positions.
export function formationFromSlots(positions) {
  const ROW = { GK: 90, LB: 70, RB: 70, CB: 74, LWB: 62, RWB: 62, CDM: 56, LM: 44, RM: 44, CM: 50, CAM: 36, LW: 24, RW: 24, CF: 22, ST: 15 };
  const SIDE = { LB: 12, LWB: 10, LM: 12, LW: 16, RB: 88, RWB: 90, RM: 88, RW: 84 };
  const groups = {};
  positions.forEach((p, i) => { (groups[p] = groups[p] || []).push(i); });
  const slots = positions.map(pos => ({ pos, x: 50, y: ROW[pos] ?? 50 }));
  for (const [pos, idx] of Object.entries(groups)) {
    if (SIDE[pos] != null) { idx.forEach(i => (slots[i].x = SIDE[pos])); continue; }
    const n = idx.length;
    idx.forEach((i, k) => (slots[i].x = n === 1 ? 50 : 22 + (56 * k) / (n - 1)));
  }
  return { name: 'EA squad', slots };
}
