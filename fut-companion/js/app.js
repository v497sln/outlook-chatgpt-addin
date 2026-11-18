import * as E from './engine.js';
import { FORMATIONS, formation, TEMPLATES, sampleClub, DEFAULT_PRICES } from './data.js';
import * as Store from './store.js';
import * as EA from './ea.js';

let S = Store.load();

const ui = {
  tab: 'home',
  sbc: { name: 'Custom SBC', formation: '4-3-3', reqs: [{ t: 'rating', n: 83 }], from: null, custom: null, ea: null, unsupported: [] },
  newReq: { t: 'rating', op: 'min', n: 1, attr: 'league', value: '', min: '', max: '' },
  result: null,
  solving: false,
  seed: 1,
  undo: null,
  clubQ: '', clubFilter: 'all', clubSort: 'rating',
  pickSlot: null, pickQ: '',
};

// ---------------------------------------------------------------- helpers
const sbcFormation = () => ui.sbc.custom || formation(ui.sbc.formation);
const $ = s => document.querySelector(s);
const h = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const fmt = n => Math.round(n || 0).toLocaleString();
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
const clone = o => JSON.parse(JSON.stringify(o));

function persist() {
  if (!Store.save(S)) toast('Could not save — browser storage is full or blocked. Export a backup.', 'bad');
}

let toastTimer;
function toast(msg, kind = '') {
  const t = $('#toast');
  t.className = `toast show ${kind}`;
  t.textContent = msg;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => (t.className = 'toast'), 3200);
}

function openModal(html) {
  const m = $('#modal');
  $('#modal-body').innerHTML = html;
  if (!m.open) m.showModal();
}
// In-app replacements for confirm()/prompt(), which embedded web views often block.
const pending = {};
function askConfirm(msg, yesLabel, onYes, danger = true) {
  pending.yes = onYes;
  openModal(`<h2>Are you sure?</h2><p>${h(msg)}</p>
    <div class="row gap"><button class="btn ${danger ? 'danger' : 'primary'}" data-a="dlgYes">${h(yesLabel)}</button><button class="btn" data-a="close">Cancel</button></div>`);
}
function askNumber(title, value, okLabel, onOk) {
  pending.num = onOk;
  openModal(`<h2>${h(title)}</h2><form data-form="dlgNum"><label>Coins<input id="dlg-num" name="v" type="number" inputmode="numeric" min="0" value="${h(value)}" required></label>
    <div class="row gap"><button class="btn primary">${h(okLabel)}</button><button type="button" class="btn" data-a="close">Cancel</button></div></form>`);
  setTimeout(() => $('#dlg-num')?.select(), 30);
}
function showText(title, filename, text, type) {
  pending.file = { filename, text, type };
  openModal(`<h2>${h(title)}</h2>
    <p class="hint">Copy this text and keep it somewhere safe, for example in a note.</p>
    <textarea id="dlg-text" rows="10" readonly>${h(text)}</textarea>
    <div class="row gap wrap"><button class="btn primary" data-a="dlgCopy">Copy</button>${globalThis.__SOLVER_SRC ? '' : '<button class="btn" data-a="dlgDownload">Download file</button>'}<button class="btn" data-a="close">Close</button></div>`);
}

function closeModal() { const m = $('#modal'); if (m.open) m.close(); ui.pickSlot = null; }

const coin = '<span class="coin" aria-hidden="true"></span>';
const qClass = r => E.quality(r);

function uniqueValues(attr) {
  return [...new Set(S.club.map(p => p[attr]).filter(Boolean))].sort();
}

function clubStats() {
  let cards = 0, untr = 0, prot = 0, value = 0;
  for (const p of S.club) {
    const c = p.count || 1;
    cards += c;
    if (p.untradeable) untr += c;
    if (p.protected) prot += c;
    if (!p.untradeable) value += E.afterTax(p.price || S.prices[p.rating] || 0) * c;
  }
  return { cards, untr, prot, value };
}

function tradeStats() {
  const t = { listed: 0, potIncome: 0, potProfit: 0, soldIncome: 0, soldProfit: 0, sold: 0, unsold: 0 };
  for (const x of S.trades) {
    if (x.status === 'listed') { t.listed++; t.potIncome += E.afterTax(x.list); t.potProfit += E.afterTax(x.list) - x.buy; }
    if (x.status === 'unsold') t.unsold++;
    if (x.status === 'sold') { t.sold++; t.soldIncome += E.afterTax(x.soldFor); t.soldProfit += E.afterTax(x.soldFor) - x.buy; }
  }
  return t;
}

// ---------------------------------------------------------------- components
function card(p, chem, opts = {}) {
  if (!p) return `<div class="pcard empty">${opts.label ? `<b>${h(opts.label)}</b>` : ''}<span>+</span></div>`;
  const tags = [];
  if (p.source === 'market') tags.push('<i class="tag buy">BUY</i>');
  else if (p.untradeable) tags.push('<i class="tag ut">UT</i>');
  if (p.duplicate) tags.push('<i class="tag dup">DUP</i>');
  const pos = opts.label || (p.positions || [])[0] || '';
  const chemDots = chem == null ? '' : `<span class="chem c${chem}" title="${chem} chemistry">${'◆'.repeat(chem)}${'◇'.repeat(3 - chem)}</span>`;
  return `<div class="pcard ${qClass(p.rating)} ${p.rarity === 'icon' || p.rarity === 'hero' ? 'special' : ''} ${p.rare ? 'rare' : ''}">
    <div class="pc-top"><b>${p.rating}</b><small>${h(pos)}</small></div>
    <div class="pc-name">${h(p.generic ? `Any ${p.rating}` : p.name.split(' ').slice(-1)[0])}</div>
    <div class="pc-meta">${p.generic ? 'market' : h((p.nation || '').slice(0, 3).toUpperCase())}</div>
    ${chemDots}<div class="pc-tags">${tags.join('')}</div>
  </div>`;
}

function pitch(squad, f, per, clickable) {
  return `<div class="pitch" role="group" aria-label="Formation ${h(f.name)}">
    <div class="pitch-lines"></div>
    ${f.slots.map((s, i) => `<div class="slot" style="left:${s.x}%;top:${s.y}%">
      ${clickable ? `<button class="slot-btn" data-a="pickSlot" data-i="${i}" aria-label="Choose ${s.pos}">` : ''}
      ${card(squad[i], squad[i] ? per[i] : null, { label: s.pos })}
      ${clickable ? '</button>' : ''}
    </div>`).join('')}
  </div>`;
}

function statTile(label, value, cls = '') {
  return `<div class="stat ${cls}"><span>${h(label)}</span><b>${value}</b></div>`;
}

function select(name, options, current, attrs = '') {
  return `<select ${attrs} data-in="${name}">${options.map(o => {
    const [v, l] = Array.isArray(o) ? o : [o, o];
    return `<option value="${h(v)}" ${String(v) === String(current) ? 'selected' : ''}>${h(l)}</option>`;
  }).join('')}</select>`;
}

// ---------------------------------------------------------------- views
const VIEWS = {
  home() {
    if (EA.available()) return liveHome();
    const c = clubStats();
    const t = tradeStats();
    const empty = !S.club.length;
    return `
      <section class="hero">
        <h1>Your FUT companion that never touches your account.</h1>
        <p>Unlimited SBC solving, squad building and trade tracking — on your device, offline, free.</p>
        <div class="row gap">
          <button class="btn primary" data-a="go" data-tab="sbc">Solve an SBC</button>
          <button class="btn" data-a="go" data-tab="trades">Track trades</button>
        </div>
      </section>
      ${eaCard()}
      ${empty ? `<div class="callout">
        <b>Your club is empty.</b> Import your club as CSV, add players by hand, or try the app with a fictional sample club.
        <div class="row gap"><button class="btn small primary" data-a="importOpen">Import club</button>
        <button class="btn small" data-a="loadSample">Load sample club</button></div></div>` : ''}
      <h2 class="sec">Club</h2>
      <div class="stats">
        ${statTile('Cards', fmt(c.cards))}
        ${statTile('Untradeable', fmt(c.untr))}
        ${statTile('Protected from SBCs', fmt(c.prot))}
        ${statTile('Tradeable value (after tax)', `${fmt(c.value)} ${coin}`)}
      </div>
      <h2 class="sec">Trading</h2>
      <div class="stats">
        ${statTile('On the market', t.listed)}
        ${statTile('Profit if all sell', `${t.potProfit >= 0 ? '+' : ''}${fmt(t.potProfit)} ${coin}`, t.potProfit >= 0 ? 'good' : 'bad')}
        ${statTile('Banked profit', `${t.soldProfit >= 0 ? '+' : ''}${fmt(t.soldProfit)} ${coin}`, t.soldProfit >= 0 ? 'good' : 'bad')}
        ${statTile('Waiting to relist', t.unsold)}
      </div>
      ${S.history.length ? `<h2 class="sec">Recent SBCs</h2><ul class="list">
        ${S.history.slice(0, 5).map(x => `<li class="li"><div><b>${h(x.name)}</b><small>${new Date(x.at).toLocaleString()} · rating ${x.rating} · chem ${x.chem}</small></div>
        <span class="muted">${x.buy ? `bought ${fmt(x.buy)} ${coin}` : 'club only'}</span></li>`).join('')}</ul>` : ''}
      <h2 class="sec">Built from user reviews</h2>
      <ul class="fixes">
        <li><b>No login, no ban worries.</b> Nothing connects to your EA account — you apply the squad yourself.</li>
        <li><b>No daily credits.</b> Solve as many times as you like; a failed attempt costs nothing.</li>
        <li><b>Pre-checked squads.</b> Every requirement is verified and rating is rounded down, so “valid” really means valid.</li>
        <li><b>Works when servers don't.</b> Installs as an app and runs fully offline.</li>
        <li><b>No ads, no paywall, one price: free.</b></li>
        <li><b>Your keepers stay yours.</b> Protect any card and the solver will never use it.</li>
        <li><b>Clear trade numbers.</b> Every figure includes the 5% EA tax and says exactly what it means.</li>
      </ul>`;
  },

  sbc() {
    const s = ui.sbc;
    const saved = S.savedSbcs || [];
    const nr = ui.newReq;
    const attrOpts = [['league', 'League'], ['nation', 'Nation'], ['club', 'Club'], ['rarity', 'Rarity'], ['quality', 'Quality']];
    const valueList = nr.attr === 'quality' ? ['gold', 'silver', 'bronze'] : nr.attr === 'rarity' ? ['common', 'rare', 'icon', 'hero'] : uniqueValues(nr.attr);
    let fields = '';
    if (['rating', 'chem', 'playerChem'].includes(nr.t)) {
      fields = `<label>Minimum<input type="number" inputmode="numeric" data-in="nr.n" value="${h(nr.n)}" min="0" max="${nr.t === 'rating' ? 99 : nr.t === 'chem' ? 33 : 3}"></label>`;
    } else if (nr.t === 'count') {
      fields = `<label>Rule${select('nr.op', [['min', 'At least'], ['max', 'At most'], ['exact', 'Exactly']], nr.op)}</label>
        <label>Players<input type="number" inputmode="numeric" data-in="nr.n" value="${h(nr.n)}" min="0" max="11"></label>
        <label>With${select('nr.attr', [...attrOpts, ['rare', 'Rare (any)']], nr.attr)}</label>
        ${nr.attr === 'rare' ? '' : `<label>Value<input list="vals" data-in="nr.value" value="${h(nr.value)}" placeholder="e.g. ${h(valueList[0] || 'Premier League')}">
        <datalist id="vals">${valueList.map(v => `<option value="${h(v)}">`).join('')}</datalist></label>`}`;
    } else if (nr.t === 'same' || nr.t === 'distinct') {
      fields = `<label>Rule${select('nr.op', [['min', 'At least'], ['max', 'At most'], ['exact', 'Exactly']], nr.op)}</label>
        <label>Number<input type="number" inputmode="numeric" data-in="nr.n" value="${h(nr.n)}" min="0" max="11"></label>
        <label>${nr.t === 'same' ? 'From same' : 'Different'}${select('nr.attr', attrOpts.slice(0, 3), nr.attr)}</label>`;
    } else if (nr.t === 'ovr') {
      fields = `<label>Min player rating<input type="number" data-in="nr.min" value="${h(nr.min)}" placeholder="—"></label>
        <label>Max player rating<input type="number" data-in="nr.max" value="${h(nr.max)}" placeholder="—"></label>`;
    }

    return `
      ${eaSbcBar()}
      ${EA.available() ? '<details class="practice"><summary>Practice templates and your own SBCs</summary>' : '<h2 class="sec first">Pick a challenge</h2>'}
      <div class="chips">
        ${TEMPLATES.map(t => `<button class="chip ${s.from === t.id ? 'on' : ''}" data-a="tpl" data-id="${t.id}">${h(t.name)}</button>`).join('')}
        ${saved.map(t => `<span class="chip-wrap"><button class="chip mine ${s.from === t.id ? 'on' : ''}" data-a="tpl" data-id="${t.id}">★ ${h(t.name)}</button><button class="chip-x" data-a="delSbc" data-id="${t.id}" aria-label="Delete ${h(t.name)}">×</button></span>`).join('')}
      </div>
      ${EA.available() ? '</details>' : ''}

      <div class="panel" id="solver">
        <div class="grid2">
          <label>Name<input data-in="sbc.name" value="${h(s.name)}"></label>
          <label>Formation${s.custom ? `<input value="EA squad: ${h(s.custom.slots.map(x => x.pos).join(' '))}" disabled>` : select('sbc.formation', Object.keys(FORMATIONS), s.formation)}</label>
        </div>
        <h3>Requirements</h3>
        <ul class="reqs">
          ${s.reqs.length ? s.reqs.map((r, i) => `<li><span>${h(E.describeReq(r))}</span><button class="icon-btn" data-a="rmReq" data-i="${i}" aria-label="Remove">×</button></li>`).join('') : '<li class="muted">No requirements yet — add one below.</li>'}
        </ul>
        ${s.ea && s.unsupported.length ? `<div class="callout bad">Not understood yet, so check these yourself (auto-submit is off): ${s.unsupported.map(h).join(', ')}</div>` : ''}
        <details class="addreq" ${s.reqs.length ? '' : 'open'}>
          <summary>Add requirement</summary>
          <div class="grid2">
            <label>Type${select('nr.t', [['rating', 'Team rating'], ['chem', 'Squad chemistry'], ['playerChem', 'Chemistry per player'], ['count', 'Number of players with…'], ['same', 'Players from the same…'], ['distinct', 'Different leagues/nations/clubs'], ['ovr', 'Player rating range']], nr.t)}</label>
            ${fields}
          </div>
          <button class="btn small" data-a="addReq">Add</button>
        </details>
        <div class="toggles">
          <label class="switch"><input type="checkbox" data-in="set.useMarket" ${S.settings.useMarket ? 'checked' : ''}><span></span>Allow buying fodder from the market</label>
          <label class="switch"><input type="checkbox" data-in="set.thorough" ${S.settings.thorough ? 'checked' : ''}><span></span>Thorough search (slower, cheaper results)</label>
        </div>
        <div class="row gap">
          <button class="btn primary grow" data-a="solve" ${ui.solving ? 'disabled' : ''}>${ui.solving ? '<span class="spin"></span> Solving…' : 'Solve'}</button>
          <button class="btn" data-a="saveSbc">Save</button>
        </div>
        <p class="hint">${fmt(clubStats().cards - clubStats().prot)} usable club cards · ${clubStats().prot} protected · solving is unlimited</p>
      </div>
      <div id="result">${renderResult()}</div>`;
  },

  squad() {
    const f = formation(S.squad.formation);
    const byId = new Map(S.club.map(p => [p.id, p]));
    const squad = S.squad.ids.map(id => byId.get(id) || null);
    const ev = E.evaluate(squad, f, []);
    const filled = squad.filter(Boolean).length;
    return `
      <div class="row between center first">
        <h2 class="sec flat">Squad builder</h2>
        <label class="inline">${select('sq.formation', Object.keys(FORMATIONS), S.squad.formation)}</label>
      </div>
      <div class="stats three">
        ${statTile('Rating', filled === 11 ? ev.rating : '—')}
        ${statTile('Chemistry', `${ev.chem.total}/33`)}
        ${statTile('Players', `${filled}/11`)}
      </div>
      ${pitch(squad, f, ev.chem.per, true)}
      <p class="hint">Tap a position to pick a player. Players out of position get 0 chemistry.</p>
      <div class="row gap"><button class="btn" data-a="clearSquad">Clear squad</button>
      <button class="btn" data-a="squadToSbc">Use as SBC formation</button></div>`;
  },

  club() {
    const q = ui.clubQ.toLowerCase();
    let list = S.club.filter(p => !q || [p.name, p.nation, p.league, p.club, (p.positions || []).join(' ')].join(' ').toLowerCase().includes(q));
    const f = ui.clubFilter;
    if (f === 'untradeable') list = list.filter(p => p.untradeable);
    if (f === 'tradeable') list = list.filter(p => !p.untradeable);
    if (f === 'protected') list = list.filter(p => p.protected);
    if (f === 'duplicates') list = list.filter(p => (p.count || 1) > 1);
    const sorts = {
      rating: (a, b) => b.rating - a.rating || a.name.localeCompare(b.name),
      name: (a, b) => a.name.localeCompare(b.name),
      value: (a, b) => (b.price || S.prices[b.rating] || 0) - (a.price || S.prices[a.rating] || 0),
    };
    list.sort(sorts[ui.clubSort]);
    const shown = list.slice(0, 300);
    return `
      <div class="row between center first">
        <h2 class="sec flat">My club <small class="muted">${fmt(clubStats().cards)} cards</small></h2>
        ${EA.available() ? `<span class="muted small">${S.settings.eaLastSync ? `from EA · ${ago(S.settings.eaLastSync)}` : ''}</span>` : '<button class="btn small primary" data-a="editP">+ Add</button>'}
      </div>
      <div class="toolbar">
        <input type="search" placeholder="Search name, nation, league, club, position" data-in="clubQ" value="${h(ui.clubQ)}" aria-label="Search club">
        ${select('clubFilter', [['all', 'All'], ['untradeable', 'Untradeable'], ['tradeable', 'Tradeable'], ['protected', 'Protected'], ['duplicates', 'Duplicates']], ui.clubFilter, 'aria-label="Filter"')}
        ${select('clubSort', [['rating', 'Rating'], ['name', 'Name'], ['value', 'Value']], ui.clubSort, 'aria-label="Sort"')}
      </div>
      <div class="row gap wrap" ${EA.available() ? 'hidden' : ''}>
        <button class="btn small" data-a="importOpen">Import CSV</button>
        <button class="btn small" data-a="exportCsv" ${S.club.length ? '' : 'disabled'}>Export CSV</button>
        <button class="btn small" data-a="loadSample">Sample club</button>
        <button class="btn small danger" data-a="clearClub" ${S.club.length ? '' : 'disabled'}>Clear</button>
      </div>
      <ul class="list">
        ${shown.map(p => `<li class="li player">
          <span class="badge ${qClass(p.rating)} ${p.rare ? 'rare' : ''}">${p.rating}</span>
          <button class="li-main" data-a="editP" data-id="${p.id}">
            <b>${h(p.name)}${(p.count || 1) > 1 ? ` <small class="muted">×${p.count}</small>` : ''}</b>
            <small>${h((p.positions || []).join('/'))} · ${h(p.nation)} · ${h(p.league)} · ${h(p.club)}</small>
          </button>
          <span class="pills">${p.untradeable ? '<i class="tag ut">UT</i>' : ''}${p.rarity === 'icon' || p.rarity === 'hero' ? `<i class="tag sp">${h(p.rarity)}</i>` : ''}</span>
          <button class="icon-btn ${p.protected ? 'on' : ''}" data-a="protect" data-id="${p.id}" aria-pressed="${p.protected}" title="${p.protected ? 'Protected — never used in SBCs' : 'Protect from SBCs'}">${p.protected ? '🔒' : '🔓'}</button>
        </li>`).join('') || '<li class="li muted">No players match.</li>'}
      </ul>
      ${list.length > shown.length ? `<p class="hint">Showing first ${shown.length} of ${list.length}. Refine your search.</p>` : ''}`;
  },

  trades() {
    const t = tradeStats();
    const groups = { listed: [], unsold: [], sold: [] };
    S.trades.forEach(x => groups[x.status]?.push(x));
    const row = x => {
      const sell = x.status === 'sold' ? x.soldFor : x.list;
      const profit = E.afterTax(sell) - x.buy;
      const breakeven = E.roundToMarket(Math.ceil(x.buy / (1 - E.EA_TAX)), 'up');
      return `<li class="li trade">
        <div class="grow"><b>${h(x.name)}</b>
          <small>bought ${fmt(x.buy)} → ${x.status === 'sold' ? 'sold' : 'listed'} ${fmt(sell)} · break-even ${fmt(breakeven)}</small></div>
        <span class="${profit >= 0 ? 'good' : 'bad'} num">${profit >= 0 ? '+' : ''}${fmt(profit)}</span>
        <div class="acts">
          ${x.status === 'listed' ? `<button class="btn tiny" data-a="tSold" data-id="${x.id}">Sold</button><button class="btn tiny" data-a="tUnsold" data-id="${x.id}">Expired</button>` : ''}
          ${x.status === 'unsold' ? `<button class="btn tiny" data-a="tRelist" data-id="${x.id}">Relist</button>` : ''}
          <button class="icon-btn" data-a="tDel" data-id="${x.id}" aria-label="Delete">×</button>
        </div></li>`;
    };
    const sec = (title, arr, extra = '') => `<div class="row between center"><h2 class="sec flat">${title} <small class="muted">${arr.length}</small></h2>${extra}</div>
      <ul class="list">${arr.map(row).join('') || '<li class="li muted">Nothing here.</li>'}</ul>`;
    return `
      <h2 class="sec first">Trade tracker</h2>
      <div class="stats">
        ${statTile('Income if listed items sell', `${fmt(t.potIncome)} ${coin}`)}
        ${statTile('Profit if listed items sell', `${t.potProfit >= 0 ? '+' : ''}${fmt(t.potProfit)} ${coin}`, t.potProfit >= 0 ? 'good' : 'bad')}
        ${statTile('Income from sold items', `${fmt(t.soldIncome)} ${coin}`)}
        ${statTile('Profit from sold items', `${t.soldProfit >= 0 ? '+' : ''}${fmt(t.soldProfit)} ${coin}`, t.soldProfit >= 0 ? 'good' : 'bad')}
      </div>
      <p class="hint">All income and profit figures are after the 5% EA transfer tax.</p>
      ${EA.available() ? `<div class="row gap wrap"><button class="btn small primary" data-a="eaSync" ${EA.state.status === 'ready' ? '' : 'disabled'}>Refresh from EA</button><button class="btn small" data-a="eaRelist" ${EA.state.status === 'ready' ? '' : 'disabled'}>Relist expired on EA</button></div>
      <details class="practice"><summary>Add a trade by hand</summary>` : ''}
      <form class="panel addtrade" data-form="addTrade">
        <label>Card<input name="name" required placeholder="e.g. 86 CB fodder"></label>
        <label>Bought for<input name="buy" type="number" inputmode="numeric" min="0" required></label>
        <label>List at (BIN)<input name="list" type="number" inputmode="numeric" min="150" required></label>
        <button class="btn primary">Add</button>
      </form>
      ${EA.available() ? '</details>' : ''}
      ${sec('On the market', groups.listed)}
      ${sec('Expired / unsold', groups.unsold, groups.unsold.length ? `<span class="row gap center">${select('relistAdj', [['0', 'same price'], ['-5', '−5%'], ['-10', '−10%'], ['5', '+5%'], ['10', '+10%']], ui.relistAdj || '0', 'aria-label="Price change"')}<button class="btn small" data-a="relistAll">Relist all</button></span>` : '')}
      <p class="hint">Relisting never goes below a card's break-even price unless you set it yourself.</p>
      ${sec('Sold', groups.sold, groups.sold.length ? '<button class="btn small" data-a="clearSold">Clear</button>' : '')}`;
  },
};

function ago(t) {
  const m = Math.round((Date.now() - t) / 60000);
  return m < 1 ? 'just now' : m < 60 ? `${m} min ago` : `${Math.round(m / 60)} h ago`;
}

// Live sync: club + transfer list, run automatically after login and on app resume.
let syncing = false;
async function eaSyncAll(quiet = true) {
  if (syncing || EA.state.status !== 'ready') return;
  syncing = true;
  try {
    await EA.readAccount().catch(() => {});
    render();
    await A.eaTransfers(true).catch(e => toast(e.message, 'bad'));
    await loadEaSets().catch(() => {});
    await A.eaImport(quiet);
    S.settings.eaLastSync = Date.now(); S.settings.eaAuto = true; persist(); render();
    if (!quiet) toast('Synced with EA');
  } catch (e) { toast(e.message, 'bad'); } finally { syncing = false; }
}

function eaSubmitBlock(r) {
  if (!ui.sbc.ea) return 'this challenge was not loaded from EA';
  if (EA.state.status !== 'ready' && EA.state.status !== 'busy') return 'you are not logged in with EA';
  if (ui.sbc.unsupported.length) return 'some EA requirements could not be read';
  if (r.squad.some(p => p.source === 'market')) return 'the squad needs cards you would have to buy first';
  if (r.squad.some(p => !p.eaId)) return 'some cards were not imported from EA';
  return '';
}
const canEaSubmit = r => !eaSubmitBlock(r);

function eaCard() {
  const st = EA.state;
  const native = EA.available();
  const label = { off: 'Not connected', opening: 'Opening EA Web App…', waiting: 'Waiting for you to log in…', ready: `Connected · ${st.persona}`, busy: 'Working…' }[st.status];
  const busy = st.status === 'busy' ? 'disabled' : '';
  return `<section class="panel ea">
    <div class="row between center"><h2 class="sec flat">EA account</h2><span class="pill ${st.status === 'ready' || st.status === 'busy' ? 'on' : ''}">${h(label)}</span></div>
    ${!native ? '<p class="hint">Connecting your EA account works in the SquadSage Android app. In a browser, import your club as CSV instead.</p>' : ''}
    ${native && st.status === 'off' ? `<p class="hint">You log in on EA's own Web App page inside SquadSage. Your password goes only to EA. <b class="bad">Automating the Web App is against EA's terms and can get your account restricted or banned.</b></p>
      <button class="btn primary" data-a="eaConnect">Connect EA account</button>` : ''}
    ${native && (st.status === 'waiting' || st.status === 'opening') ? '<div class="row gap wrap"><button class="btn" data-a="eaShow">Show EA login</button><button class="btn" data-a="eaOff">Cancel</button></div>' : ''}
    ${st.status === 'ready' || st.status === 'busy' ? `<p class="hint">${st.status === 'busy' ? 'Syncing with EA…' : S.settings.eaLastSync ? `Club and transfer list synced ${ago(S.settings.eaLastSync)}. Syncs again when you reopen the app.` : 'Syncing…'}</p>
      <div class="row gap wrap">
      <button class="btn small primary" data-a="eaSync" ${busy}>Sync now</button>
      <button class="btn small" data-a="eaRelist" ${busy}>Relist expired</button>
      <button class="btn small" data-a="eaShow">Open Web App</button>
      <button class="btn small danger" data-a="eaOff">Disconnect</button></div>` : ''}
    ${st.log.length ? `<details class="ealog"><summary>Activity (${st.log.length})</summary><ul>${st.log.slice(0, 20).map(l => `<li class="${l.kind}"><time>${new Date(l.at).toLocaleTimeString()}</time> ${h(l.msg)}</li>`).join('')}</ul>
      <button class="btn tiny" data-a="eaCopyLog">Copy log</button></details>` : ''}
  </section>`;
}

async function loadEaSets() {
  ui.eaSets = await EA.loadSbcs();
  render();
}

function eaSbcBar() {
  if (!EA.available()) return '';
  const st = EA.state.status;
  if (st !== 'ready' && st !== 'busy') return `<div class="callout first">Log in with EA on the Home tab to see your live SBCs.</div>`;
  const sets = ui.eaSets;
  if (!sets) return `<h2 class="sec first">Live SBCs</h2><div class="callout"><span class="spin"></span> Loading your SBCs…</div>`;
  const open = ui.eaOpenSet;
  return `<div class="row between center first"><h2 class="sec flat">Live SBCs <small class="muted">${sets.length}</small></h2>
      <button class="btn small" data-a="eaSetsRefresh" ${st === 'busy' ? 'disabled' : ''}>Refresh</button></div>
    <ul class="list sets">${sets.map(x => `<li class="set ${open === x.id ? 'open' : ''}">
      <button class="li li-main" data-a="eaSet" data-id="${x.id}"><b>${h(x.name)}</b><small>${x.done}/${x.challenges} done</small></button>
      ${open === x.id ? (ui.eaChallenges ? `<ul class="chs">${ui.eaChallenges.map((c, i) => `<li><button class="li li-main" data-a="eaPick" data-i="${i}"><b>${h(c.name)}</b><small>${c.reqs.length} requirements · tap to solve</small></button></li>`).join('') || '<li class="muted">All done in this set.</li>'}</ul>`
        : '<div class="hint"><span class="spin"></span> Loading challenges…</div>') : ''}
    </li>`).join('') || '<li class="li muted">No open SBCs right now.</li>'}</ul>`;
}

function liveHome() {
  const st = EA.state;
  const connected = st.status === 'ready' || st.status === 'busy';
  if (!connected) {
    const reconnecting = S.settings.eaAuto && (st.status === 'opening' || st.status === 'waiting');
    return `<section class="login">
      <img src="icon.svg" alt="" width="72" height="72">
      <h1>${reconnecting ? 'Reconnecting to EA…' : 'Log in with your EA account'}</h1>
      <p class="muted">${reconnecting ? 'Using your saved EA login. EA\'s page opens only if it asks you to log in again.' : 'Your club, transfer list, coins and live SBCs load as soon as you log in. You type your password on EA\'s own page.'}</p>
      ${reconnecting ? '<div class="spin big"></div><button class="btn" data-a="eaShow">Show EA page</button>'
        : st.status === 'off' ? '<button class="btn primary wide" data-a="eaConnect">Log in with EA</button>'
        : '<button class="btn primary wide" data-a="eaShow">Continue to EA login</button><button class="btn" data-a="eaOff">Cancel</button>'}
      <p class="hint warnline">Automating the EA Web App is against EA's terms and can get an account restricted or banned.</p>
      ${logBlock()}
    </section>`;
  }
  const c = clubStats();
  const t = tradeStats();
  const expired = S.trades.filter(x => x.status === 'unsold').length;
  return `
    <section class="acct">
      <div><small class="muted">Logged in as</small><b>${h(st.persona)}</b></div>
      <div class="coins">${st.coins != null ? `${fmt(st.coins)} ${coin}` : ''}</div>
    </section>
    ${st.task ? `<div class="syncbar"><span class="spin"></span>${h(st.task)}</div>` : `<p class="hint">Updated ${S.settings.eaLastSync ? ago(S.settings.eaLastSync) : 'just now'} · refreshes every few minutes</p>`}
    <div class="stats">
      ${statTile('Club players', S.club.length || (st.task ? '…' : 0))}
      ${statTile('Untradeable', fmt(c.untr))}
      ${statTile('On transfer list', t.listed)}
      ${statTile('Expired', expired, expired ? 'warn' : '')}
      ${statTile('Profit if listed items sell', `${t.potProfit >= 0 ? '+' : ''}${fmt(t.potProfit)} ${coin}`, t.potProfit >= 0 ? 'good' : 'bad')}
      ${statTile('Sold, after tax', `${fmt(t.soldIncome)} ${coin}`)}
    </div>
    <div class="row gap wrap actions">
      <button class="btn primary" data-a="go" data-tab="sbc">Live SBCs${ui.eaSets ? ` (${ui.eaSets.length})` : ''}</button>
      ${expired ? `<button class="btn" data-a="eaRelist" ${st.status === 'busy' ? 'disabled' : ''}>Relist ${expired} expired</button>` : ''}
      <button class="btn" data-a="eaSync" ${st.status === 'busy' ? 'disabled' : ''}>Refresh</button>
    </div>
    ${S.history.length ? `<h2 class="sec">Recent SBCs</h2><ul class="list">
      ${S.history.slice(0, 5).map(x => `<li class="li"><div><b>${h(x.name)}</b><small>${new Date(x.at).toLocaleString()} · rating ${x.rating} · chem ${x.chem}</small></div></li>`).join('')}</ul>` : ''}
    <div class="row gap wrap foot"><button class="btn small" data-a="eaShow">Open EA Web App</button><button class="btn small danger" data-a="eaOff">Log out</button></div>
    ${logBlock()}`;
}

function logBlock() {
  const st = EA.state;
  return st.log.length ? `<details class="ealog"><summary>Activity (${st.log.length})</summary><ul>${st.log.slice(0, 25).map(l => `<li class="${l.kind}"><time>${new Date(l.at).toLocaleTimeString()}</time> ${h(l.msg)}</li>`).join('')}</ul>
    <button class="btn tiny" data-a="eaCopyLog">Copy log</button></details>` : '';
}

function renderResult() {
  const r = ui.result;
  if (!r) return '';
  if (r.error) return `<div class="callout bad">${h(r.error)}</div>`;
  if (!r.squad) return `<div class="callout bad">${h(r.reason)}</div>`;
  const f = sbcFormation();
  const ev = r.eval;
  const fromClub = r.squad.filter(p => p.source === 'club');
  const buys = r.squad.filter(p => p.source === 'market');
  const clubValue = fromClub.filter(p => !p.untradeable).reduce((a, p) => a + E.afterTax(p.price || S.prices[p.rating] || 0), 0);
  return `
    <div class="callout ${r.ok ? 'good' : 'bad'}">${r.ok ? '✔ All requirements met. Build this squad in-game and submit.' : h(r.reason)}</div>
    <div class="stats">
      ${statTile('Team rating', ev.rating)}
      ${statTile('Chemistry', `${ev.chem.total}/33`)}
      ${statTile('Coins to spend', `${fmt(r.buyCost)} ${coin}`, r.buyCost ? 'warn' : 'good')}
      ${statTile('Tradeable club value used', `${fmt(clubValue)} ${coin}`)}
    </div>
    ${pitch(r.squad, f, ev.chem.per, false)}
    <h3>Requirement check</h3>
    <ul class="checks">${ev.results.map(x => `<li class="${x.ok ? 'ok' : 'no'}"><span>${x.ok ? '✔' : '✖'}</span>${h(x.label)}<em>${h(x.actual)}</em></li>`).join('')}</ul>
    ${buys.length ? `<h3>Buy from the market</h3><ul class="list">${buys.map(p => `<li class="li"><span class="badge ${qClass(p.rating)}">${p.rating}</span><div class="grow">Any ${p.rating}-rated card</div><span class="num">~${fmt(p.price)} ${coin}</span></li>`).join('')}</ul>` : ''}
    <h3>From your club</h3>
    <ul class="list">${fromClub.map(p => `<li class="li"><span class="badge ${qClass(p.rating)} ${p.rare ? 'rare' : ''}">${p.rating}</span>
      <div class="grow"><b>${h(p.name)}</b><small>${h(p.nation)} · ${h(p.league)} · ${h(p.club)}</small></div>
      ${p.untradeable ? '<i class="tag ut">UT</i>' : ''}${p.duplicate ? '<i class="tag dup">DUP</i>' : ''}</li>`).join('')}</ul>
    <div class="row gap wrap">
      <button class="btn" data-a="solveAgain">Try another solution</button>
      <button class="btn" data-a="copyList">Copy list</button>
      ${r.ok && canEaSubmit(r) ? '<button class="btn primary" data-a="eaSubmit">Submit to EA</button>' : ''}
      ${r.ok ? `<button class="btn ${canEaSubmit(r) ? '' : 'primary'}" data-a="submitted">I submitted it</button>` : ''}
    </div>
    ${r.ok && ui.sbc.ea && !canEaSubmit(r) ? `<p class="hint">Submit to EA is off: ${h(eaSubmitBlock(r))}.</p>` : ''}
    <p class="hint">“I submitted it” removes the club cards used so your next solve stays accurate. You can undo.</p>`;
}

// ---------------------------------------------------------------- render
function render() {
  $('#main').innerHTML = VIEWS[ui.tab]();
  document.querySelectorAll('.nav button').forEach(b => {
    const on = b.dataset.tab === ui.tab;
    b.classList.toggle('on', on);
    b.setAttribute('aria-current', on ? 'page' : 'false');
  });
}

// ---------------------------------------------------------------- solving
let worker = null;
function getWorker() {
  if (worker) return worker;
  try {
    const src = globalThis.__SOLVER_SRC;
    worker = src
      ? new Worker(URL.createObjectURL(new Blob([src], { type: 'text/javascript' })))
      : new Worker(new URL('./solver-worker.js', import.meta.url), { type: 'module' });
  } catch { worker = null; }
  return worker;
}

function runSolve(newSeed = false) {
  if (!ui.sbc.reqs.length) return toast('Add at least one requirement', 'bad');
  if (newSeed) ui.seed++;
  ui.solving = true; render();
  const payload = {
    club: S.club, formation: sbcFormation(), reqs: ui.sbc.reqs,
    priceTable: S.prices, useMarket: S.settings.useMarket,
    timeMs: S.settings.thorough ? 4000 : 1500, seed: ui.seed * 7919,
  };
  const done = res => {
    ui.solving = false;
    ui.result = res;
    render();
    $('#result')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  };
  const w = getWorker();
  if (w) {
    w.onmessage = e => done(e.data.ok ? e.data.result : { error: e.data.error });
    w.onerror = () => { worker = null; setTimeout(() => done(E.solve(payload)), 0); };
    w.postMessage(payload);
  } else setTimeout(() => done(E.solve(payload)), 30);
}

// ---------------------------------------------------------------- actions
const A = {
  go: el => { ui.tab = el.dataset.tab; render(); scrollTo(0, 0); },
  loadSample() {
    const go = () => { S.club = sampleClub(); persist(); closeModal(); render(); toast('Sample club loaded (fictional players)'); };
    if (S.club.length) askConfirm('Replace your club with the fictional sample club?', 'Replace club', go); else go();
  },
  clearClub() {
    askConfirm('Delete every player in your club? Export first if you want a backup.', 'Delete all players', () => {
      S.club = []; persist(); closeModal(); render(); toast('Club cleared');
    });
  },
  exportCsv: () => showText('Export club (CSV)', 'club.csv', Store.toCsv(S.club), 'text/csv'),
  dlgYes() { const f = pending.yes; pending.yes = null; f?.(); },
  async dlgCopy() {
    const ta = $('#dlg-text');
    try { await navigator.clipboard.writeText(ta.value); toast('Copied'); }
    catch { ta.focus(); ta.select(); toast('Text selected — copy it from the menu'); }
  },
  dlgDownload() { const f = pending.file; if (f) Store.download(f.filename, f.text, f.type); },
  importOpen() {
    openModal(`<h2>Import club</h2>
      <p class="hint">CSV columns: <code>${Store.CSV_HEADER}</code>. Positions are separated by <code>|</code>, e.g. <code>CM|CDM</code>. Only name and rating are required.</p>
      <label>Choose a file<input type="file" accept=".csv,text/csv,text/plain" id="imp-file"></label>
      <label>…or paste CSV<textarea id="imp-text" rows="6" placeholder="${h(Store.CSV_HEADER)}"></textarea></label>
      <label class="switch"><input type="checkbox" id="imp-replace"><span></span>Replace my current club</label>
      <div id="imp-err"></div>
      <div class="row gap"><button class="btn primary" data-a="importRun">Import</button><button class="btn" data-a="close">Cancel</button></div>`);
  },
  async importRun() {
    const file = $('#imp-file').files[0];
    const text = file ? await file.text() : $('#imp-text').value;
    const { players, errors } = Store.parseCsv(text);
    if (!players.length) {
      $('#imp-err').innerHTML = `<div class="callout bad">No players imported.<br>${errors.slice(0, 5).map(h).join('<br>')}</div>`;
      return;
    }
    S.club = $('#imp-replace').checked ? players : [...S.club, ...players];
    persist(); closeModal(); render();
    toast(`Imported ${players.length} players${errors.length ? ` · ${errors.length} rows skipped` : ''}`);
  },
  protect(el) {
    const p = S.club.find(x => x.id === el.dataset.id);
    if (!p) return;
    p.protected = !p.protected; persist(); render();
    toast(p.protected ? `${p.name} will never be used in SBCs` : `${p.name} can be used in SBCs`);
  },
  editP(el) {
    const p = S.club.find(x => x.id === el.dataset.id) || { id: '', name: '', rating: 80, positions: [], nation: '', league: '', club: '', rarity: 'common', untradeable: false, count: 1, price: 0, protected: false };
    const dl = (id, attr) => `<datalist id="${id}">${uniqueValues(attr).map(v => `<option value="${h(v)}">`).join('')}</datalist>`;
    openModal(`<h2>${p.id ? 'Edit player' : 'Add player'}</h2>
      <form data-form="saveP" data-id="${h(p.id)}" class="grid2">
        <label class="span2">Name<input name="name" required value="${h(p.name)}"></label>
        <label>Rating<input name="rating" type="number" min="1" max="99" required value="${p.rating}"></label>
        <label>Positions<input name="positions" placeholder="CM|CDM" value="${h((p.positions || []).join('|'))}"></label>
        <label>Nation<input name="nation" list="dl-n" value="${h(p.nation)}">${dl('dl-n', 'nation')}</label>
        <label>League<input name="league" list="dl-l" value="${h(p.league)}">${dl('dl-l', 'league')}</label>
        <label>Club<input name="club" list="dl-c" value="${h(p.club)}">${dl('dl-c', 'club')}</label>
        <label>Rarity<select name="rarity">${['common', 'rare', 'icon', 'hero'].map(r => `<option ${p.rarity === r ? 'selected' : ''}>${r}</option>`).join('')}</select></label>
        <label>Copies<input name="count" type="number" min="1" value="${p.count || 1}"></label>
        <label>Price (optional)<input name="price" type="number" min="0" value="${p.price || ''}" placeholder="uses fodder table"></label>
        <label class="switch span2"><input type="checkbox" name="untradeable" ${p.untradeable ? 'checked' : ''}><span></span>Untradeable</label>
        <label class="switch span2"><input type="checkbox" name="protected" ${p.protected ? 'checked' : ''}><span></span>Protect — never use in SBCs</label>
        <div class="row gap span2"><button class="btn primary">Save</button>
        ${p.id ? `<button type="button" class="btn danger" data-a="delP" data-id="${p.id}">Delete</button>` : ''}
        <button type="button" class="btn" data-a="close">Cancel</button></div>
      </form>`);
  },
  delP(el) {
    S.club = S.club.filter(p => p.id !== el.dataset.id); persist(); closeModal(); render();
  },
  close: closeModal,

  // EA account
  eaConnect() {
    if (S.settings.eaWarned) return EA.connect();
    askConfirm("SquadSage will automate the EA Web App for you: reading your club, submitting SBCs and relisting. This is against EA's terms and EA can restrict or ban accounts for it. You use it at your own risk.", 'I understand, connect', () => {
      S.settings.eaWarned = true; persist(); closeModal(); EA.connect();
    });
  },
  eaShow: () => EA.show(),
  eaOff: () => { EA.disconnect(); S.settings.eaAuto = false; ui.eaSets = null; ui.eaOpenSet = null; persist(); render(); },
  async eaCopyLog() {
    const text = EA.state.log.map(l => `${new Date(l.at).toISOString()} [${l.kind}] ${l.msg}`).join('\n');
    try { await navigator.clipboard.writeText(text); toast('Log copied'); } catch { showText('Activity log', 'log.txt', text, 'text/plain'); }
  },
  eaSync: () => eaSyncAll(false),
  async eaImport(quiet) {
    try {
      const players = await EA.importClub();
      const keep = new Map(S.club.filter(p => p.eaId).map(p => [p.definitionId, p]));
      const imported = players.map(p => {
        const old = keep.get(p.definitionId);
        return { id: `ea${p.eaId}`, ...p, count: 1, price: old?.price || 0, protected: old?.protected || false };
      });
      S.club = [...S.club.filter(p => !p.eaId), ...imported];
      persist(); render(); if (!quiet) toast(`Imported ${imported.length} players from EA`);
    } catch (e) { if (quiet) throw e; toast(e.message, 'bad'); }
  },
  async eaTransfers(quiet) {
    try {
      const items = await EA.readTransferList();
      const byEa = new Map(S.trades.filter(t => t.eaId).map(t => [t.eaId, t]));
      for (const i of items) {
        const status = i.tradeState === 'closed' ? 'sold' : i.tradeState === 'expired' ? 'unsold' : i.tradeState === 'active' ? 'listed' : null;
        if (!status) continue;
        const t = byEa.get(i.eaId) || { id: uid(), eaId: i.eaId, at: Date.now(), buy: i.paid || 0 };
        Object.assign(t, { name: `${i.rating} ${i.name}`, list: i.bin || t.list || 0, status });
        if (status === 'sold') t.soldFor = i.bid || i.bin;
        if (!byEa.has(i.eaId)) S.trades.unshift(t);
      }
      persist(); render(); if (!quiet) toast(`Synced ${items.length} transfer list items`);
    } catch (e) { if (quiet) throw e; toast(e.message, 'bad'); }
  },
  async eaRelist() {
    try { await EA.relistExpired(); await A.eaTransfers(); } catch (e) { toast(e.message, 'bad'); }
  },
  async eaSbcs() {
    try {
      const sets = await EA.loadSbcs();
      openModal(`<h2>EA SBCs</h2><ul class="list picker">${sets.map(s => `<li><button class="li li-main" data-a="eaSet" data-id="${s.id}"><b>${h(s.name)}</b><small>${s.done}/${s.challenges} done</small></button></li>`).join('') || '<li class="li muted">No open SBCs found.</li>'}</ul>
        <button class="btn" data-a="close">Close</button>`);
    } catch (e) { toast(e.message, 'bad'); }
  },
  async eaSet(el) {
    const id = +el.dataset.id;
    if (ui.eaOpenSet === id) { ui.eaOpenSet = null; render(); return; }
    ui.eaOpenSet = id; ui.eaChallenges = null; render();
    try { ui.eaChallenges = await EA.loadChallenges(id); } catch (e) { ui.eaOpenSet = null; toast(e.message, 'bad'); }
    render();
  },
  async eaSetsRefresh() { try { await loadEaSets(); } catch (e) { toast(e.message, 'bad'); } },
  eaPick(el) {
    const c = ui.eaChallenges[+el.dataset.i];
    const reqs = []; const unsupported = [];
    for (const r of c.reqs) {
      const t = EA.translateReq(r);
      if (t) reqs.push(t); else unsupported.push(`${r.key} ${r.scope} ${r.values.join('/')}`);
    }
    if (c.bricks) unsupported.push('locked (brick) slots');
    const slots = c.slots.filter(Boolean);
    if (slots.length !== 11) unsupported.push('squad positions (using 4-3-3, chemistry may differ)');
    ui.sbc = { name: c.name, formation: '4-3-3', reqs, from: null, unsupported,
      custom: slots.length === 11 ? EA.formationFromSlots(slots) : null,
      ea: { setId: c.setId, challengeId: c.id } };
    ui.result = null; closeModal(); render();
    document.getElementById('solver')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
    runSolve(false);
  },
  eaSubmit() {
    const r = ui.result;
    if (!r?.ok || !canEaSubmit(r)) return;
    askConfirm(`Submit "${ui.sbc.name}" to EA with these ${r.squad.length} players? They will be used up.`, 'Submit to EA', async () => {
      closeModal();
      try {
        await EA.submitSbc({ ...ui.sbc.ea, eaIds: r.squad.map(p => p.eaId) });
        A.submitted();
        toast('Submitted to EA');
      } catch (e) { toast(e.message, 'bad'); }
    }, false);
  },

  // SBC
  tpl(el) {
    const t = [...TEMPLATES, ...(S.savedSbcs || [])].find(x => x.id === el.dataset.id);
    if (!t) return;
    ui.sbc = { name: t.name, formation: t.formation, reqs: clone(t.reqs), from: t.id, custom: null, ea: null, unsupported: [] };
    ui.result = null; render();
  },
  delSbc(el) {
    S.savedSbcs = S.savedSbcs.filter(x => x.id !== el.dataset.id); persist(); render();
  },
  rmReq(el) { ui.sbc.reqs.splice(+el.dataset.i, 1); ui.sbc.from = null; render(); },
  addReq() {
    const n = ui.newReq;
    const num = v => (v === '' || v == null ? null : Number(v));
    let r;
    if (['rating', 'chem', 'playerChem'].includes(n.t)) r = { t: n.t, n: num(n.n) ?? 0 };
    else if (n.t === 'count') {
      if (n.attr !== 'rare' && !String(n.value).trim()) return toast('Enter a value, e.g. a league name', 'bad');
      r = { t: 'count', attr: n.attr, value: n.attr === 'rare' ? true : String(n.value).trim(), op: n.op, n: num(n.n) ?? 1 };
    } else if (n.t === 'same' || n.t === 'distinct') r = { t: n.t, attr: n.attr === 'rarity' || n.attr === 'quality' || n.attr === 'rare' ? 'league' : n.attr, op: n.op, n: num(n.n) ?? 1 };
    else if (n.t === 'ovr') {
      r = { t: 'ovr', min: num(n.min), max: num(n.max) };
      if (r.min == null && r.max == null) return toast('Set a min or max rating', 'bad');
      ui.sbc.reqs = ui.sbc.reqs.filter(x => x.t !== 'ovr');
    }
    if (['rating', 'chem', 'playerChem'].includes(r.t)) ui.sbc.reqs = ui.sbc.reqs.filter(x => x.t !== r.t);
    ui.sbc.reqs.push(r); ui.sbc.from = null; render();
  },
  saveSbc() {
    const s = ui.sbc;
    if (!s.reqs.length) return toast('Add requirements first', 'bad');
    const existing = (S.savedSbcs || []).find(x => x.name === s.name && x.id.startsWith('u-'));
    if (existing) Object.assign(existing, { formation: s.formation, reqs: clone(s.reqs) });
    else S.savedSbcs = [...(S.savedSbcs || []), { id: `u-${uid()}`, name: s.name || 'My SBC', formation: s.formation, reqs: clone(s.reqs) }];
    persist(); render(); toast('Saved to your challenges');
  },
  solve: () => runSolve(false),
  solveAgain: () => runSolve(true),
  async copyList() {
    const r = ui.result; if (!r?.squad) return;
    const f = sbcFormation();
    const text = `${ui.sbc.name} (${f.name}) — rating ${r.eval.rating}, chem ${r.eval.chem.total}\n` +
      r.squad.map((p, i) => `${f.slots[i].pos}: ${p.generic ? `BUY any ${p.rating} (~${fmt(p.price)})` : `${p.name} ${p.rating}`}`).join('\n');
    try { await navigator.clipboard.writeText(text); toast('Copied'); } catch { showText('Squad list', 'squad.txt', text, 'text/plain'); }
  },
  submitted() {
    const r = ui.result; if (!r?.ok) return;
    ui.undo = { club: clone(S.club), history: clone(S.history) };
    for (const p of r.squad) {
      if (p.source !== 'club') continue;
      const c = S.club.find(x => x.id === p.id);
      if (!c) continue;
      c.count = (c.count || 1) - 1;
    }
    S.club = S.club.filter(p => (p.count || 1) > 0 && p.count !== 0);
    S.history = [{ name: ui.sbc.name, at: Date.now(), rating: r.eval.rating, chem: r.eval.chem.total, buy: r.buyCost }, ...S.history].slice(0, 50);
    ui.result = null; persist(); render();
    showUndo();
  },
  undo() {
    if (!ui.undo) return;
    S.club = ui.undo.club; S.history = ui.undo.history; ui.undo = null;
    persist(); render(); toast('Restored your club');
  },

  // Squad
  pickSlot(el) { ui.pickSlot = +el.dataset.i; ui.pickQ = ''; openPicker(); },
  assign(el) {
    const i = ui.pickSlot;
    S.squad.ids = S.squad.ids.map(id => (id === el.dataset.id ? null : id));
    S.squad.ids[i] = el.dataset.id || null;
    persist(); closeModal(); render();
  },
  clearSquad() { S.squad.ids = Array(11).fill(null); persist(); render(); },
  squadToSbc() { ui.sbc.formation = S.squad.formation; ui.tab = 'sbc'; render(); toast(`SBC formation set to ${S.squad.formation}`); },

  // Trades
  tSold(el) {
    const x = S.trades.find(t => t.id === el.dataset.id);
    askNumber(`${x.name} sold for (before tax)`, x.list, 'Mark sold', v => {
      const n = parseInt(v, 10);
      if (!(n > 0)) return toast('Enter a valid price', 'bad');
      x.status = 'sold'; x.soldFor = n; persist(); closeModal(); render();
    });
  },
  tUnsold(el) { const x = S.trades.find(t => t.id === el.dataset.id); x.status = 'unsold'; persist(); render(); },
  tRelist(el) {
    const x = S.trades.find(t => t.id === el.dataset.id);
    askNumber(`Relist ${x.name} at (BIN)`, x.list, 'Relist', v => {
      x.list = E.roundToMarket(parseInt(v, 10) || x.list); x.status = 'listed'; persist(); closeModal(); render();
    });
  },
  tDel(el) { S.trades = S.trades.filter(t => t.id !== el.dataset.id); persist(); render(); },
  relistAll() {
    const adj = +(ui.relistAdj || 0) / 100;
    let floored = 0;
    for (const x of S.trades) {
      if (x.status !== 'unsold') continue;
      const breakeven = E.roundToMarket(Math.ceil(x.buy / (1 - E.EA_TAX)), 'up');
      let p = E.roundToMarket(x.list * (1 + adj));
      if (adj < 0 && p < breakeven && x.list >= breakeven) { p = breakeven; floored++; }
      x.list = p; x.status = 'listed';
    }
    persist(); render();
    toast(floored ? `Relisted · ${floored} kept at break-even to avoid a loss` : 'Relisted');
  },
  clearSold() { S.trades = S.trades.filter(t => t.status !== 'sold'); persist(); render(); },

  // Settings
  settings() {
    openModal(`<h2>Settings</h2>
      <h3>Fodder prices</h3>
      <p class="hint">The solver uses these when it suggests buying “any card” of a rating, and to value tradeable club cards without a price. Update them to today's market.</p>
      <div class="prices">${Object.keys(S.prices).sort((a, b) => a - b).map(r => `<label><span class="badge ${qClass(+r)}">${r}</span><input type="number" min="0" data-price="${r}" value="${S.prices[r]}"></label>`).join('')}</div>
      <div class="row gap"><button class="btn small" data-a="savePrices">Save prices</button><button class="btn small" data-a="resetPrices">Reset to defaults</button></div>
      <h3>Backup</h3>
      <p class="hint">Everything lives only on this device. Export a backup to move to another phone.</p>
      <div class="row gap wrap"><button class="btn small" data-a="backup">Export backup</button>
      <label class="btn small">Restore from file<input type="file" accept=".json,application/json" id="restore" hidden></label>
      <button class="btn small" data-a="restorePaste">Restore from text</button>
      <button class="btn small danger" data-a="resetAll">Erase all data</button></div>
      <h3>About</h3>
      <p class="hint">SquadSage is an independent fan tool. It is not affiliated with or endorsed by EA. It never logs in to or automates your game account.</p>
      <div class="row"><button class="btn" data-a="close">Done</button></div>`);
  },
  restorePaste() {
    openModal(`<h2>Restore backup</h2><p class="hint">Paste the backup text you copied earlier. This replaces everything on this device.</p>
      <textarea id="restore-text" rows="8"></textarea>
      <div class="row gap"><button class="btn primary" data-a="restoreRun">Restore</button><button class="btn" data-a="close">Cancel</button></div>`);
  },
  restoreRun() { restoreFrom($('#restore-text').value); },
  savePrices() {
    document.querySelectorAll('[data-price]').forEach(i => { S.prices[i.dataset.price] = Math.max(0, parseInt(i.value, 10) || 0); });
    persist(); toast('Prices saved');
  },
  resetPrices() { S.prices = { ...DEFAULT_PRICES }; persist(); A.settings(); },
  backup: () => showText('Backup', `squadsage-backup-${new Date().toISOString().slice(0, 10)}.json`, JSON.stringify(S), 'application/json'),
  resetAll() {
    askConfirm('Erase your club, trades, saved SBCs and settings from this device?', 'Erase everything', () => {
      S = Store.defaultState(); persist(); closeModal(); ui.result = null; render(); toast('All data erased');
    });
  },
};

function showUndo() {
  const t = $('#toast');
  t.className = 'toast show';
  t.innerHTML = 'Club updated — used cards removed. <button class="link" data-a="undo">Undo</button>';
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => (t.className = 'toast'), 8000);
}

function openPicker() {
  const f = formation(S.squad.formation);
  const pos = f.slots[ui.pickSlot].pos;
  const q = ui.pickQ.toLowerCase();
  const list = S.club
    .filter(p => !q || `${p.name} ${p.nation} ${p.league} ${p.club}`.toLowerCase().includes(q))
    .sort((a, b) => (E.inPosition(b, pos) - E.inPosition(a, pos)) || b.rating - a.rating)
    .slice(0, 120);
  const html = `<h2>Choose ${h(pos)}</h2>
    <input type="search" id="pick-q" placeholder="Search" value="${h(ui.pickQ)}" aria-label="Search players">
    <ul class="list picker">
      ${S.squad.ids[ui.pickSlot] ? '<li><button class="li-main" data-a="assign" data-id=""><b>Remove player</b></button></li>' : ''}
      ${list.map(p => `<li class="li"><span class="badge ${qClass(p.rating)} ${p.rare ? 'rare' : ''}">${p.rating}</span>
        <button class="li-main" data-a="assign" data-id="${p.id}"><b>${h(p.name)}</b>
        <small>${h((p.positions || []).join('/'))} · ${h(p.nation)} · ${h(p.league)} · ${h(p.club)}</small></button>
        ${E.inPosition(p, pos) ? '' : '<i class="tag dup">OOP</i>'}</li>`).join('') || '<li class="li muted">No players. Import your club first.</li>'}
    </ul>
    <button class="btn" data-a="close">Cancel</button>`;
  openModal(html);
  const inp = $('#pick-q');
  inp.oninput = () => { ui.pickQ = inp.value; const pos2 = inp.selectionStart; openPicker(); const i2 = $('#pick-q'); i2.focus(); i2.setSelectionRange(pos2, pos2); };
}

// ---------------------------------------------------------------- events
document.addEventListener('click', e => {
  const el = e.target.closest('[data-a]');
  if (!el) {
    const nav = e.target.closest('.nav button');
    if (nav) { ui.tab = nav.dataset.tab; render(); scrollTo(0, 0); }
    return;
  }
  const fn = A[el.dataset.a];
  if (fn) { e.preventDefault(); fn(el, e); }
});

document.addEventListener('submit', e => {
  const form = e.target.closest('[data-form]');
  if (!form) return;
  e.preventDefault();
  const fd = new FormData(form);
  if (form.dataset.form === 'dlgNum') { const f = pending.num; pending.num = null; f?.(fd.get('v')); return; }
  if (form.dataset.form === 'addTrade') {
    const buy = parseInt(fd.get('buy'), 10), list = parseInt(fd.get('list'), 10);
    if (!(buy >= 0) || !(list >= 150)) return toast('Enter valid prices', 'bad');
    S.trades.unshift({ id: uid(), name: String(fd.get('name')).trim(), buy, list: E.roundToMarket(list), status: 'listed', at: Date.now() });
    persist(); render();
    if (E.afterTax(list) < buy) toast('Heads up: this listing loses coins after the 5% tax', 'bad');
  }
  if (form.dataset.form === 'saveP') {
    const id = form.dataset.id;
    const rarity = fd.get('rarity');
    const data = {
      name: String(fd.get('name')).trim(),
      rating: Math.min(99, Math.max(1, parseInt(fd.get('rating'), 10) || 1)),
      positions: String(fd.get('positions') || '').split(/[|/ ,;]+/).map(s => s.toUpperCase()).filter(Boolean),
      nation: String(fd.get('nation')).trim(), league: String(fd.get('league')).trim(), club: String(fd.get('club')).trim(),
      rarity, rare: rarity !== 'common',
      count: Math.max(1, parseInt(fd.get('count'), 10) || 1),
      price: Math.max(0, parseInt(fd.get('price'), 10) || 0),
      untradeable: fd.get('untradeable') === 'on', protected: fd.get('protected') === 'on',
    };
    if (id) Object.assign(S.club.find(p => p.id === id), data);
    else S.club.unshift({ id: `p${uid()}`, ...data });
    persist(); closeModal(); render(); toast('Saved');
  }
});

function restoreFrom(t) {
  try {
    const data = JSON.parse(t);
    if (!Array.isArray(data.club)) throw new Error('bad');
    S = { ...Store.defaultState(), ...data }; persist(); closeModal(); render(); toast('Backup restored');
  } catch { toast('That is not a SquadSage backup', 'bad'); }
}

function onInput(e) {
  const el = e.target;
  if (el.id === 'restore' && el.files?.[0]) {
    el.files[0].text().then(restoreFrom);
    return;
  }
  const k = el.dataset.in;
  if (!k) return;
  const v = el.type === 'checkbox' ? el.checked : el.value;
  const rerender = e.type === 'change';
  if (k.startsWith('nr.')) { ui.newReq[k.slice(3)] = v; if (rerender && el.tagName === 'SELECT') render(); return; }
  if (k === 'sbc.name') { ui.sbc.name = v; return; }
  if (k === 'sbc.formation') { ui.sbc.formation = v; ui.sbc.custom = null; ui.sbc.ea = null; ui.result = null; render(); return; }
  if (k.startsWith('set.')) { S.settings[k.slice(4)] = v; persist(); return; }
  if (k === 'sq.formation') { S.squad.formation = v; persist(); render(); return; }
  if (k === 'clubQ') {
    ui.clubQ = v;
    const pos = el.selectionStart; render();
    const i = document.querySelector('[data-in="clubQ"]'); i.focus(); i.setSelectionRange(pos, pos);
    return;
  }
  if (k === 'clubFilter' || k === 'clubSort') { ui[k] = v; render(); return; }
  if (k === 'relistAdj') { ui.relistAdj = v; }
}
document.addEventListener('input', e => { if (e.target.dataset.in === 'clubQ') onInput(e); else if (e.target.dataset.in?.startsWith('nr.') && e.target.tagName !== 'SELECT') onInput(e); else if (e.target.dataset.in === 'sbc.name') onInput(e); });
document.addEventListener('change', e => { if (e.target.dataset.in !== 'clubQ') onInput(e); });

$('#modal').addEventListener('close', () => { ui.pickSlot = null; });
$('#modal').addEventListener('click', e => { if (e.target.id === 'modal') closeModal(); });

// Offline support
EA.setReadyHandler(() => eaSyncAll(true));
document.addEventListener('deviceready', () => {
  render();
  // Reconnect in the background if you've connected before.
  if (S.settings.eaAuto && EA.available()) EA.connect({ hidden: true });
});
setInterval(() => {
  if (document.visibilityState === 'visible' && EA.state.status === 'ready' && Date.now() - (S.settings.eaLastSync || 0) > 5 * 60000) eaSyncAll(true);
}, 60000);
document.addEventListener('resume', () => {
  if (EA.state.status === 'ready' && Date.now() - (S.settings.eaLastSync || 0) > 3 * 60000) eaSyncAll(true);
});
if ('serviceWorker' in navigator && location.protocol !== 'file:' && !globalThis.__SOLVER_SRC && !window.Capacitor) {
  navigator.serviceWorker.register('./sw.js').catch(() => {});
}
window.addEventListener('online', () => $('#net').textContent = 'Online');
window.addEventListener('offline', () => $('#net').textContent = 'Offline — still works');

EA.setListener(() => { if (!document.activeElement || document.activeElement.tagName !== 'INPUT') render(); });
render();
