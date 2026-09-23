'use strict';
/* ============ CONFIG: edit before deploying ============ */
const CONFIG = {
  clientId: '00000000-0000-0000-0000-000000000000',        // Entra app registration (SPA platform)
  tenantId: '00000000-0000-0000-0000-000000000000',        // UiPath tenant ID
  site: 'uipath.sharepoint.com:/sites/HauntedTrail2026',   // "<host>:/sites/<path>" of the SharePoint site holding the lists
  lists: { vault: 'HQ_Vault', pub: 'HQ_Public', subs: 'HQ_Submissions' },
  opensAt: '09:00', closesAt: '20:00',                      // IST, same for every level
  offsite: '2026-10-09T12:00:00+05:30',
  // the drive from Onyx to Rosetta along NH75; km are approximate road distance from Bengaluru
  routeKm: 230,
  levels: [
    { date: '2026-09-24', title: 'The Whispering Gate', icon: '🕸️', kind: 'Riddle', place: 'Onyx, Bengaluru', km: 0, scene: 'city' },
    { date: '2026-09-26', title: 'The Cursed Toll Booth', icon: '🚧', kind: 'Riddle', place: 'Nelamangala', km: 28, scene: 'city' },
    { date: '2026-09-29', title: 'Lake of Lost Souls', icon: '🪷', kind: 'Riddle', place: 'Kunigal', km: 70, scene: 'plains' },
    { date: '2026-10-02', title: 'The Midnight Dhaba', icon: '🍛', kind: 'Riddle', place: 'Yediyur', km: 88, scene: 'plains' },
    { date: '2026-10-05', title: 'Night of the Living Bugs', icon: '🕷️', kind: 'Daily quiz', place: 'Channarayapatna', km: 145, scene: 'hills' },
    { date: '2026-10-06', title: 'The Crypt of Complexity', icon: '⚰️', kind: 'Daily quiz', place: 'Hassan', km: 185, scene: 'hills' },
    { date: '2026-10-07', title: 'Ghosts of Production Past', icon: '👻', kind: 'Daily quiz', place: 'Hemavathi River', km: 222, scene: 'ghats' },
    { date: '2026-10-08', title: 'The Final Séance', icon: '🔮', kind: 'Daily quiz', place: 'Manjarabad Fort', km: 226, scene: 'ghats' },
  ],
};
/* ======================================================= */

const TZ = 'Asia/Kolkata', LISTS = CONFIG.lists, SCOPES = ['User.Read', 'Sites.ReadWrite.All'];
const params = new URLSearchParams(location.search);
const DEMO = params.has('demo'), ADMIN = params.has('admin');
const BOOT = Date.now(), NOW0 = params.get('now') ? +new Date(params.get('now')) : BOOT;
const now = () => (DEMO ? NOW0 + (Date.now() - BOOT) : Date.now());   // ?demo&now=2026-10-06T21:00 time-travels
const OFFSITE = +new Date(CONFIG.offsite);
const LEVELS = CONFIG.levels.map((l, i) => ({
  ...l, n: i + 1,
  open: +new Date(`${l.date}T${CONFIG.opensAt}:00+05:30`),
  close: +new Date(`${l.date}T${CONFIG.closesAt}:00+05:30`),
}));

const S = { me: null, pub: { levels: {} }, board: [], mine: {}, view: 'path', admin: false, stats: null, scrolled: false };
let api;

/* ---------- tiny DOM helper: children are always text nodes, never HTML ---------- */
const $ = s => document.querySelector(s);
function h(tag, attrs, ...kids) {
  const el = tag.startsWith('svg:') ? document.createElementNS('http://www.w3.org/2000/svg', tag.slice(4)) : document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v == null || v === false) continue;
    if (k === 'class') el.setAttribute('class', v);
    else if (k === 'style') Object.assign(el.style, v);
    else if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
    else el.setAttribute(k, v === true ? '' : v);
  }
  for (const c of kids.flat(Infinity)) if (c != null && c !== false) el.append(c instanceof Node ? c : String(c));
  return el;
}
const fmt = (t, o) => new Intl.DateTimeFormat('en-IN', { timeZone: TZ, ...o }).format(t);
const day = t => fmt(t, { weekday: 'short', day: 'numeric', month: 'short' });
const clock = t => fmt(t, { hour: 'numeric', minute: '2-digit' });
const dur = ms => {
  const m = Math.max(0, Math.round(ms / 6e4)), d = Math.floor(m / 1440), hh = Math.floor(m % 1440 / 60);
  return d ? `${d}d ${hh}h` : hh ? `${hh}h ${m % 60}m` : `${m % 60}m`;
};
let toastT;
function toast(msg) {
  const t = $('#toast'); t.textContent = msg; t.classList.add('on');
  clearTimeout(toastT); toastT = setTimeout(() => t.classList.remove('on'), 3200);
}
const store = { get: k => { try { return localStorage.getItem(k); } catch { return null; } },
                set: (k, v) => { try { localStorage.setItem(k, v); } catch {} } };

/* ---------- auth: MSAL, tokens cached in localStorage so people sign in once ---------- */
let pca, account;
async function auth() {
  pca = new msal.PublicClientApplication({
    auth: { clientId: CONFIG.clientId, authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
            redirectUri: new URL('.', location.href).href },
    cache: { cacheLocation: 'localStorage' },
  });
  await pca.initialize();
  const res = await pca.handleRedirectPromise();
  account = res?.account || pca.getActiveAccount() || pca.getAllAccounts()[0];
  if (account) pca.setActiveAccount(account);
  return account;
}
async function token() {
  try {
    return (await pca.acquireTokenSilent({ scopes: SCOPES, account })).accessToken;   // refresh token, then Entra SSO cookie
  } catch (e) {
    if (e instanceof msal.InteractionRequiredAuthError) await pca.acquireTokenRedirect({ scopes: SCOPES, account, redirectStartPage: location.href });
    throw e;
  }
}

/* ---------- storage: Microsoft Graph -> SharePoint lists ---------- */
function graphApi() {
  let siteId;
  async function g(path, opt = {}) {
    const r = await fetch(path.startsWith('https://') ? path : `https://graph.microsoft.com/v1.0${path}`, {
      ...opt, headers: { Authorization: `Bearer ${await token()}`, 'Content-Type': 'application/json' },
    });
    if (!r.ok) {
      const e = new Error(`${r.status} ${(await r.json().catch(() => ({}))).error?.message || r.statusText}`);
      e.status = r.status; throw e;
    }
    return r.status === 204 ? null : r.json();
  }
  const site = async () => (siteId ??= (await g(`/sites/${CONFIG.site}?$select=id`)).id);
  const norm = it => ({
    id: it.id, f: it.fields,
    by: { id: String(it.createdBy?.user?.email || it.createdBy?.user?.id || '').toLowerCase(), name: it.createdBy?.user?.displayName || 'Unknown soul' },
    at: +new Date(it.lastModifiedDateTime),   // stamped by SharePoint, not the browser
  });
  return {
    async items(list) {
      const out = [];
      let url = `/sites/${await site()}/lists/${list}/items?$expand=fields&$top=999`;
      while (url) { const j = await g(url); out.push(...j.value); url = j['@odata.nextLink']; }
      return out.map(norm);
    },
    add: async (list, fields) => g(`/sites/${await site()}/lists/${list}/items`, { method: 'POST', body: JSON.stringify({ fields }) }),
    patch: async (list, id, fields) => g(`/sites/${await site()}/lists/${list}/items/${id}/fields`, { method: 'PATCH', body: JSON.stringify(fields) }),
    createList: async (name, columns) => g(`/sites/${await site()}/lists`, {
      method: 'POST', body: JSON.stringify({ displayName: name, columns, list: { template: 'genericList' } }) }),
  };
}

/* ---------- demo: in-browser fake SharePoint, with its own throwaway questions ---------- */
function demoApi() {
  const KEY = 'hq-demo-v1', me = { id: 'demo@you', name: 'You (demo)' };
  let db = null;
  try { db = JSON.parse(store.get(KEY)); } catch {}
  if (!db) {
    const Q = [
      [1, 'What gets wetter the more it dries?', '', 'towel', 'A towel, obviously.'],
      [2, 'Which structure is Last-In-First-Out, like bodies in a crypt?', 'Queue\nStack\nHeap\nTrie', 'Stack', ''],
      [3, 'What has many teeth but never bites?', '', 'comb', ''],
      [4, 'What runs but never walks, has a mouth but never talks?', '', 'river', ''],
      [5, "In JavaScript, '2' + 2 is…", '4\n22\nNaN\nTypeError', '22', 'String concatenation wins.'],
      [5, 'Which month is Halloween in?', 'September\nOctober\nNovember', 'October', ''],
      [6, 'Bits in a byte?', '4\n8\n16', '8', ''], [6, 'Dracula is a…', 'Werewolf\nVampire\nZombie', 'Vampire', ''],
      [7, 'HTTP 404 means…', 'Not Found\nForbidden\nGone', 'Not Found', ''], [7, 'The Red Planet?', 'Venus\nMars\nJupiter', 'Mars', ''],
      [8, 'Git command to see history?', 'git log\ngit show\ngit blame', 'git log', ''], [8, 'Who wrote Frankenstein?', 'Bram Stoker\nMary Shelley\nEdgar Allan Poe', 'Mary Shelley', ''],
    ];
    let seq = 1, rnd = 7;
    const rand = () => (rnd = (rnd * 16807) % 2147483647) / 2147483647;
    const vault = Q.map(([Level, Prompt, Options, Answer, Explanation], i) => ({ id: String(seq++), f: { Title: '', Level, Seq: i, Prompt, Options, Answer, Explanation }, by: me, at: 0 }));
    const ghosts = ['Casper', 'Count Byte', 'Null Pointer', 'Wanda the Witch', 'Sir Segfault', 'Lady Latency', 'The Headless Build', 'Mummy Merge', 'Pumpkin Spice Dev', 'Ghoul-ang'];
    const subs = [];
    for (const [gi, name] of ghosts.entries()) for (const L of LEVELS) {
      if (rand() < 0.2) continue;
      const qs = Q.filter(q => q[0] === L.n);
      const answers = qs.map(q => (rand() < 0.75 - gi * 0.04 ? q[3] : (q[2].split('\n')[1] || 'dunno')));
      subs.push({ id: String(seq++), f: { Level: L.n, Answers: JSON.stringify(answers) }, by: { id: `ghost${gi}`, name }, at: L.open + rand() * (L.close - L.open) });
    }
    // you (demo) answered the first four: three right, one wrong
    [['towel'], ['Stack'], ['brush'], ['river']].forEach((a, i) => subs.push({ id: String(seq++), f: { Level: i + 1, Answers: JSON.stringify(a) }, by: me, at: LEVELS[i].open + 36e5 }));
    db = { seq, lists: { [LISTS.vault]: vault, [LISTS.pub]: [], [LISTS.subs]: subs } };
  }
  const save = () => store.set(KEY, JSON.stringify(db));
  const need = l => { if (!db.lists[l]) throw Object.assign(new Error('404 list not found'), { status: 404 }); return db.lists[l]; };
  return {
    me,
    items: async l => structuredClone(need(l)),
    add: async (l, f) => { need(l).push({ id: String(db.seq++), f, by: me, at: now() }); save(); },
    patch: async (l, id, f) => { const it = need(l).find(i => i.id === id); Object.assign(it.f, f); it.at = now(); save(); },
    createList: async l => { db.lists[l] ??= []; save(); },
  };
}

/* ---------- keeper sync: publish opened levels, reveal closed ones, recompute board ---------- */
async function sync() {
  const [vault, pubItems, subs] = await Promise.all([api.items(LISTS.vault), api.items(LISTS.pub), api.items(LISTS.subs)]);
  const t = now(), levels = {}, keys = {}, stats = {};
  for (const L of LEVELS) {
    const rows = vault.filter(r => +r.f.Level === L.n).sort((a, b) => (a.f.Seq ?? 0) - (b.f.Seq ?? 0));
    stats[L.n] = { q: rows.length, subs: subs.filter(s => +s.f.Level === L.n).length };
    if (!rows.length || t < L.open) continue;
    const lv = levels[L.n] = { questions: rows.map(r => ({ prompt: r.f.Prompt || '', options: String(r.f.Options || '').split('\n').map(s => s.trim()).filter(Boolean) })) };
    if (t >= L.close) {
      keys[L.n] = lv.answers = rows.map(r => String(r.f.Answer || ''));
      lv.explanations = rows.map(r => r.f.Explanation || '');
    }
  }
  const board = Score.board(LEVELS, keys, subs.map(s => ({ uid: s.by.id, name: s.by.name, level: +s.f.Level, answers: s.f.Answers, at: s.at })));
  const fields = { Data: JSON.stringify({ levels }), Board: JSON.stringify(board) };
  const cur = pubItems[0];
  if (!cur) await api.add(LISTS.pub, { Title: 'state', ...fields });
  else if (cur.f.Data !== fields.Data || cur.f.Board !== fields.Board) await api.patch(LISTS.pub, cur.id, fields);
  S.stats = { ...stats, at: t };
}

/* ---------- load player view ---------- */
const isMe = id => S.me.ids.includes(id);
const myEntry = () => S.board.find(e => isMe(e.id));
async function load() {
  if (S.admin === true) await sync();
  const [pubItems, subs] = await Promise.all([api.items(LISTS.pub), api.items(LISTS.subs)]);   // SharePoint only returns MY submissions
  const f = pubItems[0]?.f || {};
  S.pub = f.Data ? JSON.parse(f.Data) : { levels: {} };
  S.board = f.Board ? JSON.parse(f.Board) : [];
  S.mine = {};
  for (const s of subs) {
    const L = LEVELS[s.f.Level - 1];
    if (!isMe(s.by.id) || !L || s.at < L.open || s.at > L.close) continue;
    let a; try { a = JSON.parse(s.f.Answers); } catch { continue; }
    if (!S.mine[L.n] || s.at > S.mine[L.n].at) S.mine[L.n] = { answers: a, at: s.at };
  }
}

function stateOf(L) {
  const t = now(), pub = S.pub.levels[L.n], mine = S.mine[L.n];
  if (t < L.open) return 'locked';
  if (pub?.answers) {
    if (!mine) return 'missed';
    const r = myEntry()?.r?.[L.n];
    return r && r[0] > 0 && r[0] * 2 >= r[1] ? 'won' : 'lost';
  }
  if (t >= L.close) return mine ? 'pending' : 'missed';
  if (!pub) return 'brewing';
  return mine ? 'sealed' : 'open';
}
const BADGE = { locked: '🔒', brewing: '🧪', sealed: '🕯️', pending: '⏳', won: '✓', lost: '✗', missed: '💀' };

/* ---------- render ---------- */
function render() {
  renderHud();
  const app = $('#app');
  app.replaceChildren(...[S.view === 'board' ? renderBoard() : renderTrail(), ADMIN && renderAdmin()].filter(Boolean),
    h('footer', {}, 'Made with 🎃 for the India PnE offsite · Bengaluru → Rosetta Sakleshpur, 9–11 Oct',
      DEMO && h('div', {}, h('a', { href: '#', onclick: e => { e.preventDefault(); try { localStorage.removeItem('hq-demo-v1'); } catch {} location.reload(); } }, 'reset demo'))));
  celebrate();
  if (!S.scrolled && S.view === 'path') { S.scrolled = true; app.querySelector('.cur')?.scrollIntoView({ block: 'center' }); }
}

function nextText() {
  const t = now(), L = LEVELS.find(l => t < l.close);
  const nights = Math.ceil((OFFSITE - t) / 864e5);
  const tail = nights > 0 ? ` · ${nights} night${nights > 1 ? 's' : ''} to Rosetta` : '';
  if (!L) return t < OFFSITE ? `🚌 ${dur(OFFSITE - t)} until Rosetta Sakleshpur` : '🎃 Happy Halloween, Sakleshpur!';
  return (t < L.open ? `🗝️ Level ${L.n} awakens in ${dur(L.open - t)}` : `⏳ Level ${L.n} seals in ${dur(L.close - t)}`) + tail;
}

function renderHud() {
  const me = myEntry(), rank = me ? S.board.indexOf(me) + 1 : 0;
  const chip = (icon, v, label) => h('div', { class: 'chip' }, icon, h('div', {}, v, h('small', {}, label)));
  const tab = (v, label) => h('button', { class: 'tab' + (S.view === v ? ' on' : ''), onclick: () => { S.view = v; render(); scrollTo(0, 0); } }, label);
  $('#hud').replaceChildren(
    h('div', { class: 'brand' }, h('h1', {}, 'The Haunted Trail'), h('p', {}, 'Bengaluru → Sakleshpur · NH75', DEMO && h('span', { class: 'tag' }, 'DEMO'))),
    h('div', { class: 'stats' }, chip('🎃', me?.p ?? 0, 'points'), chip('🔥', me?.s ?? 0, 'streak'), chip('🏆', rank ? `#${rank}` : '–', `of ${S.board.length || '–'}`)),
    h('div', { class: 'next' }, nextText()),
    routeBar(),
    h('nav', { class: 'tabs' }, tab('path', '🗺️ Trail'), tab('board', '🪦 Hall of Haunts'), h('button', { class: 'tab', onclick: rules }, '📜 Rules')));
}

function routeBar() {
  const lit = LEVELS.filter(L => stateOf(L) !== 'locked').at(-1);
  const km = now() >= OFFSITE ? CONFIG.routeKm : lit?.km ?? 0, pct = Math.max(2, Math.min(98, 100 * km / CONFIG.routeKm));
  return h('div', { class: 'route', title: `${km} of ~${CONFIG.routeKm} km` },
    h('div', { class: 'fill', style: { width: `${pct}%` } }), h('span', { class: 'bus', style: { left: `${pct}%` } }, '🚌'),
    h('small', {}, `km ${km}`), h('small', {}, `~${CONFIG.routeKm} km · Rosetta`));
}

// Silhouettes along the route: skyline, then palms and fields, boulder hills, then misty Ghats with coffee rows.
function scenery(pts, H) {
  let seed = 11;
  const r = () => (seed = (seed * 16807) % 2147483647) / 2147483647;
  const g = h('svg:svg', { class: 'scene', viewBox: `0 0 400 ${H}`, preserveAspectRatio: 'none', height: H });
  const add = (tag, a) => g.append(h(`svg:${tag}`, a));
  const palm = (x, y, s) => {
    add('path', { d: `M${x} ${y} q${4 * s} ${-22 * s} ${1 * s} ${-44 * s}`, class: 'trunk', 'stroke-width': 3 * s });
    for (const [dx, dy] of [[-22, 6], [-14, -8], [0, -14], [14, -8], [22, 6]])
      add('path', { d: `M${x + s} ${y - 44 * s} q${dx * s / 2} ${(dy - 12) * s} ${dx * s} ${dy * s}`, class: 'frond', 'stroke-width': 3 * s });
  };
  LEVELS.forEach((L, i) => {
    const { y } = pts[i], left = pts[i].x > 50, edge = left ? 0 : 290;   // scenery on the side away from the road
    if (L.scene === 'city') {
      for (let x = edge; x < edge + 110; x += 14 + r() * 10) {
        const w = 14 + r() * 16, bh = 40 + r() * 90;
        add('rect', { x, y: y + 40 - bh, width: w, height: bh, class: 'bldg' });
        for (let wy = y + 48 - bh; wy < y + 34; wy += 9) if (r() < 0.35) add('rect', { x: x + 3 + r() * (w - 8), y: wy, width: 3, height: 4, class: 'win' });
      }
    } else if (L.scene === 'plains') {
      for (let k = 0; k < 3; k++) add('rect', { x: 0, y: y + 30 + k * 10, width: 400, height: 4, class: 'field' });
      for (let k = 0; k < 3; k++) palm(edge + 18 + k * 36 + r() * 10, y + 34, 0.8 + r() * 0.5);
    } else if (L.scene === 'hills') {
      add('path', { d: `M0 ${y + 60} Q70 ${y - 20} 140 ${y + 40} T280 ${y + 30} T400 ${y + 10} V${y + 90} H0Z`, class: 'hill' });
      add('ellipse', { cx: edge + 60, cy: y + 18, rx: 34, ry: 26, class: 'rock' });
    } else {
      add('path', { d: `M0 ${y + 70} L50 ${y - 50} L110 ${y + 10} L180 ${y - 80} L250 ${y} L320 ${y - 60} L400 ${y - 10} V${y + 110} H0Z`, class: 'peak' });
      for (let row = 0; row < 3; row++) for (let x = 8; x < 400; x += 16) add('circle', { cx: x + (row % 2) * 8, cy: y + 62 + row * 12, r: 4, class: 'coffee' });
      add('ellipse', { cx: 120 + r() * 160, cy: y + 20, rx: 180, ry: 18, class: 'mist' });
    }
  });
  return g;
}

function renderTrail() {
  const ROW = 170, TOP = 110, n = LEVELS.length;
  const pts = [...LEVELS, null].map((_, i) => ({ x: 50 + Math.sin(i * 1.2) * 24, y: TOP + i * ROW }));
  const H = TOP + n * ROW + 120;
  const states = LEVELS.map(stateOf);
  const lastLit = states.findLastIndex(s => s !== 'locked');
  const curve = list => list.map((p, i) => (i ? `C${list[i - 1].x} ${list[i - 1].y + ROW / 2} ${p.x} ${p.y - ROW / 2} ${p.x} ${p.y}` : `M${p.x} ${p.y}`)).join('');
  const cur = states.findIndex(s => s === 'open' || s === 'brewing' || s === 'sealed');
  const svg = h('svg:svg', { viewBox: `0 0 100 ${H}`, preserveAspectRatio: 'none', height: H },
    h('svg:path', { class: 'road', d: curve(pts) }), h('svg:path', { class: 'lane', d: curve(pts) }),
    lastLit >= 0 && h('svg:path', { class: 'lit', d: curve(pts.slice(0, lastLit + 1)) }));
  const stops = LEVELS.map((L, i) => {
    const st = states[i];
    return h('div', { class: 'stop' + (i === (cur >= 0 ? cur : lastLit) ? ' cur' : ''), style: { left: `${pts[i].x}%`, top: `${pts[i].y}px` } },
      st === 'open' && h('div', { class: 'bubble' }, 'PLAY'),
      st === 'sealed' && h('div', { class: 'bubble alt' }, 'SEALED'),
      h('button', { class: `node ${st}`, 'aria-label': `Level ${L.n}: ${L.title}, ${st}`, onclick: () => openLevel(L) }, L.icon, BADGE[st] && h('span', { class: 'badge' }, BADGE[st])),
      h('div', { class: 'lbl' }, h('b', {}, L.title), h('span', { class: 'place' }, `📍 ${L.place} · ${L.km ? `~${L.km} km` : 'start'}`), `L${L.n} · ${day(L.open)}`));
  });
  const home = pts[n];
  stops.push(h('div', { class: 'stop', style: { left: `${home.x}%`, top: `${home.y}px` } },
    h('button', { class: 'node home', 'aria-label': 'Rosetta Sakleshpur', onclick: () => toast('The haunted mansion awaits, 9 Oct. Costumes mandatory. 🧛') }, '🏚️'),
    h('div', { class: 'lbl', style: { top: '64px' } }, h('b', {}, 'Rosetta Sakleshpur'), 'Halloween night · 9–11 Oct')));
  return h('div', { class: 'trail', style: { height: `${H}px` } }, scenery(pts.map(p => ({ x: p.x, y: p.y })), H), svg, stops);
}

function renderBoard() {
  const first = LEVELS[0];
  return h('div', {}, h('h2', { class: 'board-title' }, 'Hall of Haunts'),
    S.board.length
      ? h('ol', { class: 'board' }, S.board.map((e, i) => h('li', { class: isMe(e.id) ? 'me' : '' },
          h('span', { class: 'rk' }, ['🥇', '🥈', '🥉'][i] || i + 1),
          h('span', { class: 'nm' }, e.name, isMe(e.id) && ' (you)'),
          h('span', { class: 'st' }, e.s ? `🔥${e.s}` : ''),
          h('span', { class: 'pt' }, e.p))))
      : h('p', { class: 'empty' }, `🕯️ No souls ranked yet. The first reveal is ${day(first.close)} at ${clock(first.close)}.`));
}

/* ---------- level sheet ---------- */
function sheet(...kids) {
  const s = $('#sheet');
  s.hidden = false;
  s.replaceChildren(h('div', { class: 'scrim', onclick: closeSheet }),
    h('div', { class: 'panel', role: 'dialog', 'aria-modal': 'true' }, h('button', { class: 'x', onclick: closeSheet, 'aria-label': 'Close' }, '✕'), ...kids));
}
function closeSheet() { $('#sheet').hidden = true; $('#sheet').replaceChildren(); }
addEventListener('keydown', e => e.key === 'Escape' && closeSheet());

function openLevel(L) {
  const st = stateOf(L), pub = S.pub.levels[L.n], mine = S.mine[L.n];
  const kind = pub ? (pub.questions.length > 1 ? `Quiz · ${pub.questions.length} questions` : 'Riddle') : L.kind;
  const head = h('div', { class: 'lvhead' }, h('div', { class: 'lvicon' }, L.icon),
    h('div', {}, h('small', {}, `Level ${L.n} · ${kind}`), h('h2', {}, L.title), h('small', {}, `${day(L.open)} · ${clock(L.open)} – ${clock(L.close)} IST`)));
  const p = (t, c) => h('p', { class: c }, t);
  if (st === 'locked') return sheet(head, p('🔒 This crypt is still sealed.'), p(`It creaks open ${day(L.open)} at ${clock(L.open)} and seals at ${clock(L.close)}. Answers are revealed after it seals.`, 'muted'));
  if (st === 'brewing') return sheet(head, p('🧪 The spirits are still scribbling this one. Check back in a little while.'));
  if (st === 'open' || st === 'sealed') return sheet(head, quizForm(L, pub, mine));
  if (!pub?.answers) {
    return sheet(head, mine ? p(`⏳ Your answer is sealed. The keeper lights the lanterns after ${clock(L.close)}, and then the truth comes out.`) : p('💀 You slept through this one. The next crypt awaits.'),
      mine && pub && pub.questions.map((q, i) => [h('div', { class: 'qp' }, q.prompt), h('div', { class: 'opt sel' }, mine.answers[i] ?? '—')]));
  }
  const r = myEntry()?.r?.[L.n];
  sheet(head,
    h('div', { class: 'result' }, !mine ? '💀 No answer sealed. The ghosts took this one.' : r ? `${r[0]}/${r[1]} correct · +${r[2]} 🎃` : '✗ No points for this seal.'),
    pub.questions.map((q, i) => {
      const key = pub.answers[i], given = mine?.answers[i];
      const opts = q.options.length ? q.options : [given, key.split('|')[0]].filter((v, j, a) => v && a.indexOf(v) === j);
      return [h('div', { class: 'qp', style: { marginTop: '18px' } }, q.prompt),
        opts.map(o => h('div', { class: 'opt' + (Score.match(o, key) ? ' right' : o === given ? ' wrong' : '') },
          o, Score.match(o, key) ? '  ✓' : o === given ? '  ✗ (you)' : '')),
        pub.explanations[i] && h('p', { class: 'expl' }, `🕯️ ${pub.explanations[i]}`)];
    }));
}

function quizForm(L, pub, mine) {
  const ans = pub.questions.map((_, i) => mine?.answers[i] ?? '');
  const form = h('form', {});
  pub.questions.forEach((q, i) => {
    const fs = h('fieldset', { class: 'q' }, h('legend', {}, pub.questions.length > 1 ? `${i + 1}. ` : '', q.prompt));
    if (q.options.length) {
      for (const o of q.options) {
        const b = h('button', { type: 'button', class: 'opt' + (ans[i] === o ? ' sel' : ''), 'aria-pressed': String(ans[i] === o),
          onclick: () => { ans[i] = o; fs.querySelectorAll('.opt').forEach(x => { x.classList.toggle('sel', x === b); x.setAttribute('aria-pressed', String(x === b)); }); } }, o);
        fs.append(b);
      }
    } else {
      fs.append(h('input', { class: 'txt', value: ans[i], maxlength: 120, placeholder: 'Whisper your answer…', autocomplete: 'off', 'aria-label': 'Your answer', oninput: e => { ans[i] = e.target.value; } }));
    }
    form.append(fs);
  });
  const btn = h('button', { class: 'btn', type: 'submit' }, mine ? 'Re-seal my answer 🕯️' : 'Seal my answer 🕯️');
  form.append(h('p', { class: 'muted' }, mine
    ? `Sealed at ${clock(mine.at)}. You can change it until ${clock(L.close)}, and the last seal counts (but re-sealing later lowers your speed bonus).`
    : `Answers stay sealed until ${clock(L.close)}. Seal early for a bigger speed bonus.`), btn);
  form.addEventListener('submit', async e => {
    e.preventDefault();
    const clean = ans.map(a => String(a).trim());
    if (clean.some(a => !a)) return toast('Answer every question, mortal. 👻');
    if (now() >= L.close) return toast('Too late. The crypt has sealed.');
    btn.disabled = true; btn.textContent = 'Sealing…';
    try {
      await api.add(LISTS.subs, { Title: `L${L.n}`, Level: L.n, Answers: JSON.stringify(clean) });
      S.mine[L.n] = { answers: clean, at: now() };
      closeSheet(); render(); toast('Sealed in the crypt 🕯️ See you after 8 PM.');
    } catch (err) {
      btn.disabled = false; btn.textContent = 'Try again';
      toast(`The spirits refused: ${err.message}`);
    }
  });
  return form;
}

function rules() {
  sheet(h('div', { class: 'lvhead' }, h('div', { class: 'lvicon' }, '📜'), h('h2', {}, 'Rules of the Trail')),
    h('ul', { class: 'rules' },
      h('li', {}, `${LEVELS.length} haunted stops on the drive from Onyx to Rosetta: two a week, then one every day from Mon 5 to Thu 8 Oct.`),
      h('li', {}, `Each crypt opens at ${clock(LEVELS[0].open)} and seals at ${clock(LEVELS[0].close)} IST the same day.`),
      h('li', {}, '100 🎃 per correct answer. Seal early for up to +50 speed bonus.'),
      h('li', {}, 'Clear crypts back to back (at least half right) to build a 🔥 streak: +20 per streak step, up to +100.'),
      h('li', {}, 'Change your mind before it seals. Your last seal counts.'),
      h('li', {}, 'Answers and the Hall of Haunts update after the crypt seals.'),
      h('li', {}, '😈 To the hackers: SharePoint stamps every seal, not your browser. A seal edited after closing is void, and nobody can read anyone else’s seals. Happy hunting.')));
}

let seenWins = null;
function celebrate() {
  seenWins ??= new Set(JSON.parse(store.get('hq-wins') || '[]'));
  const fresh = LEVELS.filter(L => stateOf(L) === 'won' && !seenWins.has(L.n));
  if (!fresh.length) return;
  fresh.forEach(L => seenWins.add(L.n));
  store.set('hq-wins', JSON.stringify([...seenWins]));
  for (let i = 0; i < 14; i++) {
    const b = h('span', { class: 'boom' }, i % 3 ? '🦇' : '🎃');
    b.style.setProperty('--dx', `${(Math.random() - 0.5) * 90}vw`);
    b.style.setProperty('--dy', `${(Math.random() - 0.7) * 70}vh`);
    b.style.setProperty('--r', `${(Math.random() - 0.5) * 720}deg`);
    document.body.append(b); setTimeout(() => b.remove(), 1700);
  }
  toast(`Crypt cleared! +${fresh.reduce((s, L) => s + (myEntry()?.r?.[L.n]?.[2] || 0), 0)} 🎃`);
}

/* ---------- keeper (admin) panel: ?admin. Real access control = SharePoint permissions on HQ_Vault ---------- */
async function initAdmin() {
  try { await api.items(LISTS.vault); S.admin = true; } catch (e) { S.admin = e.status === 404 ? 'missing' : false; S.adminErr = e.message; }
}
function renderAdmin() {
  const log = h('pre', {});
  const say = m => { log.textContent += m + '\n'; };
  const run = fn => async () => { try { await fn(); } catch (e) { say(`✗ ${e.message}`); } };
  if (S.admin === false) return h('div', { class: 'admin' }, h('h3', {}, '🗝️ Keeper'), `Not the keeper (vault not readable): ${S.adminErr}`);
  if (S.admin === 'missing') {
    return h('div', { class: 'admin' }, h('h3', {}, '🗝️ Keeper · first-time setup'), 'Lists not found on the site.',
      h('button', { class: 'btn', onclick: run(async () => {
        const num = name => ({ name, number: {} }), multi = name => ({ name, text: { allowMultipleLines: true, textType: 'plain' } });
        const defs = { [LISTS.vault]: [num('Level'), num('Seq'), multi('Prompt'), multi('Options'), { name: 'Answer', text: {} }, multi('Explanation')],
                       [LISTS.pub]: [multi('Data'), multi('Board')], [LISTS.subs]: [num('Level'), multi('Answers')] };
        for (const [name, cols] of Object.entries(defs)) {
          try { await api.createList(name, cols); say(`✓ created ${name}`); } catch (e) { if (e.status !== 409) throw e; say(`• ${name} exists`); }
        }
        if (!(await api.items(LISTS.pub)).length) await api.add(LISTS.pub, { Title: 'state', Data: '{"levels":{}}', Board: '[]' });
        say('✓ Done. NOW lock down permissions (SETUP.md step 5), then reload.');
      }) }, 'Create the 3 lists'), log);
  }
  const st = S.stats || {};
  return h('div', { class: 'admin' }, h('h3', {}, '🗝️ Keeper'),
    h('div', { class: 'muted' }, `Auto-syncs on open and every 5 min while this tab is open. Last sync: ${st.at ? `${day(st.at)} ${clock(st.at)}` : '—'}`),
    h('table', {}, h('tr', {}, h('th', {}, 'L'), h('th', {}, 'Date'), h('th', {}, 'Qs in vault'), h('th', {}, 'Seals'), h('th', {}, 'State')),
      LEVELS.map(L => h('tr', {}, h('td', {}, L.n), h('td', {}, day(L.open)), h('td', {}, st[L.n]?.q ?? '?', st[L.n]?.q === 0 && now() > L.open - 864e5 ? ' ⚠️' : ''), h('td', {}, st[L.n]?.subs ?? '?'), h('td', {}, S.pub.levels[L.n]?.answers ? 'revealed' : S.pub.levels[L.n] ? 'published' : now() < L.open ? 'locked' : '—')))),
    h('button', { class: 'btn', onclick: run(async () => { await load(); render(); toast('Synced 🕯️'); }) }, 'Sync now'),
    h('p', { class: 'muted' }, 'Import questions (JSON from questions.json, appended to the vault):'),
    h('input', { type: 'file', accept: '.json,application/json', onchange: run(async e => {
      const qs = JSON.parse(await e.target.files[0].text());
      for (const q of qs) {
        await api.add(LISTS.vault, { Title: LEVELS[q.level - 1]?.title || '', Level: q.level, Seq: q.seq ?? 0, Prompt: q.prompt,
          Options: (q.options || []).join('\n'), Answer: q.answer, Explanation: q.explanation || '' });
        say(`✓ L${q.level}: ${q.prompt.slice(0, 40)}…`);
      }
      await load(); render(); toast(`Imported ${qs.length} questions`);
    }) }), log);
}

/* ---------- boot ---------- */
function splash(err) {
  $('#hud').replaceChildren();
  $('#app').replaceChildren(h('div', { class: 'splash' }, h('div', { class: 'pumpkin' }, '🎃'), h('h1', {}, 'The Haunted Trail'),
    h('p', {}, `${LEVELS.length} haunted stops on NH75, Bengaluru to Rosetta Sakleshpur.`),
    err ? h('p', { style: { color: 'var(--blood)' } }, err)
        : h('button', { class: 'btn', onclick: () => pca.loginRedirect({ scopes: SCOPES, redirectStartPage: location.href }) }, 'Enter with Microsoft 365'),
    h('p', { class: 'muted' }, 'Sign in once. We keep you signed in on this device.')));
}
async function refresh() {
  try { await load(); render(); } catch (e) {
    console.error(e);
    $('#app').replaceChildren(h('p', { class: 'empty' }, e.status === 403 || e.status === 404
      ? '🚷 The trail is hidden from you. Ask the organisers to add you to the guest list.'
      : `🌫️ The fog is too thick: ${e.message}`), h('button', { class: 'btn ghost', onclick: refresh }, 'Try again'));
  }
}
(async function main() {
  try {
    if (DEMO) { api = demoApi(); S.me = { name: api.me.name, ids: [api.me.id] }; }
    else {
      if (typeof msal === 'undefined') return splash('Could not load the Microsoft sign-in library. Check your network.');
      if (!await auth()) return splash();
      api = graphApi();
      S.me = { name: account.name, ids: [account.username.toLowerCase(), account.localAccountId] };
    }
    if (ADMIN || DEMO) await initAdmin();   // demo: you are the keeper too, so reveals happen on reload
    await refresh();
    setInterval(() => { if ($('#sheet').hidden) render(); }, 30e3);
    setInterval(refresh, 5 * 60e3);
    document.addEventListener('visibilitychange', () => document.visibilityState === 'visible' && refresh());
  } catch (e) { console.error(e); splash(`Sign-in failed: ${e.message}`); }
})();
