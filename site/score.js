// Scoring is a pure function of server-stamped data, so the keeper's browser can recompute it any time.
(function (root) {
  'use strict';
  const norm = s => String(s ?? '').normalize('NFKD').replace(/[̀-ͯ]/g, '').toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ').trim().replace(/^(a|an|the) /, '');
  // key may hold alternatives: "revert|git revert"
  const match = (given, key) => norm(given) !== '' && String(key ?? '').split('|').some(k => norm(k) === norm(given));

  // levels: [{n, open, close}] ms · keys: {n: [answer, ...]} for REVEALED levels only
  // subs: [{uid, name, level, answers: JSON string, at: SharePoint lastModified ms}]
  function board(levels, keys, subs) {
    const lv = new Map(levels.map(l => [l.n, l])), latest = new Map(), names = new Map();
    for (const s of subs) {
      const l = lv.get(s.level);
      // server timestamp must sit inside the window: edits after 8 PM void the seal
      if (!l || !keys[l.n] || !(s.at >= l.open && s.at <= l.close) || String(s.answers).length > 2000) continue;
      let a;
      try { a = JSON.parse(s.answers); } catch { continue; }
      if (!Array.isArray(a)) continue;
      names.set(s.uid, s.name);
      const k = `${s.uid}|${l.n}`, prev = latest.get(k);
      if (!prev || s.at > prev.at) latest.set(k, { at: s.at, a });
    }
    const revealed = levels.filter(l => keys[l.n]).sort((x, y) => x.n - y.n);
    return [...names].map(([id, name]) => {
      let p = 0, s = 0, b = 0;
      const r = {};
      for (const l of revealed) {
        const sub = latest.get(`${id}|${l.n}`), key = keys[l.n];
        if (!sub) { s = 0; continue; }
        const c = key.filter((k, i) => match(sub.a[i], k)).length;
        let pts = 100 * c;
        if (c > 0 && c * 2 >= key.length) {       // cleared = at least half right
          s++; b = Math.max(b, s);
          pts += Math.round(50 * (l.close - sub.at) / (l.close - l.open)) + 20 * Math.min(s - 1, 5);
        } else s = 0;
        p += pts;
        r[l.n] = [c, key.length, pts];
      }
      return { id, name, p, s, b, r };
    }).sort((x, y) => y.p - x.p || x.name.localeCompare(y.name));
  }

  root.Score = { norm, match, board };
  if (typeof module === 'object') module.exports = root.Score;
})(globalThis);
