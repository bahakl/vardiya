/**
 * Vardiya backend (Google Apps Script project in the same Google account as the roster spreadsheet).
 * Deploy: Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone.
 * Tabs used: "eylul" (master roster), "İstekler" (written), "Yöneticiler" (managers: Ad), "Giriş kodları" (Rol, Ad, Kod; made by makeCodes()),
 * "Çalışanlar" (workers: Ad), "_db" (hidden, created automatically; app data).
 */
const TZ = 'Europe/Istanbul';
const ROSTER_TAB = 'eylul', REQ_TAB = 'İstekler', MGR_TAB = 'Yöneticiler', WRK_TAB = 'Çalışanlar', DB_TAB = '_db', CODE_TAB = 'Giriş kodları';
const FROM_DATE = '2026-10-28';
const TOKEN_DAYS = 90;

/* ---------- helpers ---------- */
const SHEET_ID = 'PASTE_YOUR_SPREADSHEET_ID_HERE';
function ss_() { return SpreadsheetApp.openById(SHEET_ID); }
function slug_(n) {
  return String(n || '').replace(/İ/g, 'i').toLowerCase().replace(/ı/g, 'i').replace(/ş/g, 's').replace(/ğ/g, 'g')
    .replace(/ü/g, 'u').replace(/ö/g, 'o').replace(/ç/g, 'c').replace(/[^a-z0-9]/g, '');
}
function pad_(n) { return (n < 10 ? '0' : '') + n; }
function secret_() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty('SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SECRET', s); }
  return s;
}
function sign_(payload) {
  const p = Utilities.base64EncodeWebSafe(Utilities.newBlob(JSON.stringify(payload)).getBytes());
  const sig = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(p, secret_()));
  return p + '.' + sig;
}
function unsign_(tok) {
  if (!tok || typeof tok !== 'string' || tok.indexOf('.') < 0) return null;
  const parts = tok.split('.'), p = parts[0], sig = parts[1];
  const good = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(p, secret_()));
  if (good !== sig) return null;
  try {
    const o = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(p)).getDataAsString());
    return o.exp > Date.now() ? o : null;
  } catch (e) { return null; }
}
function readPhones_(tab, withEmpty) {
  const sh = ss_().getSheetByName(tab); if (!sh) return [];
  const v = sh.getDataRange().getDisplayValues(), out = [];
  for (let i = 1; i < v.length; i++) {
    const name = String(v[i][0] || '').trim();
    const d = String(v[i][1] || '').replace(/\D/g, '');
    if (!name) continue;
    if (d.length >= 10) out.push({ name: name, ph: d.slice(-10) });
    else if (withEmpty) out.push({ name: name, ph: '' });
  }
  return out;
}

/* ---------- storage (hidden "_db" tab: path | json | updatedAt) ---------- */
function dbSheet_() {
  let sh = ss_().getSheetByName(DB_TAB);
  if (!sh) { sh = ss_().insertSheet(DB_TAB); sh.appendRow(['path', 'json', 'updatedAt']); sh.hideSheet(); }
  return sh;
}
function dbLoad_() {
  const sh = dbSheet_(), n = sh.getLastRow(), map = {};
  if (n > 1) {
    const v = sh.getRange(2, 1, n - 1, 3).getValues();
    v.forEach(function (r, i) { if (r[0]) map[r[0]] = { json: r[1], row: i + 2 }; });
  }
  return { sh: sh, map: map };
}
function dbGet_(path) {
  const m = dbLoad_().map[path];
  if (!m) return null;
  try { return JSON.parse(m.json); } catch (e) { return null; }
}
function dbSet_(path, obj) {
  const js = JSON.stringify(obj);
  if (js.length > 48000) throw new Error('too big');
  const L = dbLoad_(), now = new Date().toISOString();
  if (L.map[path]) L.sh.getRange(L.map[path].row, 2, 1, 2).setValues([[js, now]]);
  else L.sh.appendRow([path, js, now]);
}

/* ---------- auth ---------- */
function rateOk_() {
  const c = CacheService.getScriptCache(), k = 'fail' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHH');
  return Number(c.get(k) || 0) < 80;
}
function rateFail_() {
  const c = CacheService.getScriptCache(), k = 'fail' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHH');
  c.put(k, String(Number(c.get(k) || 0) + 1), 3600);
}
function normCode_(c) { return String(c || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); }
function readCodes_() {
  const sh = ss_().getSheetByName(CODE_TAB); if (!sh) return [];
  const v = sh.getDataRange().getDisplayValues(), out = [];
  for (let i = 1; i < v.length; i++) {
    const role = String(v[i][0] || '').trim().toLowerCase(), name = String(v[i][1] || '').trim(), code = normCode_(v[i][2]);
    if (name && code.length >= 6) out.push({ mgr: role.indexOf('y') === 0, name: name, code: code });
  }
  return out;
}
function info_() {
  const w = readPhones_(WRK_TAB, true);
  return { ok: true, names: w.map(function (x) { return x.name; }), open: !readCodes_().length };
}
function verify_(req) {
  if (!rateOk_()) return { ok: false, err: 'rate' };
  const code = normCode_(req.code), nm = String(req.name || '').trim();
  const mk = function (r, w) { return { ok: true, r: r, w: w, token: sign_({ r: r, w: w, exp: Date.now() + TOKEN_DAYS * 864e5 }) }; };
  if (code.length >= 6) {
    const hit = readCodes_().filter(function (x) { return x.code === code; })[0];
    if (hit && hit.mgr) return mk('manager', nm || hit.name);
    if (hit && (!nm || slug_(hit.name) === slug_(nm))) return mk('worker', hit.name);
    rateFail_(); return { ok: false, err: 'bad' };
  }
  const inf = info_();
  if (inf.open && nm && inf.names.some(function (n) { return slug_(n) === slug_(nm); })) return mk('worker', nm);
  rateFail_(); return { ok: false, err: 'bad' };
}
function auth_(tok) { return unsign_(tok); }

/* ---------- API ---------- */
function snap_(req) {
  const me = auth_(req.token); if (!me) return { ok: false, err: 'auth' };
  const L = dbLoad_(), docs = {};
  Object.keys(L.map).forEach(function (p) {
    const col = p.split('/')[0];
    let ok = col === 'roster' || col === 'patron' || col === 'swaps';
    if (col === 'requests') ok = me.r === 'manager' || p === 'requests/w' + slug_(me.w);
    if (ok) { try { docs[p] = JSON.parse(L.map[p].json); } catch (e) {} }
  });
  return { ok: true, me: { r: me.r, w: me.w }, docs: docs };
}
function names_(line) {
  const f = String(line || '').split('|'), out = [];
  for (let i = 1; i <= 9; i++) String(f[i] || '').split(',').forEach(function (x) { x = x.trim(); if (x) out.push(x); });
  return out.sort().join(',');
}
function write_(req) {
  const me = auth_(req.token); if (!me) return { ok: false, err: 'auth' };
  const path = String(req.path || ''), data = req.data, parts = path.split('/');
  if (parts.length !== 2 || !data || typeof data !== 'object') return { ok: false, err: 'path' };
  const col = parts[0], id = parts[1], mgr = me.r === 'manager', mine = slug_(me.w);
  const cur = dbGet_(path);
  if (col === 'patron') {
    if (id === 'status') { if (!mgr) return { ok: false, err: 'deny' }; }
    else if (id === 'overlay') {
      if (!mgr) {
        // a worker may only move names between columns on the dates of an open swap addressed to them
        const oldRows = (cur && cur.rows) || {}, base = (dbGet_('roster/current') || {}).rows || {}, newRows = data.rows || {};
        const swaps = Object.keys(dbLoad_().map).filter(function (p) { return p.indexOf('swaps/') === 0; }).map(function (p) { return dbGet_(p); });
        const dates = {};
        Object.keys(newRows).concat(Object.keys(oldRows)).forEach(function (d) { dates[d] = 1; });
        for (const d in dates) {
          const a = oldRows[d] != null ? oldRows[d] : base[d], b = newRows[d] != null ? newRows[d] : base[d];
          if (a === b) continue;
          const okSwap = swaps.some(function (s) { return s && s.status === 'open' && slug_(s.to) === mine && (s.d === d || s.d2 === d); });
          if (!okSwap || names_(a) !== names_(b)) return { ok: false, err: 'deny' };
        }
      }
    } else return { ok: false, err: 'deny' };
  } else if (col === 'swaps') {
    if (!mgr) {
      if (cur) {
        if (slug_(cur.from) !== mine && slug_(cur.to) !== mine) return { ok: false, err: 'deny' };
        if (data.from !== cur.from || data.to !== cur.to || data.id !== cur.id) return { ok: false, err: 'deny' };
      } else if (slug_(data.from) !== mine || data.id !== id) return { ok: false, err: 'deny' };
    }
  } else if (col === 'requests') {
    if (!mgr) {
      if (id !== 'w' + mine) return { ok: false, err: 'deny' };
      const en = data.entries || {};
      for (const k in en) if (slug_(en[k].worker) !== mine) return { ok: false, err: 'deny' };
    }
  } else return { ok: false, err: 'deny' };
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { dbSet_(path, data); } finally { lock.releaseLock(); }
  return { ok: true };
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doGet() { return out_(info_()); }
function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    if (req.a === 'info') return out_(info_());
    if (req.a === 'verify') return out_(verify_(req));
    if (req.a === 'snap') return out_(snap_(req));
    if (req.a === 'set') return out_(write_(req));
    if (req.a === 'readimg') return out_(readImg_(req));
    if (req.a === 'applyroster') return out_(applyRoster_(req));
    if (req.a === 'undoroster') return out_(undoRoster_(req));
    return out_({ ok: false, err: 'op' });
  } catch (err) { return out_({ ok: false, err: 'server' }); }
}

/* ---------- picture -> sheet (manager only; needs Script Property GEMINI_KEY) ---------- */
const BACKUP_TAB = 'eylul-yedek';
const COL_ICON = ['🌙', '🔋', '💻', '📌', '📱', '🌆', '💤', '🏖️', '🤒'];
const GUN_ = ['Paz', 'Pzt', 'Sal', 'Çar', 'Per', 'Cum', 'Cmt'];
function isMgr_(req) { const me = auth_(req.token); return me && me.r === 'manager' ? me : null; }
function readImg_(req) {
  if (!isMgr_(req)) return { ok: false, err: 'auth' };
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
  if (!key) return { ok: false, err: 'nokey' };
  const imgs = (req.images || []).slice(0, 3), prompt = String(req.prompt || '').slice(0, 6000);
  if (!imgs.length || !prompt) return { ok: false, err: 'bad' };
  const parts = [{ text: prompt }];
  imgs.forEach(function (b) { parts.push({ inline_data: { mime_type: 'image/jpeg', data: String(b) } }); });
  const body = JSON.stringify({ contents: [{ role: 'user', parts: parts }], generationConfig: { temperature: 0, responseMimeType: 'application/json' } });
  const models = ['gemini-2.5-pro', 'gemini-2.5-flash'];
  let err = 'ai';
  for (let i = 0; i < models.length; i++) {
    let r;
    try {
      r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + models[i] + ':generateContent',
        { method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key }, payload: body, muteHttpExceptions: true });
    } catch (e) { err = 'net'; continue; }
    const code = r.getResponseCode();
    if (code === 429) { err = 'quota'; continue; }
    if (code !== 200) { err = (code === 400 || code === 403) && i === 0 && /API key/i.test(r.getContentText()) ? 'nokey' : 'ai'; if (err === 'nokey') return { ok: false, err: err }; continue; }
    try {
      const j = JSON.parse(r.getContentText());
      let t = ((((j.candidates || [])[0] || {}).content || {}).parts || []).map(function (p) { return p.text || ''; }).join('');
      t = t.replace(/^\s*```(?:json)?/i, '').replace(/```\s*$/, '');
      return { ok: true, data: JSON.parse(t), model: models[i] };
    } catch (e) { err = 'json'; }
  }
  return { ok: false, err: err };
}
function cellText_(i, line) {
  const names = String(line || '').split(',').map(function (x) { return x.trim().replace(/\s*\(/, ' ('); }).filter(Boolean);
  return names.length ? COL_ICON[i] + ' ' + names.join(', ') : '';
}
function applyRoster_(req) {
  if (!isMgr_(req)) return { ok: false, err: 'auth' };
  const rows = req.rows || {}, dates = Object.keys(rows).sort();
  if (!dates.length || dates.length > 400) return { ok: false, err: 'bad' };
  const parsed = {};
  for (let i = 0; i < dates.length; i++) {
    const d = dates[i], p = String(rows[d]).split('|');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d) || p.length !== 10 || p[0] !== d) return { ok: false, err: 'bad' };
    parsed[d] = p.slice(1, 10).map(function (x, k) { return cellText_(k, x.slice(0, 300)); });
  }
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const ss = ss_(), sh = ss.getSheetByName(ROSTER_TAB); if (!sh) return { ok: false, err: 'nosheet' };
    const old = ss.getSheetByName(BACKUP_TAB); if (old) ss.deleteSheet(old);
    const bk = sh.copyTo(ss); bk.setName(BACKUP_TAB); bk.hideSheet();
    const disp = sh.getDataRange().getDisplayValues(), at = {}, lastCol = Math.max(11, sh.getLastColumn());
    let lastDateRow = 1;
    const order = [];
    for (let i = 1; i < disp.length; i++) {
      const m = String(disp[i][0]).trim().match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/); if (!m) continue;
      const iso = m[3] + '-' + pad_(+m[2]) + '-' + pad_(+m[1]); at[iso] = i + 1; lastDateRow = i + 1; order.push(iso);
    }
    let written = 0;
    dates.forEach(function (d) {
      let row = at[d];
      if (!row) {
        // new date: insert in order, copying the look of the row above
        const prevKeys = order.filter(function (k) { return k < d; }), after = prevKeys.length ? prevKeys[prevKeys.length - 1] : null;
        const prevRow = after ? at[after] : lastDateRow;
        const y = +d.slice(0, 4), mo = +d.slice(5, 7), dd = +d.slice(8, 10), dt = new Date(y, mo - 1, dd, 12);
        const wk = function (k) { const g = new Date(+k.slice(0, 4), +k.slice(5, 7) - 1, +k.slice(8, 10), 12).getDay(); return g === 0 || g === 6; };
        const same = order.filter(function (k) { return k !== d && wk(k) === wk(d); }).sort(), src = same.length ? Math.max.apply(null, [at[same.filter(function (k) { return k < d; }).pop() || same[0]]]) : prevRow;
        sh.insertRowAfter(prevRow); row = prevRow + 1;
        const srcRow = src > prevRow ? src + 1 : src;
        sh.getRange(srcRow, 1, 1, lastCol).copyTo(sh.getRange(row, 1, 1, lastCol), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
        sh.getRange(row, 1).setValue(dt);
        sh.getRange(row, 2).setValue(GUN_[dt.getDay()]);
        Object.keys(at).forEach(function (k) { if (at[k] >= row) at[k]++; });
        at[d] = row; order.push(d); order.sort(); lastDateRow++;
      }
      sh.getRange(row, 3, 1, 9).setValues([parsed[d]]); written++;
    });
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  const n = syncRoster_();
  return { ok: true, written: dates.length, days: n };
}
function undoRoster_(req) {
  if (!isMgr_(req)) return { ok: false, err: 'auth' };
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const ss = ss_(), sh = ss.getSheetByName(ROSTER_TAB), bk = ss.getSheetByName(BACKUP_TAB);
    if (!sh || !bk) return { ok: false, err: 'nobackup' };
    const v = bk.getDataRange().getValues();
    sh.getDataRange().clearContent();
    sh.getRange(1, 1, v.length, v[0].length).setValues(v);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  syncRoster_();
  return { ok: true };
}

/* ---------- hourly sync ---------- */
function clean_(s) {
  return String(s || '').split(',').map(function (x) {
    return x.trim().replace(/^[^\p{L}\p{N}(]+/u, '').trim().replace(/\s+\(/g, '(');
  }).filter(Boolean).join(',');
}
function syncRoster_() {
  const sh = ss_().getSheetByName(ROSTER_TAB); if (!sh) return 0;
  const v = sh.getDataRange().getDisplayValues(), rows = {};
  for (let i = 1; i < v.length; i++) {
    const m = String(v[i][0]).trim().match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/); if (!m) continue;
    const iso = m[3] + '-' + pad_(+m[2]) + '-' + pad_(+m[1]); if (iso < FROM_DATE) continue;
    const cells = []; for (let c = 2; c < 11; c++) cells.push(clean_(v[i][c]));
    if (cells.some(function (x) { return x; })) rows[iso] = iso + '|' + cells.join('|');
  }
  const cur = dbGet_('roster/current');
  if (!cur || JSON.stringify(cur.rows) !== JSON.stringify(rows))
    dbSet_('roster/current', { rows: rows, updatedAt: new Date().toISOString(), days: Object.keys(rows).length, source: 'sheet' });
  return Object.keys(rows).length;
}
const CODE_LBL = { s0: '🌙 00-08', s1: '🔋 08-16', s2: '💻 9-17', s3: '📌 11-19', s4: '📱 14-22', s5: '🌆 16-24', off: '💤 İzin' };
const LBL_FILL = { s0: '#F4B0B0', s1: '#BFE3FA', s2: '#FFF4B8', s3: '#F2D568', s4: '#CDEFC6', s5: '#FFD0D0', off: '#E3E7E3' };
function syncRequests_() {
  const now = new Date(), today = Utilities.formatDate(now, TZ, 'yyyy-MM-dd'), u = Number(Utilities.formatDate(now, TZ, 'u'));
  const add = function (iso, n) { const d = new Date(iso + 'T12:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); };
  const mon = add(today, 8 - u), days = []; for (let i = 0; i < 7; i++) days.push(add(mon, i));
  const L = dbLoad_(), best = {};
  Object.keys(L.map).filter(function (p) { return p.indexOf('requests/') === 0; }).forEach(function (p) {
    let d; try { d = JSON.parse(L.map[p].json); } catch (e) { return; }
    const list = []; Object.keys(d.entries || {}).forEach(function (k) { list.push(d.entries[k]); });
    Object.keys(d.weeks || {}).forEach(function (k) { const w = d.weeks[k]; if (w) { w.week = w.week || k; list.push(w); } });
    list.forEach(function (e) { if (e && e.week === mon && e.worker) { const k = slug_(e.worker); if (!best[k] || String(e.at) > String(best[k].at)) best[k] = e; } });
  });
  const status = (dbGet_('patron/status') || {}).items || {}, cur = dbGet_('roster/current') || { rows: {} };
  const annual = function (date, name) {
    const f = String((cur.rows || {})[date] || '').split('|'); const s = slug_(name);
    return String(f[8] || '').split(',').some(function (x) { return slug_(x.replace(/\(.*$/, '')) === s; });
  };
  const dh = ['Pzt', 'Sal', 'Çar', 'Per', 'Cum', 'Cmt', 'Paz'];
  const header = ['Çalışan'].concat(days.map(function (d, i) { return dh[i] + ' ' + d.slice(8) + '.' + d.slice(5, 7); }), ['Gönderildi', 'Durum']);
  const tr = function (a, b) { return a.worker.localeCompare(b.worker, 'tr'); };
  const rows = [header], fills = [header.map(function () { return null; })], fmtAt = function (iso) { return Utilities.formatDate(new Date(iso), TZ, 'dd.MM HH:mm'); };
  const entries = Object.keys(best).map(function (k) { return best[k]; }).sort(tr), submitted = {};
  entries.forEach(function (e) {
    submitted[slug_(e.worker)] = 1;
    const row = [e.worker], fl = [null];
    days.forEach(function (d) {
      let c = (e.days || {})[d] || '', lbl = CODE_LBL[c] || '', fill = LBL_FILL[c] || null;
      if (!c && annual(d, e.worker)) { lbl = '🏖️ Yıllık izin'; fill = '#E5CCF5'; }
      row.push(lbl); fl.push(fill);
    });
    const st = status[mon + '_' + slug_(e.worker)];
    row.push(fmtAt(e.at), st && st.reqAt === e.at ? (st.state === 'approved' ? 'Onaylandı' : st.state === 'rejected' ? 'Reddedildi' : 'Bekliyor') : 'Bekliyor');
    fl.push(null, null); rows.push(row); fills.push(fl);
  });
  const names = readPhones_(WRK_TAB, true).map(function (x) { return x.name; });
  const missing = names.filter(function (n) { return !submitted[slug_(n)]; });
  rows.push(header.map(function () { return ''; })); fills.push(header.map(function () { return null; }));
  rows.push(['Henüz göndermeyenler:'].concat(missing, Array(Math.max(0, 9 - 1 - missing.length)).fill('')));
  fills.push(rows[rows.length - 1].map(function () { return null; }));
  const w = Math.max.apply(null, rows.map(function (r) { return r.length; }));
  rows.forEach(function (r) { while (r.length < w) r.push(''); }); fills.forEach(function (r) { while (r.length < w) r.push(null); });
  const sig = JSON.stringify(rows), p = PropertiesService.getScriptProperties();
  if (p.getProperty('REQ_SIG') === sig) return;
  let sh = ss_().getSheetByName(REQ_TAB); if (!sh) sh = ss_().insertSheet(REQ_TAB);
  sh.clear();
  const rg = sh.getRange(1, 1, rows.length, w); rg.setValues(rows); rg.setBackgrounds(fills.map(function (r) { return r.map(function (x) { return x || '#FFFFFF'; }); }));
  sh.getRange(1, 1, 1, w).setFontWeight('bold');
  p.setProperty('REQ_SIG', sig);
}
function syncAll() { const n = syncRoster_(); syncRequests_(); return n; }

/** Run once from the editor: authorises the script and installs the hourly sync. */
function setup() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'syncAll') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('syncAll').timeBased().everyHours(1).create();
  dbSheet_(); secret_();
  return 'roster days: ' + syncAll();
}

/** Run from the editor: creates/fills the "Giriş kodları" tab. Existing codes are kept; clear a cell and run again to issue a new code. */
function makeCodes() {
  const A = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  const gen = function (n) {
    const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Utilities.getUuid() + Math.random());
    let o = ''; for (let i = 0; i < n; i++) o += A[(b[i] + 256) % 256 % A.length]; return o;
  };
  let sh = ss_().getSheetByName(CODE_TAB);
  if (!sh) { sh = ss_().insertSheet(CODE_TAB); sh.appendRow(['Rol', 'Ad', 'Kod']); sh.getRange(1, 1, 1, 3).setFontWeight('bold'); }
  const have = {}; const v = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < v.length; i++) if (v[i][1]) have[slug_(v[i][1]) + '|' + (String(v[i][0]).toLowerCase().indexOf('y') === 0 ? 'm' : 'w')] = i + 1;
  const want = readPhones_(WRK_TAB, true).map(function (x) { return ['Çalışan', x.name, 6]; })
    .concat(readPhones_(MGR_TAB, true).map(function (x) { return ['Yönetici', x.name, 10]; }));
  want.forEach(function (w) {
    const key = slug_(w[1]) + '|' + (w[0] === 'Yönetici' ? 'm' : 'w'), row = have[key];
    if (row) { if (!normCode_(sh.getRange(row, 3).getValue())) sh.getRange(row, 3).setValue(gen(w[2])); }
    else sh.appendRow([w[0], w[1], gen(w[2])]);
  });
  return 'codes: ' + readCodes_().length;
}

/** Run once from the editor after adding GEMINI_KEY: asks Google for the external-request permission. */
function authorize() {
  const k = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
  const r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?key=' + encodeURIComponent(k || 'x'), { muteHttpExceptions: true });
  return r.getResponseCode() === 200 ? 'Gemini key works' : 'Permission OK, but key problem: HTTP ' + r.getResponseCode();
}
