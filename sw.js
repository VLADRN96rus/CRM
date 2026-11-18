// CRM отдела v2 — серверная часть (Apps Script + Google Таблица).
// Этап 1: персональные коды, роли, фильтрация данных НА СЕРВЕРЕ, построчное хранилище, журнал действий.
//
// ПЕРВЫЙ ЗАПУСК (один раз, из редактора): migrate()
//   — создаёт вас администратором; ваш прежний код доступа остаётся рабочим, устройства перенастраивать не надо;
//   — проставляет владельца старым задачам и расходам.
// Новый сотрудник: впишите имя в NEW_USER ниже → запустите createUser() → в «Журнале выполнения» код и готовая ссылка-ключ.
// Потеря телефона / увольнение: впишите id в TARGET_ID → removeUser() (отключить) или resetCode() (новый код). Список: listUsers().
// После правки: Развёртывание → Управление → ✎ → Новая версия → Развернуть (адрес /exec прежний).

const SHEET = 'data';
const PREFIXES = ['t_', 'd_', 'mp_', 'r_', 'o_'];   // личные задачи, задачи отдела, медиаплан, постоянные и разовые расходы
const PER_USER = ['budget', 'prefs'];               // личные настройки: в таблице хранятся как budget@<id сотрудника>
const SHARED_SINGLE = ['dict'];                     // общие справочники (Фокус, Тактика)
const ADMIN_ID = 'u_admin', ADMIN_NAME = 'Владислав';
const NEW_USER = { name: 'Имя Фамилия', role: 'user' };   // role: 'user' | 'admin'
const TARGET_ID = '';

const P_ = () => PropertiesService.getScriptProperties();

// ---------- сотрудники и коды ----------
function users_() { try { return JSON.parse(P_().getProperty('USERS') || '{}'); } catch (e) { return {}; } }
function saveUsers_(u) { P_().setProperty('USERS', JSON.stringify(u)); }
function hash_(code) {
  const salt = P_().getProperty('SALT') || '';
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + String(code))
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}
function newCode_() { return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').slice(0, 32); }
function auth_(k) {
  if (!k || String(k).length < 16) return null;
  const h = hash_(k), u = users_();
  for (const id in u) if (u[id].hash === h && !u[id].off) return { id: id, name: u[id].name, role: u[id].role };
  return null;
}
function setup() {
  const p = P_();
  p.setProperty('SS_ID', SpreadsheetApp.getActive().getId());
  if (!p.getProperty('SALT')) p.setProperty('SALT', Utilities.getUuid());
  sheet_();
}
function migrate() {
  setup();
  const p = P_(), u = users_();
  if (!u[ADMIN_ID]) {
    const old = p.getProperty('TOKEN');
    const code = old || newCode_();
    u[ADMIN_ID] = { name: ADMIN_NAME, role: 'admin', hash: hash_(code), created: Date.now() };
    saveUsers_(u);
    Logger.log(old ? 'Администратор создан. Ваш прежний код доступа продолжает работать.' : 'Администратор создан. Код: ' + code);
  }
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {   // одноразово: владелец старым записям
    const s = sheet_(), n = s.getLastRow(); if (n < 2) return;
    const v = s.getRange(2, 1, n - 1, 2).getValues(); let ch = 0;
    v.forEach(r => {
      let id = String(r[0]), it; try { it = JSON.parse(r[1]); } catch (e) { return; }
      if (PER_USER.indexOf(id) >= 0) { r[0] = id + '@' + ADMIN_ID; ch++; }
      else if (['t_', 'r_', 'o_'].indexOf(pref_(id)) >= 0 && !it.owner) { it.owner = ADMIN_ID; r[1] = JSON.stringify(it); ch++; }
    });
    if (ch) s.getRange(2, 1, v.length, 2).setValues(v);
    Logger.log('Записей обновлено: ' + ch + '. Старое свойство TOKEN можно удалить: Параметры проекта → Свойства скрипта.');
  } finally { lock.releaseLock(); }
}
function createUser() {
  setup();
  const u = users_(), id = 'u' + Utilities.getUuid().replace(/-/g, '').slice(0, 8), code = newCode_();
  u[id] = { name: NEW_USER.name, role: NEW_USER.role === 'admin' ? 'admin' : 'user', hash: hash_(code), created: Date.now() };
  saveUsers_(u);
  Logger.log('Сотрудник: ' + NEW_USER.name + ' (id ' + id + ')\nКод доступа: ' + code +
    '\nСсылка для входа: <адрес вашего сайта>/index.html#u=<адрес /exec>&k=' + code + '\nКод показан один раз — передайте лично.');
}
function resetCode() {
  const u = users_(); if (!u[TARGET_ID]) return Logger.log('Нет такого id');
  const code = newCode_(); u[TARGET_ID].hash = hash_(code); u[TARGET_ID].off = false; saveUsers_(u);
  Logger.log('Новый код для ' + u[TARGET_ID].name + ': ' + code);
}
function removeUser() {
  const u = users_(); if (!u[TARGET_ID]) return Logger.log('Нет такого id');
  if (TARGET_ID === ADMIN_ID) return Logger.log('Администратора отключать нельзя');
  u[TARGET_ID].off = true; saveUsers_(u); Logger.log('Доступ отключён: ' + u[TARGET_ID].name);
}
function listUsers() { const u = users_(); Object.keys(u).forEach(id => Logger.log(id + ' · ' + u[id].name + ' · ' + u[id].role + (u[id].off ? ' · ОТКЛЮЧЁН' : ''))); }
function pubUsers_() { const u = users_(), o = []; Object.keys(u).forEach(id => { if (!u[id].off) o.push({ id: id, name: u[id].name }); }); return o; }

// ---------- API ----------
function doGet() { return out_({ ok: true, hint: 'CRM API' }); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad_request' }); }
  const user = auth_(req.k);
  if (!user) return out_({ ok: false, error: 'auth' });
  if (req.act) return mail_(req, user);

  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const s = sheet_(), n = s.getLastRow(), rows = {}, data = {};
    if (n > 1) s.getRange(2, 1, n - 1, 2).getValues().forEach((r, i) => { if (r[0]) { rows[r[0]] = i + 2; try { data[r[0]] = JSON.parse(r[1]); } catch (x) {} } });
    const upd = {}, add = {}, del = [], aud = [], rej = [], now = Date.now();
    (req.ops || []).forEach(op => {
      if (!op || !valid_(op.id)) return;
      const key = key_(op.id, user), cur = data[key], pf = pref_(op.id);
      if (op.t === 'd') {
        if (!cur || !allow_(op.id, cur, user)) return;
        delete data[key]; delete upd[key];
        if (add[key]) delete add[key]; else del.push(rows[key]);
        if (pf === 'd_' || pf === 'mp_') aud.push([new Date(), user.name, 'delete', op.id, String(cur.title || '').slice(0, 80)]);
      } else if (op.t === 'p' && op.d && typeof op.d === 'object') {
        const d = Object.assign({}, op.d); delete d.owner; delete d.ver; delete d.by; delete d.upd;   // служебные поля задаёт только сервер
        let nv;
        if (cur) {
          if (!allow_(op.id, cur, user)) return;
          nv = Object.assign({}, cur, d); nv.ver = (cur.ver || 0) + 1;
        } else if (op.c) {
          nv = d; nv.ver = 1; if (SHARED_SINGLE.indexOf(op.id) < 0 && PER_USER.indexOf(op.id) < 0) nv.owner = user.id;
        } else return;
        if (pf) { nv.by = user.id; nv.upd = now; }
        if (JSON.stringify(nv).length > 48000) { rej.push(op.id); return; }
        data[key] = nv;
        if (cur) { if (!add[key]) upd[key] = 1; } else add[key] = 1;
        if (pf === 'd_' || pf === 'mp_') aud.push([new Date(), user.name, cur ? 'edit' : 'create', op.id, Object.keys(d).join(',').slice(0, 120)]);
      }
    });
    Object.keys(upd).forEach(k => { if (data[k] && rows[k]) s.getRange(rows[k], 2).setValue(JSON.stringify(data[k])); });
    const ak = Object.keys(add).filter(k => data[k]);
    if (ak.length) { ensure_(s, s.getLastRow(), ak.length); s.getRange(s.getLastRow() + 1, 1, ak.length, 2).setValues(ak.map(k => [k, JSON.stringify(data[k])])); }
    del.sort((a, b) => b - a).forEach(r => s.deleteRow(r));
    if (aud.length) { const a = tab_('audit'); ensure_(a, a.getLastRow(), aud.length); a.getRange(a.getLastRow() + 1, 1, aud.length, 5).setValues(aud); }
    return out_({ ok: true, items: view_(data, user), me: user, users: pubUsers_(), rej: rej });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally { lock.releaseLock(); }
}

function pref_(id) { for (let i = 0; i < PREFIXES.length; i++) if (id.indexOf(PREFIXES[i]) === 0) return PREFIXES[i]; return ''; }
function valid_(id) { return typeof id === 'string' && id.length < 60 && id.indexOf('@') < 0 && (PER_USER.indexOf(id) >= 0 || SHARED_SINGLE.indexOf(id) >= 0 || pref_(id) !== ''); }
function key_(id, u) { return PER_USER.indexOf(id) >= 0 ? id + '@' + u.id : id; }
function allow_(id, it, u) {   // менять и удалять
  const p = pref_(id);
  if (p === 't_' || p === 'r_' || p === 'o_') return it.owner === u.id;
  return true;   // задачи отдела, медиаплан, справочники — общие
}
function view_(data, u) {      // что сотрудник вправе видеть
  const out = {};
  for (const k in data) {
    const it = data[k], at = k.indexOf('@');
    if (at > 0) { if (k.slice(at + 1) === u.id) out[k.slice(0, at)] = it; continue; }
    const p = pref_(k);
    if (p === 't_') { if (it.owner === u.id || it.shared) out[k] = it; }
    else if (p === 'r_' || p === 'o_') { if (it.owner === u.id) out[k] = it; }
    else out[k] = it;
  }
  return out;
}
function sheet_() {
  const id = P_().getProperty('SS_ID');
  if (!id) throw new Error('Сначала запустите migrate()');
  const ss = SpreadsheetApp.openById(id);
  let s = ss.getSheetByName(SHEET);
  if (!s) { s = ss.insertSheet(SHEET); s.appendRow(['id', 'json']); s.setFrozenRows(1); }
  return s;
}

// ================= ПИСЬМА: очередь отправки =================
// Письма в листе «outbox» (5-й столбец — владелец). Отправщик сотрудника (его код доступа) забирает ТОЛЬКО свои письма.
// Счётчики очереди пока общие на всех (персональные — в этапе «Письма»).
const PIECE = 45000;
const MSK = 'Europe/Moscow';
const MAIL_RE = /^[^@\s,;]+@[^@\s,;]+\.[^@\s,;]+$/;

function tab_(n) { const ss = SpreadsheetApp.openById(P_().getProperty('SS_ID')); return ss.getSheetByName(n) || ss.insertSheet(n); }
function get_(n) { const s = tab_(n), k = s.getLastRow(), o = []; if (k) s.getRange(1, 1, k, 1).getValues().forEach(r => { try { o.push(JSON.parse(r[0])); } catch (e) {} }); return o; }
function set_(n, arr) {
  const s = tab_(n); s.clear(); if (!arr.length) return;
  const v = arr.map(x => { const j = JSON.stringify(x); if (j.length > 49000) throw new Error('Шаблон слишком большой (лимит ячейки ~49 000 знаков)'); return [j]; });
  s.getRange(1, 1, v.length, 1).setValues(v);
}
function ensure_(s, last, need) { const have = s.getMaxRows(); if (have < last + need) s.insertRowsAfter(have, last + need - have); }
function day_() { return Utilities.formatDate(new Date(), MSK, 'yyyy-MM-dd'); }
const own_ = (x, u) => (x[4] || ADMIN_ID) === u.id;

function putSt_(s) { P_().setProperty('ST', JSON.stringify(s)); }
function rebuildSt_() {
  const c = { p: 0, c: 0, s: 0, e: 0, day: day_() }, s = tab_('outbox'), k = s.getLastRow();
  if (k) s.getRange(1, 1, k, 3).getValues().forEach(x => { if (x[1] === 0) { if (x[2] === 'pending') c.p++; else if (x[2] === 'claimed') c.c++; } });
  const l = tab_('mail_log'), n = l.getLastRow();
  if (n) l.getRange(Math.max(1, n - 999), 1, Math.min(n, 1000), 5).getValues().forEach(x => { if (Utilities.formatDate(new Date(x[0]), MSK, 'yyyy-MM-dd') === c.day) { if (x[3] === 'ok') c.s++; else c.e++; } });
  putSt_(c); return c;
}
function st_() { let s = null; try { s = JSON.parse(P_().getProperty('ST')); } catch (e) {} return (s && s.day === day_()) ? s : rebuildSt_(); }
function q_() { const s = st_(); return { p: s.p, c: s.c, s: s.s, e: s.e }; }
function cnt_(v, st) { let p = 0, c = 0; v.forEach((x, i) => { if (x[1] === 0) { if (st[i][0] === 'pending') p++; else if (st[i][0] === 'claimed') c++; } }); return { p: p, c: c }; }
function setCnt_(cn) { const s = st_(); s.p = cn.p; s.c = cn.c; putSt_(s); }

function logPush_(items, u) {
  const cache = CacheService.getScriptCache(); let a = []; try { a = JSON.parse(cache.get('LOG') || '[]'); } catch (e) {}
  items.forEach(x => a.unshift({ t: Utilities.formatDate(new Date(), MSK, 'dd.MM HH:mm'), to: x.to || '', s: x.subject || '', ok: !!x.ok, err: x.err || '', from: x.from || '', o: u.id }));
  try { cache.put('LOG', JSON.stringify(a.slice(0, 60)), 21600); } catch (e) {}
}
function recent_(m, u) {   // журнал — только свои письма
  const n = m || 15, mine = x => (x.o || ADMIN_ID) === u.id;
  try { const a = JSON.parse(CacheService.getScriptCache().get('LOG') || 'null'); if (a) { const f = a.filter(mine); if (f.length >= n) return f.slice(0, n); } } catch (e) {}
  const l = tab_('mail_log'), k = l.getLastRow(); if (!k) return [];
  return l.getRange(Math.max(1, k - 599), 1, Math.min(k, 600), 7).getValues().reverse().map(x => ({
    t: Utilities.formatDate(new Date(x[0]), MSK, 'dd.MM HH:mm'), to: x[1], s: x[2], ok: x[3] === 'ok', err: x[4], from: x[5], o: x[6] || ADMIN_ID
  })).filter(mine).slice(0, n);
}
function hb_(u) { return +P_().getProperty('HB_' + u.id) || 0; }
function hbm_(u) { return P_().getProperty('HBM_' + u.id) || ''; }

function compact_(s) {
  const k = s.getLastRow(); if (!k) return;
  const v = s.getRange(1, 1, k, 5).getValues().filter(x => x[2] === 'pending' || x[2] === 'claimed'); s.clear();
  if (v.length) s.getRange(1, 1, v.length, 5).setValues(v);
}
function purge_(s, st) {
  const live = st.filter(x => x[0] === 'pending' || x[0] === 'claimed').length;
  if (!live) { s.clear(); return; }
  let f = 0; while (f < st.length && st[f][0] === 'done') f++;
  if (f > 0) s.deleteRows(1, f);
  if (st.length - f - live > 500) compact_(s);
}
function repair_() { compact_(tab_('outbox')); return rebuildSt_(); }
function repairOutbox() { const l = LockService.getScriptLock(); l.waitLock(30000); try { Logger.log('Очередь починена: ' + JSON.stringify(repair_())); } finally { l.releaseLock(); } }

function mail_(r, u) {
  try {
    if (r.act === 'init') return out_({ ok: true, tpl: get_('tpl'), dir: [].concat.apply([], get_('dir')), q: q_(), log: recent_(15, u), hb: hb_(u), hbm: hbm_(u) });
    if (r.act === 'status') return out_({ ok: true, q: q_(), log: recent_(Math.min(Math.max(+r.n || 15, 1), 300), u), hb: hb_(u), hbm: hbm_(u) });
    if (r.act === 'tpl') { set_('tpl', r.list || []); return out_({ ok: true }); }
    if (r.act === 'dir') { set_('dir', r.list || []); return out_({ ok: true }); }
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return out_({ ok: false, error: 'busy' });
    try {
      if (r.act === 'queue') {
        const rows = [], jobs = r.jobs || [];
        jobs.forEach(j => {
          if (!MAIL_RE.test(String(j.to || '').trim())) throw new Error('Некорректная почта получателя: ' + j.to);
          String(j.cc || '').split(/[;,\s]+/).filter(Boolean).forEach(a => { if (!MAIL_RE.test(a)) throw new Error('Некорректная почта в копии: ' + a); });
          const id = 'j' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), s = JSON.stringify(j);
          for (let i = 0, k = 0; i < s.length; i += PIECE, k++) rows.push([id, k, 'pending', 'p' + s.slice(i, i + PIECE), u.id]);
        });
        const s = tab_('outbox'); if (rows.length) { ensure_(s, s.getLastRow(), rows.length); s.getRange(s.getLastRow() + 1, 1, rows.length, 5).setValues(rows); }
        const st = st_(); st.p += jobs.length; putSt_(st);
        return out_({ ok: true, queued: jobs.length });
      }
      if (r.act === 'outbox_list') {
        P_().setProperties({ ['HB_' + u.id]: String(Date.now()), ['HBM_' + u.id]: String(r.msg || '').slice(0, 200) });
        const s = tab_('outbox'), k = s.getLastRow(); if (!k) return out_({ ok: true, jobs: [] });
        const v = s.getRange(1, 1, k, 5).getValues(), by = {}, order = [];
        v.forEach((x, i) => { if (!by[x[0]]) { by[x[0]] = { first: i + 1, n: 0, st: x[2], own: own_(x, u) }; order.push(x[0]); } by[x[0]].n++; });
        const pick = order.filter(id => by[id].st === 'pending' && by[id].own).slice(0, +r.n || 10), set = {};
        const jobs = pick.map(id => {
          set[id] = 1;
          const b = by[id], d = s.getRange(b.first, 4, b.n, 1).getValues(), parts = [];
          d.forEach((row, i) => { parts[v[b.first - 1 + i][1]] = String(row[0]).slice(1); });
          return Object.assign({ id: id }, JSON.parse(parts.join('')));
        });
        if (!r.peek && pick.length) {
          pick.forEach(id => s.getRange(by[id].first, 3, by[id].n, 1).setValue('claimed'));
          setCnt_(cnt_(v, v.map(x => [set[x[0]] ? 'claimed' : x[2]])));
        }
        return out_({ ok: true, jobs: jobs });
      }
      if (r.act === 'outbox_requeue') {
        const s = tab_('outbox'), k = s.getLastRow(); let n = 0;
        if (k) {
          const v = s.getRange(1, 1, k, 5).getValues();
          const st = v.map(x => { if (x[2] === 'claimed' && own_(x, u) && (!r.id || x[0] === r.id)) { if (x[1] === 0) n++; return ['pending']; } return [x[2]]; });
          if (n) { s.getRange(1, 3, k, 1).setValues(st); setCnt_(cnt_(v, st)); }
        }
        return out_({ ok: true, n: n });
      }
      if (r.act === 'outbox_clear') {
        const s = tab_('outbox'), k = s.getLastRow(); let n = 0;
        if (k) {
          const v = s.getRange(1, 1, k, 5).getValues();
          const st = v.map(x => { if (own_(x, u) && (x[2] === 'pending' || (r.all && x[2] === 'claimed'))) { if (x[1] === 0) n++; return ['done']; } return [x[2]]; });
          s.getRange(1, 3, k, 1).setValues(st); setCnt_(cnt_(v, st)); purge_(s, st);
        }
        return out_({ ok: true, n: n });
      }
      if (r.act === 'outbox_repair') { return out_({ ok: true, st: repair_() }); }
      if (r.act === 'outbox_done') {
        let items = r.items || [r];
        const s = tab_('outbox'), k = s.getLastRow(), st0 = st_();
        if (k) {
          const v = s.getRange(1, 1, k, 5).getValues(), other = {};
          v.forEach(x => { if (!own_(x, u)) other[x[0]] = 1; });
          items = items.filter(x => !other[x.id]);   // чужие письма не трогаем
          const ids = {}; items.forEach(x => { ids[x.id] = 1; });
          const st = v.map(x => [ids[x[0]] ? 'done' : x[2]]);
          s.getRange(1, 3, k, 1).setValues(st);
          const cn = cnt_(v, st); st0.p = cn.p; st0.c = cn.c;
          purge_(s, st);
        }
        items.forEach(x => { if (x.ok) st0.s++; else st0.e++; }); putSt_(st0);
        if (items.length) {
          const l = tab_('mail_log'), rows = items.map(x => [new Date(), x.to || '', x.subject || '', x.ok ? 'ok' : 'error', x.err || '', x.from || '', u.id]);
          ensure_(l, l.getLastRow(), rows.length); l.getRange(l.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
          logPush_(items, u);
        }
        P_().setProperty('HB_' + u.id, String(Date.now()));
        return out_({ ok: true });
      }
    } finally { lock.releaseLock(); }
    return out_({ ok: false, error: 'bad_act' });
  } catch (err) { return out_({ ok: false, error: String(err && err.message || err) }); }
}
