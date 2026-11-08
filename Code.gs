/**
 * CORTOCIRCUITO — API (GitHub Pages ⇄ Apps Script)
 * Hojas: Usuarios · Sesiones · Ranking
 * Peticiones: POST con JSON en el cuerpo {accion, ...}  (y GET ?accion=top&callback= para JSONP)
 */
const TZ = 'Europe/Madrid';
const GOOGLE_CLIENT_ID = '454055120059-nuhaop4r62urc3u036cc3grpqo71ulg7.apps.googleusercontent.com';
const H_USU = ['email', 'nombre', 'hash', 'sal', 'google', 'alta', 'ultimo', 'mejor', 'partidas', 'codigo', 'codigoExp'];
const H_SES = ['token', 'email', 'creada', 'usada'];
const H_RK = ['fecha', 'dia', 'nombre', 'puntos', 'nivel', 'circuitos', 'combo', 'id', 'email'];

/* ---------- hojas ---------- */
function libro_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (ss) return ss;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  ss = SpreadsheetApp.create('cortocircuito');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}
function hoja_(nombre, cab) {
  const ss = libro_();
  let sh = ss.getSheetByName(nombre);
  if (!sh) {
    sh = ss.insertSheet(nombre);
    sh.appendRow(cab);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, cab.length).setFontWeight('bold').setBackground('#0b0f1a').setFontColor('#38f9ff');
  } else if (sh.getLastColumn() < cab.length) {
    sh.getRange(1, 1, 1, cab.length).setValues([cab]);
  }
  return sh;
}
const shUsu_ = () => hoja_('Usuarios', H_USU);
const shSes_ = () => hoja_('Sesiones', H_SES);
const shRk_ = () => hoja_('Ranking', H_RK);
function setup() { shUsu_(); shSes_(); shRk_(); }

/* ---------- entrada ---------- */
function doPost(e) {
  let d = {};
  try { d = JSON.parse(e.postData.contents || '{}'); } catch (err) {}
  return json_(rutear_(d));
}
function doGet(e) {
  const p = (e && e.parameter) || {};
  const out = rutear_(p);
  const cb = (p.callback || '').replace(/[^\w.$]/g, '');
  if (cb) return ContentService.createTextOutput(cb + '(' + JSON.stringify(out) + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return json_(out);
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function rutear_(d) {
  try {
    switch (d.accion) {
      case 'top': return top_();
      case 'registro': return conLock_(() => registro_(d));
      case 'entrar': return entrar_(d);
      case 'google': return conLock_(() => google_(d));
      case 'recuperar': return conLock_(() => recuperar_(d));
      case 'restablecer': return conLock_(() => restablecer_(d));
      case 'yo': { const u = sesion_(d.ses); return u ? { ok: true, user: pub_(u) } : { ok: false, error: 'SESION' }; }
      case 'salir': return salir_(d);
      case 'guardar': return conLock_(() => guardar_(d));
      case 'duo_crear': return conLock_(() => duoCrear_(d));
      case 'duo_unir': return conLock_(() => duoUnir_(d));
      case 'duo_estado': return duoEstado_(d);
      case 'duo_jugada': return conLock_(() => duoJugada_(d));
      case 'duo_saltar': return conLock_(() => duoSaltar_(d));
      case 'duo_salir': return conLock_(() => duoSalir_(d));
      case 'duo_revancha': return conLock_(() => duoRevancha_(d));
      default: return { ok: true, juego: 'CORTOCIRCUITO' };
    }
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  }
}
function conLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------- utilidades ---------- */
function sha_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}
const limpiaEmail_ = e => String(e || '').trim().toLowerCase();
const emailOk_ = e => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e);
const limpiaNombre_ = n => String(n || '').replace(/[<>"'`]/g, '').replace(/\s+/g, ' ').trim().slice(0, 14);

function usuarios_() {
  const sh = shUsu_(), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, H_USU.length).getValues().map((r, i) => {
    const o = { _fila: i + 2 };
    H_USU.forEach((k, j) => o[k] = r[j]);
    return o;
  });
}
function buscar_(email) { email = limpiaEmail_(email); return usuarios_().find(u => String(u.email).toLowerCase() === email) || null; }
function nombreLibre_(nombre, emailPropio) {
  const n = nombre.toLowerCase();
  return !usuarios_().some(u => String(u.nombre).toLowerCase() === n && String(u.email).toLowerCase() !== emailPropio);
}
function setCampos_(u, campos) {
  const sh = shUsu_();
  Object.keys(campos).forEach(k => {
    const c = H_USU.indexOf(k) + 1;
    if (c > 0) { sh.getRange(u._fila, c).setValue(campos[k]); u[k] = campos[k]; }
  });
}
function pub_(u) { return { email: u.email, nombre: u.nombre, mejor: Number(u.mejor) || 0, partidas: Number(u.partidas) || 0, google: !!u.google }; }

function nuevaSesion_(u) {
  const token = Utilities.getUuid() + Utilities.getUuid().slice(0, 8);
  const now = new Date();
  shSes_().appendRow([token, u.email, now, now]);
  setCampos_(u, { ultimo: now });
  return { ok: true, ses: token, user: pub_(u) };
}
function sesion_(token) {
  token = String(token || '');
  if (token.length < 20) return null;
  const sh = shSes_(), last = sh.getLastRow();
  if (last < 2) return null;
  const f = sh.getRange(2, 1, last - 1, 1).createTextFinder(token).matchEntireCell(true).findNext();
  if (!f) return null;
  const email = sh.getRange(f.getRow(), 2).getValue();
  sh.getRange(f.getRow(), 4).setValue(new Date());
  return buscar_(email);
}
function salir_(d) {
  const sh = shSes_(), last = sh.getLastRow();
  if (last >= 2 && d.ses) {
    const f = sh.getRange(2, 1, last - 1, 1).createTextFinder(String(d.ses)).matchEntireCell(true).findNext();
    if (f) sh.deleteRow(f.getRow());
  }
  return { ok: true };
}

/* ---------- cuentas ---------- */
function registro_(d) {
  const email = limpiaEmail_(d.email), nombre = limpiaNombre_(d.nombre), pass = String(d.pass || '');
  if (nombre.length < 3) return { ok: false, error: 'El nombre de jugador debe tener al menos 3 letras' };
  if (!emailOk_(email)) return { ok: false, error: 'Ese correo no es válido' };
  if (pass.length < 6) return { ok: false, error: 'La contraseña debe tener al menos 6 caracteres' };
  if (buscar_(email)) return { ok: false, error: 'Ya hay una cuenta con ese correo. Entra o recupera la contraseña' };
  if (!nombreLibre_(nombre, email)) return { ok: false, error: 'Ese nombre de jugador ya está cogido. Prueba otro' };
  const sal = Utilities.getUuid();
  shUsu_().appendRow([email, nombre, sha_(sal + pass), sal, '', new Date(), new Date(), 0, 0, '', '']);
  return nuevaSesion_(buscar_(email));
}
function entrar_(d) {
  const u = buscar_(d.email);
  if (!u) return { ok: false, error: 'No hay ninguna cuenta con ese correo' };
  if (!u.hash) return { ok: false, error: 'Esa cuenta se creó con Google. Pulsa “Continuar con Google” o recupera la contraseña para crear una' };
  if (sha_(u.sal + String(d.pass || '')) !== u.hash) return { ok: false, error: 'Contraseña incorrecta' };
  return nuevaSesion_(u);
}
function google_(d) {
  const r = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(d.credential || ''), { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) return { ok: false, error: 'No se pudo comprobar tu cuenta de Google' };
  const t = JSON.parse(r.getContentText());
  if (t.aud !== GOOGLE_CLIENT_ID || String(t.email_verified) !== 'true') return { ok: false, error: 'Cuenta de Google no válida' };
  const email = limpiaEmail_(t.email);
  let u = buscar_(email);
  if (!u) {
    let base = limpiaNombre_(t.given_name || t.name || email.split('@')[0]);
    if (base.length < 3) base = (base + 'Jugador').slice(0, 14);
    let nombre = base, i = 2;
    while (!nombreLibre_(nombre, email)) nombre = base.slice(0, 11) + (i++);
    shUsu_().appendRow([email, nombre, '', '', 'sí', new Date(), new Date(), 0, 0, '', '']);
    u = buscar_(email);
  } else if (!u.google) setCampos_(u, { google: 'sí' });
  return nuevaSesion_(u);
}
function recuperar_(d) {
  const u = buscar_(d.email);
  if (!u) return { ok: false, error: 'No hay ninguna cuenta con ese correo' };
  const codigo = String(Math.floor(100000 + Math.random() * 900000));
  setCampos_(u, { codigo: sha_(codigo), codigoExp: Date.now() + 30 * 60 * 1000 });
  MailApp.sendEmail({
    to: u.email,
    subject: '⚡ CORTOCIRCUITO — tu código: ' + codigo,
    htmlBody: '<div style="font-family:Arial,sans-serif;background:#05070f;color:#dfe8ff;padding:28px;border-radius:14px;text-align:center">' +
      '<h2 style="color:#38f9ff;margin:0 0 10px">CORTOCIRCUITO</h2><p>Hola ' + u.nombre + ', tu código para poner una contraseña nueva es:</p>' +
      '<div style="font-size:34px;letter-spacing:8px;font-weight:bold;color:#ffe14d;margin:16px 0">' + codigo + '</div>' +
      '<p style="color:#7d8bb3">Caduca en 30 minutos. Si no lo has pedido tú, ignora este correo.</p><p style="color:#7d8bb3">Un juego de FuturMovil</p></div>',
    name: 'CORTOCIRCUITO'
  });
  return { ok: true };
}
function restablecer_(d) {
  const u = buscar_(d.email), pass = String(d.pass || '');
  if (!u) return { ok: false, error: 'No hay ninguna cuenta con ese correo' };
  if (!u.codigo || Date.now() > Number(u.codigoExp)) return { ok: false, error: 'El código ha caducado. Pide otro' };
  if (sha_(String(d.codigo || '').trim()) !== u.codigo) return { ok: false, error: 'Código incorrecto' };
  if (pass.length < 6) return { ok: false, error: 'La contraseña debe tener al menos 6 caracteres' };
  const sal = Utilities.getUuid();
  setCampos_(u, { hash: sha_(sal + pass), sal: sal, codigo: '', codigoExp: '' });
  return nuevaSesion_(u);
}

/* ---------- ranking ---------- */
function guardar_(d) {
  const u = sesion_(d.ses);
  if (!u) return { ok: false, error: 'SESION' };
  const puntos = Math.max(0, Math.min(99999999, parseInt(d.puntos, 10) || 0));
  if (!puntos) return { ok: false, error: 'sin puntos' };
  const id = String(d.id || '').slice(0, 40);
  const sh = shRk_(), last = sh.getLastRow();
  if (id && last > 1) {
    const ini = Math.max(2, last - 300);
    if (sh.getRange(ini, 8, last - ini + 1, 1).createTextFinder(id).matchEntireCell(true).findNext())
      return Object.assign(top_(), { ok: true, dup: true, user: pub_(u) });
  }
  const now = new Date();
  sh.appendRow([now, Utilities.formatDate(now, TZ, 'yyyy-MM-dd'), u.nombre, puntos,
    parseInt(d.nivel, 10) || 1, parseInt(d.circuitos, 10) || 0, parseInt(d.combo, 10) || 0, id, u.email]);
  setCampos_(u, { mejor: Math.max(Number(u.mejor) || 0, puntos), partidas: (Number(u.partidas) || 0) + 1 });
  CacheService.getScriptCache().remove('top');
  return Object.assign(top_(), { ok: true, user: pub_(u) });
}
function top_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('top');
  if (hit) return JSON.parse(hit);
  const sh = shRk_(), last = sh.getLastRow();
  const hoy = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const rows = last > 1 ? sh.getRange(2, 1, last - 1, 9).getValues() : [];
  const map = r => ({ nombre: r[2], puntos: r[3], nivel: r[4], circuitos: r[5] });
  const mejores = lista => {
    const best = {};
    lista.forEach(r => {
      const k = String(r[8] || r[2]).toLowerCase();
      if (!best[k] || r[3] > best[k][3]) best[k] = r;
    });
    return Object.values(best).sort((a, b) => b[3] - a[3]).slice(0, 10).map(map);
  };
  const diaStr = r => (r[1] instanceof Date) ? Utilities.formatDate(r[1], TZ, 'yyyy-MM-dd') : String(r[1]);
  const res = { ok: true, hoy: mejores(rows.filter(r => diaStr(r) === hoy)), siempre: mejores(rows), partidas: rows.length };
  cache.put('top', JSON.stringify(res), 60);
  return res;
}


/* ---------- modo DOBLES (por turnos, estado en CacheService) ---------- */
const DUO_MAX_MS = 180000;     // 3 minutos de partida
const DUO_SKIP_MS = 25000;     // si el rival no juega en 25 s se le salta el turno
function duoGet_(code) { const s = CacheService.getScriptCache().get('duo_' + String(code || '').toUpperCase()); return s ? JSON.parse(s) : null; }
function duoPut_(st) { st.v = (st.v || 0) + 1; CacheService.getScriptCache().put('duo_' + st.code, JSON.stringify(st), 21600); }
function duoPub_(st) {
  return { code: st.code, v: st.v, seed: st.seed, status: st.status, turn: st.turn, turnN: st.turnN, grid: st.grid,
    move: st.move, t0: st.t0, turnAt: st.turnAt, rev: st.rev, first: st.first, abandono: st.abandono, loser: st.loser, single: !!st.single, o: st.o || '', a: st.a || '',
    players: st.players.map(p => ({ n: p.n, s: p.s, h: p.h })), now: Date.now() };
}
function duoYo_(st, token) { return st.players.findIndex(p => p.t && p.t === String(token || '')); }
function duoCrear_(d) {
  const u = sesion_(d.ses);
  if (!u) return { ok: false, error: 'SESION' };
  const L = 'ABCDEFGHJKMNPQRSTUVWXYZ';
  let code = '';
  for (let k = 0; k < 20; k++) {
    code = ''; for (let i = 0; i < 4; i++) code += L[Math.floor(Math.random() * L.length)];
    if (!duoGet_(code)) break;
  }
  const token = Utilities.getUuid();
  const st = { code, seed: Math.floor(Math.random() * 2e9), status: 'waiting', players: [{ n: u.nombre, t: token, s: 0, h: null }],
    turn: 0, turnN: 0, grid: '', move: null, t0: 0, turnAt: 0, rev: 0, first: 0, created: Date.now(), o: String(d.o || '').slice(0, 20000), a: '', single: !!d.single };
  duoPut_(st);
  return { ok: true, token, me: 0, st: duoPub_(st) };
}
function duoUnir_(d) {
  const st = duoGet_(d.code);
  if (!st) return { ok: false, error: 'Esa partida no existe o ya ha caducado' };
  const yo = duoYo_(st, d.token);
  if (yo >= 0) return { ok: true, token: st.players[yo].t, me: yo, st: duoPub_(st) };
  if (st.status !== 'waiting' || st.players.length > 1) return { ok: false, error: 'Esa partida ya ha empezado' };
  const nombre = limpiaNombre_(d.nombre) || 'Invitado';
  const token = Utilities.getUuid();
  st.players.push({ n: nombre.toLowerCase() === String(st.players[0].n).toLowerCase() ? nombre.slice(0, 12) + ' 2' : nombre, t: token, s: 0, h: null });
  st.a = String(d.a || '').slice(0, 20000);
  st.status = 'playing'; st.first = Math.random() < .5 ? 0 : 1; st.turn = st.first;
  st.t0 = Date.now(); st.turnAt = st.t0;
  duoPut_(st);
  return { ok: true, token, me: 1, st: duoPub_(st) };
}
function duoEstado_(d) {
  const st = duoGet_(d.code);
  if (!st) return { ok: false, error: 'La partida ha caducado' };
  // fin por tiempo aunque nadie tire (con margen para el último turno)
  if (st.status === 'playing' && st.t0 && Date.now() - st.t0 > DUO_MAX_MS + 15000) {
    conLock_(() => { const s2 = duoGet_(d.code); if (s2 && s2.status === 'playing') { s2.status = 'over'; duoPut_(s2); } });
    return { ok: true, st: duoPub_(duoGet_(d.code)) };
  }
  if (d.v && Number(d.v) >= st.v) return { ok: true, same: true, now: Date.now() };
  return { ok: true, st: duoPub_(st) };
}
function duoJugada_(d) {
  const st = duoGet_(d.code);
  if (!st) return { ok: false, error: 'La partida ha caducado' };
  const yo = duoYo_(st, d.token);
  if (yo < 0) return { ok: false, error: 'No estás en esta partida' };
  if (st.status !== 'playing' || yo !== st.turn || Number(d.turnN) !== st.turnN) return { ok: true, ignorada: true, st: duoPub_(st) };
  const grid = String(d.grid || '').slice(0, 200);
  const pts = Math.max(0, Math.min(1e6, parseInt(d.pts, 10) || 0));
  const path = Array.isArray(d.path) ? d.path.slice(0, 150) : [];
  st.grid = grid;
  st.players[yo].s += pts;
  st.players[yo].h = (d.hold === null || d.hold === undefined) ? null : parseInt(d.hold, 10);
  st.move = { turnN: st.turnN, by: yo, pi: parseInt(d.pi, 10), path: path, grid: grid, pts: pts, over: !!d.over, lose: !!d.lose };
  if (d.lose) st.loser = yo;
  st.turnN++; st.turn = 1 - yo; st.turnAt = Date.now();
  if (d.over || d.lose || Date.now() - st.t0 > DUO_MAX_MS) st.status = 'over';
  duoPut_(st);
  return { ok: true, st: duoPub_(st) };
}
function duoSaltar_(d) {
  const st = duoGet_(d.code);
  if (!st) return { ok: false, error: 'La partida ha caducado' };
  const yo = duoYo_(st, d.token);
  if (yo < 0 || st.status !== 'playing' || yo === st.turn || Number(d.turnN) !== st.turnN) return { ok: true, st: duoPub_(st) };
  if (Date.now() - st.turnAt < DUO_SKIP_MS) return { ok: true, st: duoPub_(st) };
  st.move = { turnN: st.turnN, by: st.turn, skip: true, grid: st.grid, pts: 0 };
  st.turnN++; st.turn = yo; st.turnAt = Date.now();
  if (Date.now() - st.t0 > DUO_MAX_MS) st.status = 'over';
  duoPut_(st);
  return { ok: true, st: duoPub_(st) };
}
function duoSalir_(d) {
  const st = duoGet_(d.code);
  if (!st) return { ok: true };
  const yo = duoYo_(st, d.token);
  if (yo < 0) return { ok: true };
  if (st.status !== 'over') { st.status = 'over'; st.abandono = yo; duoPut_(st); }
  return { ok: true };
}
function duoRevancha_(d) {
  const st = duoGet_(d.code);
  if (!st) return { ok: false, error: 'La partida ha caducado' };
  const yo = duoYo_(st, d.token);
  if (yo < 0) return { ok: false, error: 'No estás en esta partida' };
  if (st.status !== 'over' || st.players.length < 2) return { ok: true, st: duoPub_(st) };
  st.seed = Math.floor(Math.random() * 2e9); st.status = 'playing'; st.grid = ''; st.move = null; st.abandono = null; st.loser = null;
  st.players.forEach(p => { p.s = 0; p.h = null; });
  st.first = 1 - (st.first || 0); st.turn = st.first; st.turnN = 0; st.t0 = Date.now(); st.turnAt = st.t0; st.rev = (st.rev || 0) + 1;
  duoPut_(st);
  return { ok: true, st: duoPub_(st) };
}
