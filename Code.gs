/**
 * CORTOCIRCUITO — API (GitHub Pages ⇄ Apps Script)
 * Hojas: Usuarios · Sesiones · Ranking · Licencias · Dispositivos · Pagos
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
      case 'duo_score': return conLock_(() => duoScore_(d));
      case 'duo_ev': return conLock_(() => duoEv_(d));
      case 'tv_crear': return conLock_(() => tvCrear_(d));
      case 'tv_estado': return tvEstado_(d);
      case 'tv_unir': return conLock_(() => tvUnir_(d));
      case 'tv_answer': return conLock_(() => tvSet_(d, 'a'));
      case 'tv_ui': return conLock_(() => tvSet_(d, 'ui'));
      case 'tv_reset': return conLock_(() => tvReset_(d));
      case 'tv_in': return conLock_(() => tvIn_(d));
      case 'lic_estado': return conLock_(() => licEstado_(d));
      case 'lic_pagado': return conLock_(() => licPagado_(d));
      case 'adm_info': return { ok: true, llave: !!PropertiesService.getScriptProperties().getProperty('ADMIN_NFC') };
      case 'adm_nfc': return conLock_(() => admNfc_(d));
      case 'adm': return conLock_(() => adm_(d));
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
const DUO_MAX_MS = 180000;     // 3 minutos de partida (por defecto)
const DUO_MODOS = { clasico: 180000, carrera: 600000, muerte: 180000, territorio: 180000, bombas: 180000, coop: 180000, cara: 300000, espejo: 120000 };
function duoMax_(st) { return DUO_MODOS[st.modo] || DUO_MAX_MS; }
const DUO_SKIP_MS = 25000;     // si el rival no juega en 25 s se le salta el turno
function duoGet_(code) { const s = CacheService.getScriptCache().get('duo_' + String(code || '').toUpperCase()); return s ? JSON.parse(s) : null; }
function duoPut_(st) { st.v = (st.v || 0) + 1; CacheService.getScriptCache().put('duo_' + st.code, JSON.stringify(st), 21600); }
function duoPub_(st) {
  return { code: st.code, v: st.v, seed: st.seed, status: st.status, turn: st.turn, turnN: st.turnN, grid: st.grid,
    move: st.move, t0: st.t0, turnAt: st.turnAt, rev: st.rev, first: st.first, abandono: st.abandono, loser: st.loser, winner: st.winner, single: !!st.single, modo: st.modo || 'clasico', meta: st.meta || 0, ev: st.ev || [], o: st.o || '', a: st.a || '',
    players: st.players.map(p => ({ n: p.n, s: p.s, h: p.h, done: !!p.done })), now: Date.now() };
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
    turn: 0, turnN: 0, grid: '', move: null, t0: 0, turnAt: 0, rev: 0, first: 0, created: Date.now(), o: String(d.o || '').slice(0, 20000), a: '', single: !!d.single,
    modo: DUO_MODOS[d.modo] ? d.modo : 'clasico', meta: [500, 1000, 2000].indexOf(parseInt(d.meta, 10)) >= 0 ? parseInt(d.meta, 10) : 1000, ev: [], evN: 0 };
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
  if (st.status === 'playing' && st.t0 && Date.now() - st.t0 > duoMax_(st) + 15000) {
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
  if (Array.isArray(d.pts2)) st.players.forEach((p, i) => { p.s += Math.max(0, Math.min(1e6, parseInt(d.pts2[i], 10) || 0)); });
  else st.players[yo].s += pts;
  st.players[yo].h = (d.hold === null || d.hold === undefined) ? null : parseInt(d.hold, 10);
  st.move = { turnN: st.turnN, by: yo, pi: parseInt(d.pi, 10), path: path, grid: grid, pts: pts, pts2: Array.isArray(d.pts2) ? d.pts2 : null, over: !!d.over, lose: !!d.lose, win: !!d.win };
  if (d.lose) st.loser = yo;
  if (d.win) st.winner = yo;
  st.turnN++; st.turn = 1 - yo; st.turnAt = Date.now();
  if (d.over || d.lose || d.win || Date.now() - st.t0 > duoMax_(st)) st.status = 'over';
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
  if (Date.now() - st.t0 > duoMax_(st)) st.status = 'over';
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
  st.seed = Math.floor(Math.random() * 2e9); st.status = 'playing'; st.grid = ''; st.move = null; st.abandono = null; st.loser = null; st.winner = null; st.ev = [];
  st.players.forEach(p => { p.s = 0; p.h = null; p.done = false; });
  st.first = 1 - (st.first || 0); st.turn = st.first; st.turnN = 0; st.t0 = Date.now(); st.turnAt = st.t0; st.rev = (st.rev || 0) + 1;
  duoPut_(st);
  return { ok: true, st: duoPub_(st) };
}

/* modos a la vez (cara a cara, espejo): cada uno manda su puntuación y si ha terminado */
function duoScore_(d) {
  const st = duoGet_(d.code);
  if (!st) return { ok: false, error: 'La partida ha caducado' };
  const yo = duoYo_(st, d.token);
  if (yo < 0 || st.status !== 'playing') return { ok: true, st: duoPub_(st) };
  st.players[yo].s = Math.max(0, Math.min(99999999, parseInt(d.s, 10) || 0));
  if (d.done) st.players[yo].done = true;
  if (d.lose) { st.loser = yo; st.status = 'over'; }
  if (st.players.every(p => p.done)) st.status = 'over';
  if (Date.now() - st.t0 > duoMax_(st) + 15000) st.status = 'over';
  duoPut_(st);
  return { ok: true, st: duoPub_(st) };
}
/* eventos (ataques) cuando no hay conexión directa entre los móviles */
function duoEv_(d) {
  const st = duoGet_(d.code);
  if (!st) return { ok: false, error: 'La partida ha caducado' };
  const yo = duoYo_(st, d.token);
  if (yo < 0 || st.status !== 'playing') return { ok: true };
  st.evN = (st.evN || 0) + 1;
  st.ev = (st.ev || []).concat([{ by: yo, n: st.evN, p: d.p }]).slice(-30);
  duoPut_(st);
  return { ok: true };
}

/* ---------- MODO TV: la TV muestra el juego y los móviles son mandos ----------
   Apps Script solo empareja (ofertas/respuestas WebRTC) y hace de puente si no hay conexión directa */
function tvGet_(code) { const s = CacheService.getScriptCache().get('tv_' + String(code || '').toUpperCase()); return s ? JSON.parse(s) : null; }
function tvPut_(st) { st.v = (st.v || 0) + 1; CacheService.getScriptCache().put('tv_' + st.code, JSON.stringify(st), 21600); }
function tvCrear_(d) {
  const L = 'ABCDEFGHJKMNPQRSTUVWXYZ';
  let code = '';
  for (let k = 0; k < 20; k++) { code = ''; for (let i = 0; i < 4; i++) code += L[Math.floor(Math.random() * L.length)]; if (!tvGet_(code)) break; }
  const offers = Array.isArray(d.offers) ? d.offers.slice(0, 2) : [''];
  const st = { code, n: offers.length, slots: offers.map(o => ({ o: String(o || '').slice(0, 20000), a: '', n: '', ui: null, inp: [], inN: 0 })) };
  tvPut_(st);
  return { ok: true, code, v: st.v };
}
function tvEstado_(d) {
  const st = tvGet_(d.code);
  if (!st) return { ok: false, error: 'Esa partida de TV no existe o ha caducado' };
  if (d.v && Number(d.v) >= st.v) return { ok: true, same: true };
  return { ok: true, st: st };
}
function tvUnir_(d) {
  const st = tvGet_(d.code);
  if (!st) return { ok: false, error: 'Esa partida de TV no existe o ha caducado' };
  let k = st.slots.findIndex(x => !x.n);
  if (k < 0) return { ok: false, error: 'La TV ya tiene todos los mandos conectados' };
  st.slots[k].n = limpiaNombre_(d.nombre) || ('Jugador ' + (k + 1));
  tvPut_(st);
  return { ok: true, slot: k, o: st.slots[k].o, n: st.n };
}
function tvSet_(d, campo) {
  const st = tvGet_(d.code);
  if (!st) return { ok: false, error: 'La partida de TV ha caducado' };
  const k = parseInt(d.slot, 10);
  if (!st.slots[k]) return { ok: false, error: 'Mando no válido' };
  st.slots[k][campo] = campo === 'a' ? String(d.a || '').slice(0, 20000) : d.ui;
  tvPut_(st);
  return { ok: true };
}
function tvReset_(d) {
  const st = tvGet_(d.code);
  if (!st) return { ok: false, error: 'La partida de TV ha caducado' };
  const k = parseInt(d.slot, 10);
  if (!st.slots[k]) return { ok: false };
  st.slots[k] = { o: String(d.o || '').slice(0, 20000), a: '', n: '', ui: null, inp: [], inN: st.slots[k].inN || 0 };
  tvPut_(st);
  return { ok: true };
}
function tvIn_(d) {
  const st = tvGet_(d.code);
  if (!st) return { ok: false, error: 'La partida de TV ha caducado' };
  const k = parseInt(d.slot, 10), sl = st.slots[k];
  if (!sl) return { ok: false };
  sl.inN = (sl.inN || 0) + 1;
  sl.inp = (sl.inp || []).concat([{ n: sl.inN, m: d.m }]).slice(-40);
  tvPut_(st);
  return { ok: true };
}


/* =========================================================
   LICENCIAS · prueba gratis → pago único (myPOS)
   Hojas: Licencias · Dispositivos · Pagos
   ========================================================= */
const ADMIN_EMAIL = 'futurmovil.com@gmail.com';
const H_LIC = ['clave', 'tipo', 'email', 'nombre', 'inicio', 'estado', 'intentos', 'pagoAt', 'ref', 'extra', 'dids', 'notas'];
const H_DEV = ['did', 'fp', 'alta', 'ultimo', 'emails', 'bloq', 'pagado', 'ua', 'visitas'];
const H_PAG = ['fecha', 'ref', 'clave', 'email', 'nombre', 'did', 'intento', 'estado', 'resuelto'];
const shLic_ = () => hoja_('Licencias', H_LIC);
const shDev_ = () => hoja_('Dispositivos', H_DEV);
const shPag_ = () => hoja_('Pagos', H_PAG);
const DIA_MS = 86400000;

function cfg_() {
  let c = {};
  try { c = JSON.parse(PropertiesService.getScriptProperties().getProperty('CFG') || '{}'); } catch (e) {}
  return { dias: Number(c.dias) > 0 ? Number(c.dias) : 7, precio: Number(c.precio) > 0 ? Number(c.precio) : 0.99, on: c.on !== false };
}
function filas_(sh, cab) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, cab.length).getValues().map((r, i) => { const o = { _fila: i + 2 }; cab.forEach((k, j) => o[k] = r[j]); return o; });
}
function setFila_(sh, cab, o, campos) {
  Object.keys(campos).forEach(k => { const c = cab.indexOf(k) + 1; if (c > 0) { sh.getRange(o._fila, c).setValue(campos[k]); o[k] = campos[k]; } });
}
const ms_ = v => v instanceof Date ? v.getTime() : (Number(v) || (v ? new Date(v).getTime() : 0));
const lista_ = v => String(v || '').split(',').map(x => x.trim()).filter(Boolean);
const refDe_ = clave => 'CC-' + sha_('ref' + clave).slice(0, 5).toUpperCase();

function devGet_(did, fp, ua) {
  const sh = shDev_(), rows = filas_(sh, H_DEV), now = new Date();
  did = String(did || '').replace(/[^\w-]/g, '').slice(0, 64);
  fp = String(fp || '').replace(/[^\w]/g, '').slice(0, 64);
  let d = did ? rows.find(r => r.did === did) : null;
  // si ha borrado sus datos, lo reconocemos por la huella del móvil
  if (!d && fp) { const m = rows.filter(r => r.fp === fp); if (m.length) d = m.sort((a, b) => ms_(a.alta) - ms_(b.alta))[0]; }
  if (!d) {
    const nd = did.length >= 16 ? did : Utilities.getUuid();
    sh.appendRow([nd, fp, now, now, '', '', '', String(ua || '').slice(0, 120), 1]);
    return filas_(sh, H_DEV).find(r => r.did === nd);
  }
  setFila_(sh, H_DEV, d, { ultimo: now, fp: fp || d.fp, visitas: (Number(d.visitas) || 0) + 1 });
  return d;
}
function licGet_(clave, u, crear) {
  const sh = shLic_();
  let l = filas_(sh, H_LIC).find(r => String(r.clave) === clave);
  if (!l && crear) {
    sh.appendRow([clave, u ? 'cuenta' : 'móvil', u ? u.email : '', u ? u.nombre : 'Sin cuenta', new Date(), '', 0, '', refDe_(clave), 0, '', '']);
    l = filas_(sh, H_LIC).find(r => String(r.clave) === clave);
  }
  return l;
}
// resuelve el estado de una licencia en un móvil
function licResolver_(l, dev, cfg) {
  const now = Date.now(), shL = shLic_();
  // pago avisado sin validar en 24 h → no confirmado
  if (l.estado === 'pendiente' && now - ms_(l.pagoAt) > DIA_MS) {
    const n = (Number(l.intentos) || 0) + 1;
    setFila_(shL, H_LIC, l, { estado: n >= 3 ? 'bloqueado' : 'rechazado', intentos: n });
    pagMarcar_(l.clave, n >= 3 ? 'no confirmado · bloqueado' : 'no confirmado');
    if (n >= 3) licBloquearDevs_(l, true);
  }
  const inicio = Math.min(ms_(l.inicio) || now, ms_(dev && dev.alta) || now);
  const total = cfg.dias + (Number(l.extra) || 0);
  const dias = Math.max(0, Math.ceil((inicio + total * DIA_MS - now) / DIA_MS));
  const pagado = l.estado === 'pagado' || l.estado === 'regalado' || (dev && dev.pagado === 'sí');
  let st;
  if (!cfg.on || pagado) st = 'ok';
  else if (l.estado === 'bloqueado' || (dev && dev.bloq === 'sí')) st = 'bloqueado';
  else if (l.estado === 'pendiente') st = 'pendiente';
  else if (l.estado === 'rechazado') st = 'rechazado';
  else if (dias > 0) st = 'prueba';
  else st = 'pago';
  return { st, dias, intentos: Number(l.intentos) || 0, ref: l.ref || refDe_(l.clave), precio: cfg.precio };
}
function licBloquearDevs_(l, bloq) {
  const sh = shDev_(), rows = filas_(sh, H_DEV), ids = lista_(l.dids);
  rows.filter(r => ids.indexOf(r.did) >= 0).forEach(r => setFila_(sh, H_DEV, r, { bloq: bloq ? 'sí' : '' }));
}
function pagMarcar_(clave, estado) {
  const sh = shPag_(), rows = filas_(sh, H_PAG).filter(r => String(r.clave) === clave && !r.resuelto);
  rows.forEach(r => setFila_(sh, H_PAG, r, { estado, resuelto: new Date() }));
}
function licCtx_(d) {
  const cfg = cfg_(), dev = devGet_(d.did, d.fp, d.ua);
  const u = d.ses ? sesion_(d.ses) : null;
  const clave = u ? u.email : 'dev:' + dev.did;
  const l = licGet_(clave, u, true);
  // enlazar móvil ⇄ licencia
  const ids = lista_(l.dids); if (ids.indexOf(dev.did) < 0) { ids.push(dev.did); setFila_(shLic_(), H_LIC, l, { dids: ids.slice(-20).join(',') }); }
  if (u) { const em = lista_(dev.emails); if (em.indexOf(u.email) < 0) { em.push(u.email); setFila_(shDev_(), H_DEV, dev, { emails: em.slice(-10).join(',') }); } if (l.nombre !== u.nombre) setFila_(shLic_(), H_LIC, l, { nombre: u.nombre }); }
  return { cfg, dev, u, l };
}
function licEstado_(d) {
  const c = licCtx_(d);
  return Object.assign({ ok: true, did: c.dev.did }, licResolver_(c.l, c.dev, c.cfg));
}
function licPagado_(d) {
  const c = licCtx_(d);
  const r = licResolver_(c.l, c.dev, c.cfg);
  if (r.st !== 'pago' && r.st !== 'rechazado') return Object.assign({ ok: true, did: c.dev.did }, r);
  const now = new Date(), intento = (Number(c.l.intentos) || 0) + 1;
  setFila_(shLic_(), H_LIC, c.l, { estado: 'pendiente', pagoAt: now });
  shPag_().appendRow([now, r.ref, c.l.clave, c.l.email, c.l.nombre, c.dev.did, intento, 'pendiente', '']);
  try {
    MailApp.sendEmail({ to: ADMIN_EMAIL, subject: '💶 CORTOCIRCUITO — pago por validar ' + r.ref + ' (' + c.cfg.precio.toFixed(2).replace('.', ',') + ' €)',
      htmlBody: '<div style="font-family:Arial;background:#05070f;color:#dfe8ff;padding:22px;border-radius:12px"><h2 style="color:#38f9ff;margin:0 0 8px">Pago por validar</h2>' +
        '<p><b>' + (c.l.nombre || 'Sin cuenta') + '</b> ' + (c.l.email || '(sin cuenta)') + '</p><p>Referencia: <b style="color:#ffe14d">' + r.ref + '</b> · intento ' + intento + ' de 3</p>' +
        '<p>Tienes 24 horas para validarlo en el modo admin del juego (toca 10 veces el título).</p></div>', name: 'CORTOCIRCUITO' });
  } catch (e) {}
  return Object.assign({ ok: true, did: c.dev.did }, licResolver_(c.l, c.dev, c.cfg));
}

/* ---------- admin con llave NFC ---------- */
function admNfc_(d) {
  const props = PropertiesService.getScriptProperties(), cache = CacheService.getScriptCache();
  const fails = Number(cache.get('adm_fails') || 0);
  if (fails >= 5) return { ok: false, error: 'Demasiados intentos. Espera 15 minutos.' };
  const h = String(d.h || '');
  if (h.length < 32) return { ok: false, error: 'Llave no válida' };
  let sal = props.getProperty('ADMIN_SAL'); if (!sal) { sal = Utilities.getUuid(); props.setProperty('ADMIN_SAL', sal); }
  const guardada = props.getProperty('ADMIN_NFC');
  if (!guardada) {
    if (!d.registrar) return { ok: false, error: 'No hay llave registrada' };
    props.setProperty('ADMIN_NFC', sha_(sal + h));
  } else if (sha_(sal + h) !== guardada) {
    cache.put('adm_fails', String(fails + 1), 900);
    return { ok: false, error: 'Esa llave no es la del administrador' };
  }
  cache.remove('adm_fails');
  // el móvil del admin nunca paga
  if (d.did) { const sh = shDev_(), dv = filas_(sh, H_DEV).find(r => r.did === String(d.did)); if (dv) setFila_(sh, H_DEV, dv, { pagado: 'sí' }); }
  const tok = Utilities.getUuid() + Utilities.getUuid();
  cache.put('adm_' + tok, '1', 21600);
  return { ok: true, tok, nueva: !guardada };
}
function adm_(d) {
  const cache = CacheService.getScriptCache();
  if (!d.tok || !cache.get('adm_' + d.tok)) return { ok: false, error: 'ADMIN' };
  cache.put('adm_' + d.tok, '1', 21600);
  const cfg = cfg_(), op = d.op, shL = shLic_();
  const lic = () => filas_(shL, H_LIC), devs = () => filas_(shDev_(), H_DEV);
  const findL = () => lic().find(r => String(r.clave) === String(d.clave));
  const vista = (l, dv) => { const dev = dv.find(x => lista_(l.dids).indexOf(x.did) >= 0); const r = licResolver_(l, dev, cfg);
    return { clave: l.clave, tipo: l.tipo, email: l.email, nombre: l.nombre, inicio: ms_(l.inicio), st: l.estado === 'regalado' ? 'regalo' : r.st, dias: r.dias, intentos: r.intentos, ref: r.ref, pagoAt: ms_(l.pagoAt), moviles: lista_(l.dids).length, extra: Number(l.extra) || 0 }; };
  if (op === 'resumen' || op === 'jugadores') {
    const dv = devs(), L = lic().map(l => vista(l, dv));
    const cnt = k => L.filter(x => x.st === k).length;
    const rk = shRk_(), hoy = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    const rkRows = rk.getLastRow() > 1 ? rk.getRange(2, 1, rk.getLastRow() - 1, 2).getValues() : [];
    const partHoy = rkRows.filter(r => (r[1] instanceof Date ? Utilities.formatDate(r[1], TZ, 'yyyy-MM-dd') : String(r[1])) === hoy).length;
    return { ok: true, cfg, lista: op === 'jugadores' ? L.sort((a, b) => b.inicio - a.inicio).slice(0, 400) : undefined,
      res: { cuentas: usuarios_().length, moviles: dv.length, prueba: cnt('prueba'), pago: cnt('pago'), pendiente: cnt('pendiente'), rechazado: cnt('rechazado'), bloqueado: cnt('bloqueado'), pagados: cnt('ok'), regaladas: cnt('regalo'),
        ingresos: Math.round(cnt('ok') * cfg.precio * 100) / 100, partidas: rkRows.length, partidasHoy: partHoy, activos7: dv.filter(x => Date.now() - ms_(x.ultimo) < 7 * DIA_MS).length },
      pendientes: L.filter(x => x.st === 'pendiente').sort((a, b) => a.pagoAt - b.pagoAt) };
  }
  if (op === 'cfg') {
    const c = { dias: Math.max(1, Math.min(365, parseInt(d.dias, 10) || cfg.dias)), precio: Math.max(0.1, Math.min(99, Number(String(d.precio).replace(',', '.')) || cfg.precio)), on: d.on !== false };
    PropertiesService.getScriptProperties().setProperty('CFG', JSON.stringify(c));
    return { ok: true, cfg: c };
  }
  if (op === 'nfcBorrar') { PropertiesService.getScriptProperties().deleteProperty('ADMIN_NFC'); return { ok: true }; }
  if (op === 'ranking') {
    const sh = shRk_(), last = sh.getLastRow(); if (last < 2) return { ok: true, lista: [] };
    const ini = Math.max(2, last - 199), rows = sh.getRange(ini, 1, last - ini + 1, 9).getValues();
    return { ok: true, lista: rows.map((r, i) => ({ fila: ini + i, fecha: ms_(r[0]), nombre: r[2], puntos: r[3], email: r[8] })).sort((a, b) => b.puntos - a.puntos).slice(0, 60) };
  }
  if (op === 'rkBorrar') {
    const sh = shRk_(), f = parseInt(d.fila, 10);
    if (f >= 2 && f <= sh.getLastRow() && Number(sh.getRange(f, 4).getValue()) === Number(d.puntos)) { sh.deleteRow(f); cache.remove('top'); return { ok: true }; }
    return { ok: false, error: 'No encontrado (recarga la lista)' };
  }
  const l = findL(); if (!l) return { ok: false, error: 'Jugador no encontrado' };
  if (op === 'validar' || op === 'regalar') { setFila_(shL, H_LIC, l, { estado: op === 'regalar' ? 'regalado' : 'pagado' }); pagMarcar_(l.clave, op === 'regalar' ? 'regalado' : 'validado'); licBloquearDevs_(l, false); const sh = shDev_(); devs().filter(x => lista_(l.dids).indexOf(x.did) >= 0).forEach(x => setFila_(sh, H_DEV, x, { pagado: 'sí' })); }
  else if (op === 'rechazar') { const n = (Number(l.intentos) || 0) + 1; setFila_(shL, H_LIC, l, { estado: n >= 3 ? 'bloqueado' : 'rechazado', intentos: n }); pagMarcar_(l.clave, 'rechazado'); if (n >= 3) licBloquearDevs_(l, true); }
  else if (op === 'desbloquear') { setFila_(shL, H_LIC, l, { estado: '', intentos: 0 }); licBloquearDevs_(l, false); }
  else if (op === 'bloquear') { setFila_(shL, H_LIC, l, { estado: 'bloqueado' }); licBloquearDevs_(l, true); }
  else if (op === 'quitarPago') { setFila_(shL, H_LIC, l, { estado: '' }); const sh = shDev_(); devs().filter(x => lista_(l.dids).indexOf(x.did) >= 0).forEach(x => setFila_(sh, H_DEV, x, { pagado: '' })); }
  else if (op === 'extender') setFila_(shL, H_LIC, l, { extra: (Number(l.extra) || 0) + (parseInt(d.dias, 10) || 7) });
  else if (op === 'reiniciar') { const dv = devs().find(x => lista_(l.dids).indexOf(x.did) >= 0); const ini = Math.min(ms_(l.inicio) || Date.now(), ms_(dv && dv.alta) || Date.now()); setFila_(shL, H_LIC, l, { extra: Math.ceil((Date.now() - ini) / DIA_MS), estado: (l.estado === 'pagado' || l.estado === 'regalado') ? l.estado : '', intentos: 0 }); }
  else return { ok: false, error: 'Operación desconocida' };
  return { ok: true, j: vista(findL(), devs()) };
}
