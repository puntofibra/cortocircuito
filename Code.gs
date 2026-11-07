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
