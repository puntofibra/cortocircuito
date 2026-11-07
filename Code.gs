/**
 * CORTOCIRCUITO — API de ranking (JSONP) para el juego en GitHub Pages
 * Hoja "Ranking": fecha | dia | nombre | puntos | nivel | circuitos | combo | id
 */
const HOJA = 'Ranking';
const TZ = 'Europe/Madrid';

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

function hoja_() {
  const ss = libro_();
  let sh = ss.getSheetByName(HOJA);
  if (!sh) {
    sh = ss.insertSheet(HOJA);
    sh.appendRow(['fecha', 'dia', 'nombre', 'puntos', 'nivel', 'circuitos', 'combo', 'id']);
    sh.setFrozenRows(1);
    sh.getRange('A1:H1').setFontWeight('bold').setBackground('#0b0f1a').setFontColor('#38f9ff');
  }
  return sh;
}

function setup() { hoja_(); }

function doGet(e) {
  const p = (e && e.parameter) || {};
  const cb = (p.callback || '').replace(/[^\w.$]/g, '');
  let out;
  try {
    if (p.accion === 'guardar') out = guardar_(p);
    else if (p.accion === 'top') out = top_();
    else out = { ok: true, juego: 'CORTOCIRCUITO' };
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  const json = JSON.stringify(out);
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function limpiaNombre_(n) {
  return String(n || '').replace(/[<>"'`]/g, '').trim().slice(0, 14) || 'Anónimo';
}

function guardar_(p) {
  const puntos = Math.max(0, Math.min(99999999, parseInt(p.puntos, 10) || 0));
  if (!puntos) return { ok: false, error: 'sin puntos' };
  const id = String(p.id || '').slice(0, 40);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = hoja_();
    // anti duplicados: mismo id de partida no se guarda dos veces
    if (id) {
      const last = sh.getLastRow();
      if (last > 1) {
        const ids = sh.getRange(Math.max(2, last - 200), 8, last - Math.max(2, last - 200) + 1, 1).getValues().flat();
        if (ids.indexOf(id) >= 0) return Object.assign({ ok: true, dup: true }, top_());
      }
    }
    const now = new Date();
    sh.appendRow([
      now,
      Utilities.formatDate(now, TZ, 'yyyy-MM-dd'),
      limpiaNombre_(p.nombre),
      puntos,
      parseInt(p.nivel, 10) || 1,
      parseInt(p.circuitos, 10) || 0,
      parseInt(p.combo, 10) || 0,
      id
    ]);
    CacheService.getScriptCache().remove('top');
  } finally {
    lock.releaseLock();
  }
  return Object.assign({ ok: true }, top_());
}

function top_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('top');
  if (hit) return JSON.parse(hit);
  const sh = hoja_();
  const last = sh.getLastRow();
  const hoy = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  let rows = [];
  if (last > 1) rows = sh.getRange(2, 1, last - 1, 7).getValues();
  const map = r => ({ nombre: r[2], puntos: r[3], nivel: r[4], circuitos: r[5] });
  const mejores = (lista) => {
    const best = {};
    lista.forEach(r => {
      const k = String(r[2]).toLowerCase();
      if (!best[k] || r[3] > best[k][3]) best[k] = r;
    });
    return Object.values(best).sort((a, b) => b[3] - a[3]).slice(0, 10).map(map);
  };
  const diaStr = r => (r[1] instanceof Date) ? Utilities.formatDate(r[1], TZ, 'yyyy-MM-dd') : String(r[1]);
  const res = {
    ok: true,
    hoy: mejores(rows.filter(r => diaStr(r) === hoy)),
    siempre: mejores(rows),
    partidas: rows.length
  };
  cache.put('top', JSON.stringify(res), 60);
  return res;
}
