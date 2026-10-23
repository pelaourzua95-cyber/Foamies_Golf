/**
 * Foamies Golf · conexión con Google Sheets
 * Pegar este código en Extensiones → Apps Script de la planilla y publicarlo como "Aplicación web".
 * Crea solo las pestañas que necesita:
 *   Jugadores: nombre | indice | actualizado
 *   Jornadas:  id | nombre | cancha | modo | porcentaje | creada
 *   Grupos:    jornada | grupo | hora | data | actualizado
 */
const HOJAS = {
  Jugadores: ['nombre', 'indice', 'actualizado'],
  Jornadas: ['id', 'nombre', 'cancha', 'modo', 'porcentaje', 'creada'],
  Grupos: ['jornada', 'grupo', 'hora', 'data', 'actualizado'],
};

/** Ejecuta esta función una vez desde el editor (botón Ejecutar): crea las pestañas y pide los permisos. */
function configurar() {
  Object.keys(HOJAS).forEach(hoja);
  Logger.log('Listo: pestañas Jugadores, Jornadas y Grupos creadas.');
}

function hoja(nombre) {
  if (!HOJAS[nombre]) throw new Error('Pestaña desconocida: ' + nombre);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(nombre);
  if (!s) {
    s = ss.insertSheet(nombre);
    s.appendRow(HOJAS[nombre]);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, HOJAS[nombre].length).setFontWeight('bold');
  }
  return s;
}

function filas(s) {
  const v = s.getDataRange().getValues();
  const h = v.shift() || [];
  return v.map((r, i) => {
    const o = { _fila: i + 2 };
    h.forEach((k, j) => (o[k] = r[j]));
    return o;
  });
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try { return json(atender(e.parameter || {})); } catch (err) { return json({ ok: false, error: String(err) }); }
}

function doPost(e) {
  try { return json(atender(JSON.parse((e.postData && e.postData.contents) || '{}'))); }
  catch (err) { return json({ ok: false, error: String(err) }); }
}

function ms(d) { return d ? new Date(d).getTime() : 0; }

function codigo() {
  const c = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let s = '';
  for (let i = 0; i < 5; i++) s += c[Math.floor(Math.random() * c.length)];
  return s;
}

function atender(p) {
  const candado = LockService.getScriptLock();
  candado.waitLock(15000);
  try {
    switch (p.action) {
      case 'ping':
        return { ok: true, hora: new Date().toISOString() };

      // ---------- Jugadores ----------
      case 'jugadores': {
        const r = filas(hoja('Jugadores')).filter(x => String(x.nombre).trim());
        return { ok: true, jugadores: r.map(x => ({ n: String(x.nombre), i: x.indice === '' ? null : Number(x.indice), t: ms(x.actualizado) })) };
      }
      case 'guardarJugador': {
        const s = hoja('Jugadores');
        const n = String(p.n || '').trim();
        if (!n) return { ok: false, error: 'Falta el nombre' };
        const i = p.i === null || p.i === undefined || p.i === '' ? '' : Number(p.i);
        const f = filas(s).find(x => String(x.nombre).trim().toLowerCase() === n.toLowerCase());
        const ahora = new Date();
        if (f) s.getRange(f._fila, 1, 1, 3).setValues([[n, i, ahora]]);
        else s.appendRow([n, i, ahora]);
        return { ok: true };
      }
      case 'borrarJugador': {
        const s = hoja('Jugadores');
        const n = String(p.n || '').trim().toLowerCase();
        const f = filas(s).find(x => String(x.nombre).trim().toLowerCase() === n);
        if (f) s.deleteRow(f._fila);
        return { ok: true };
      }

      // ---------- Jornadas ----------
      case 'crearJornada': {
        const s = hoja('Jornadas');
        const ids = filas(s).map(x => String(x.id));
        let id = codigo();
        while (ids.indexOf(id) >= 0) id = codigo();
        s.appendRow([id, String(p.nombre || 'Jornada'), String(p.cancha || ''), String(p.modo || 'ind'), Number(p.porcentaje || 100), new Date()]);
        return { ok: true, id: id };
      }
      case 'jornadas': {
        const r = filas(hoja('Jornadas')).sort((a, b) => ms(b.creada) - ms(a.creada)).slice(0, 20);
        return { ok: true, jornadas: r.map(x => ({ id: String(x.id), nombre: String(x.nombre), cancha: String(x.cancha), modo: String(x.modo), porcentaje: Number(x.porcentaje), creada: ms(x.creada) })) };
      }
      case 'jornada': {
        const id = String(p.id || '').trim().toUpperCase();
        const j = filas(hoja('Jornadas')).find(x => String(x.id).toUpperCase() === id);
        if (!j) return { ok: false, error: 'No existe una jornada con ese código' };
        const g = filas(hoja('Grupos')).filter(x => String(x.jornada).toUpperCase() === id).map(x => {
          let data = null;
          try { data = JSON.parse(x.data); } catch (e) {}
          return { grupo: String(x.grupo), hora: String(x.hora), data: data, t: ms(x.actualizado) };
        });
        return { ok: true, jornada: { id: id, nombre: String(j.nombre), cancha: String(j.cancha), modo: String(j.modo), porcentaje: Number(j.porcentaje), creada: ms(j.creada) }, grupos: g };
      }
      case 'guardarGrupo': {
        const s = hoja('Grupos');
        const id = String(p.jornada || '').trim().toUpperCase();
        const grupo = String(p.grupo || '').trim();
        if (!id || !grupo) return { ok: false, error: 'Faltan jornada o grupo' };
        const data = JSON.stringify(p.data || {});
        const f = filas(s).find(x => String(x.jornada).toUpperCase() === id && String(x.grupo) === grupo);
        const ahora = new Date();
        if (f) s.getRange(f._fila, 1, 1, 5).setValues([[id, grupo, String(p.hora || ''), data, ahora]]);
        else s.appendRow([id, grupo, String(p.hora || ''), data, ahora]);
        return { ok: true, t: ahora.getTime() };
      }
      default:
        return { ok: false, error: 'Acción desconocida' };
    }
  } finally {
    candado.releaseLock();
  }
}
