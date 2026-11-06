/**
 * Foamies Golf · conexión con Google Sheets
 * Pegar este código en Extensiones → Apps Script de la planilla y publicarlo como "Aplicación web".
 * Crea solo las pestañas que necesita:
 *   Jugadores: nombre | indice | actualizado
 *   Jornadas:  id | nombre | cancha | modo | porcentaje | creada
 *   Grupos:    jornada | grupo | hora | data | actualizado
 *   Tarjetas:  respaldo de cada vuelta o jornada guardada, una fila por jugador con los 18 hoyos
 */
const HOJAS = {
  Jugadores: ['nombre', 'indice', 'actualizado'],
  Jornadas: ['id', 'nombre', 'cancha', 'modo', 'porcentaje', 'creada'],
  Grupos: ['jornada', 'grupo', 'hora', 'data', 'actualizado'],
  Tarjetas: ['guardada', 'clave', 'origen', 'fecha', 'nombre', 'cancha', 'modalidad', 'porcentaje', 'grupo', 'pos', 'jugador', 'indice', 'hdc', 'juega', 'salida']
    .concat(Array.from({ length: 18 }, (_, i) => 'h' + (i + 1)))
    .concat(['out', 'in', 'gross', 'neto', 'neto_par', 'hoyos']),
};

/** Ejecuta esta función una vez desde el editor (botón Ejecutar): crea las pestañas y pide los permisos. */
function configurar() {
  Object.keys(HOJAS).forEach(hoja);
  Logger.log('Listo: pestañas Jugadores, Jornadas, Grupos y Tarjetas creadas.');
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

// Clave fija que siempre lleva a la jornada creada hoy
const ATAJO = 'FOAMIES';
const ZONA = 'America/Santiago';
function diaCL(d) { return Utilities.formatDate(d, ZONA, 'yyyy-MM-dd'); }
function jornadaDeHoy() {
  const hoy = diaCL(new Date());
  const r = filas(hoja('Jornadas')).filter(x => x.creada && diaCL(new Date(ms(x.creada))) === hoy)
    .sort((a, b) => ms(b.creada) - ms(a.creada));
  return r[0] || null;
}

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
        // Evita crear dos jornadas el mismo día por error (se puede forzar)
        if (!p.forzar) {
          const deHoy = jornadaDeHoy();
          if (deHoy) return { ok: false, error: 'Ya hay una jornada hoy', existe: { id: String(deHoy.id), nombre: String(deHoy.nombre) } };
        }
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
        let id = String(p.id || '').trim().toUpperCase();
        let j;
        if (id === ATAJO) {
          // FOAMIES = la jornada creada hoy (la más reciente)
          j = jornadaDeHoy();
          if (!j) return { ok: false, error: 'Todavía no se crea la jornada de hoy' };
          id = String(j.id).toUpperCase();
        } else {
          j = filas(hoja('Jornadas')).find(x => String(x.id).toUpperCase() === id);
        }
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

      // ---------- Respaldo de tarjetas ----------
      // Reemplaza las filas de la misma clave: guardar de nuevo la misma vuelta o jornada la actualiza, no la duplica.
      case 'guardarTarjeta': {
        const s = hoja('Tarjetas');
        const clave = String(p.clave || '').trim();
        if (!clave) return { ok: false, error: 'Falta la clave' };
        const v = s.getDataRange().getValues();
        for (let r = v.length - 1; r >= 1; r--) if (String(v[r][1]) === clave) s.deleteRow(r + 1);
        const x = val => (val === null || val === undefined ? '' : val);
        const ahora = new Date();
        const rows = (p.filas || []).map(f => {
          const g = Array.from({ length: 18 }, (_, i) => x(f.g && f.g[i] ? f.g[i] : ''));
          return [ahora, clave, x(p.origen), x(p.fecha), x(p.nombre), x(p.cancha), x(p.modalidad), x(p.porcentaje),
            x(f.grupo), x(f.pos), x(f.jugador), x(f.indice), x(f.hdc), x(f.juega), x(f.salida)]
            .concat(g).concat([x(f.out), x(f.in), x(f.gross), x(f.neto), x(f.netoPar), x(f.hoyos)]);
        });
        if (rows.length) s.getRange(s.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
        return { ok: true, filas: rows.length };
      }
      default:
        return { ok: false, error: 'Acción desconocida' };
    }
  } finally {
    candado.releaseLock();
  }
}
