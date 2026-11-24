/**
 * 🗄️ ARCHIVO DEL HISTORIAL — archivo independiente (se conecta solo con Code.gs).
 * La hoja LOG_WEB (registro de actividad) crece unas 4.000 filas al mes. Para que el portal no se vuelva lento,
 * las filas de más de ARC_DIAS días se MUEVEN (no se borran) a la hoja LOG_WEB_ARCHIVO, en el mismo Sheet.
 *   - Nada se pierde: el archivo queda en el Sheet con las mismas columnas.
 *   - SUBSANACIONES no se archiva: la usan tiempos, seguimiento y comisiones.
 *   - Se puede hacer a mano (botón en 🩺 Estado del sistema) o automático el día 1 de cada mes, a las 3 a. m.
 * Solo la administración (rol ADMIN) lo ve y lo usa.
 */
const ARC_HOJA = 'LOG_WEB', ARC_DEST = 'LOG_WEB_ARCHIVO', ARC_DIAS = 90, ARC_MAX = 20000;
try {
  ['archivoEstado', 'archivoAhora', 'archivoAuto'].forEach(f => { if (API_PUBLICA.indexOf(f) < 0) API_PUBLICA.push(f); });
  if (SOLO_LECTURA.indexOf('archivoEstado') < 0) SOLO_LECTURA.push('archivoEstado');
} catch (e) {}

function arcAdmin_(u) { if (u.rol !== 'ADMIN') throw new Error('Solo la administración archiva el historial.'); }
const arcF_ = d => d instanceof Date ? Utilities.formatDate(d, 'America/Lima', 'dd/MM/yyyy') : '';
/** Fecha de la columna A: Date, o texto "dd/mm/aaaa hh:mm:ss". */
function arcFecha_(v) {
  if (v instanceof Date) return v;
  const m = String(v || '').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  return m ? new Date(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0)) : null;
}
/** Cuántas filas del inicio (desde la fila 2) son más antiguas que el límite. LOG_WEB se escribe en orden, así que
 *  se cuenta hasta la primera fila reciente. Por seguridad, una fila con fecha que no se entiende también detiene la cuenta. */
function arcContar_(h, limite) {
  const n = h.getLastRow(); if (n < 2) return { filas: 0, viejas: 0, desde: null, hasta: null };
  const a = h.getRange(2, 1, n - 1, 1).getValues();
  let k = 0, desde = null, hasta = null;
  for (; k < a.length && k < ARC_MAX; k++) {
    const d = arcFecha_(a[k][0]);
    if (!d || isNaN(d) || d >= limite) break;
    if (!desde) desde = d; hasta = d;
  }
  return { filas: a.length, viejas: k, desde: desde, hasta: hasta, primera: arcFecha_(a[0][0]) };
}
/** Mueve las filas antiguas. Devuelve cuántas movió. */
function arcMover_() {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const ss = ss_(), h = ss.getSheetByName(ARC_HOJA); if (!h) return 0;
    const limite = new Date(Date.now() - ARC_DIAS * 864e5), c = arcContar_(h, limite); if (!c.viejas) return 0;
    const nc = h.getLastColumn(), filas = h.getRange(2, 1, c.viejas, nc).getValues();
    let d = ss.getSheetByName(ARC_DEST);
    if (!d) {
      d = ss.insertSheet(ARC_DEST);
      d.getRange(1, 1, 1, nc).setValues(h.getRange(1, 1, 1, nc).getValues()).setFontWeight('bold').setBackground('#0C1A3D').setFontColor('#FFFFFF');
      d.setFrozenRows(1);
      try { d.protect().setWarningOnly(true).setDescription('⚠️ Archivo del historial: solo lo escribe el sistema.'); } catch (e) {}
    }
    const ini = d.getLastRow() + 1;
    if (d.getMaxRows() < ini + filas.length - 1) d.insertRowsAfter(d.getMaxRows(), ini + filas.length - 1 - d.getMaxRows());
    d.getRange(ini, 1, filas.length, nc).setValues(filas);
    SpreadsheetApp.flush();
    // solo se borra de LOG_WEB si quedó copiado completo en el archivo
    const ok = d.getRange(ini, 1, filas.length, 1).getValues().every((r, i) => String(r[0]) === String(filas[i][0]));
    if (!ok) throw new Error('No se pudo confirmar la copia en ' + ARC_DEST + '; no se borró nada de ' + ARC_HOJA + '.');
    h.deleteRows(2, filas.length);
    return filas.length;
  } finally { lock.releaseLock(); }
}
const arcTarea_ = () => ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'archivoTick');

/** Administración: cuántas filas tiene el historial y cuántas se archivarían hoy. */
function archivoEstado(token) {
  const u = sesion_(token); arcAdmin_(u);
  const ss = ss_(), h = ss.getSheetByName(ARC_HOJA), d = ss.getSheetByName(ARC_DEST), limite = new Date(Date.now() - ARC_DIAS * 864e5);
  const c = h ? arcContar_(h, limite) : { filas: 0, viejas: 0 };
  return { dias: ARC_DIAS, filas: c.filas, viejas: c.viejas, desde: arcF_(c.desde), hasta: arcF_(c.hasta), primera: arcF_(c.primera), limite: arcF_(limite),
    archivadas: d ? Math.max(d.getLastRow() - 1, 0) : 0, auto: arcTarea_() };
}
/** Administración: archiva ahora. */
function archivoAhora(token) {
  const u = sesion_(token); arcAdmin_(u);
  const n = arcMover_();
  log_(u.nombre, 'HISTORIAL_ARCHIVADO', n + ' filas de ' + ARC_HOJA + ' a ' + ARC_DEST);
  return Object.assign(archivoEstado(token), { movidas: n });
}
/** Administración: enciende o apaga el archivo automático (día 1 de cada mes, 3 a. m.). */
function archivoAuto(token, on) {
  const u = sesion_(token); arcAdmin_(u);
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'archivoTick').forEach(t => ScriptApp.deleteTrigger(t));
  if (on) ScriptApp.newTrigger('archivoTick').timeBased().onMonthDay(1).atHour(3).inTimezone('America/Lima').create();
  log_(u.nombre, 'HISTORIAL_ARCHIVO_AUTO', on ? 'activado' : 'desactivado');
  return archivoEstado(token);
}
/** Tarea mensual. */
function archivoTick() {
  const n = arcMover_();
  if (n) try { log_('Sistema', 'HISTORIAL_ARCHIVADO', n + ' filas de ' + ARC_HOJA + ' a ' + ARC_DEST + ' (automático)'); } catch (e) {}
}
