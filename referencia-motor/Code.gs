/**
 * PORTAL DE VENTAS Y SUBSANACIONES — GRUPO MG  (v2: API para GitHub Pages)
 * Este script es solo el "motor": lee y escribe en tu Google Sheet "REVISIÓN DE VENTAS - MG".
 * La pantalla (web) vive en GitHub Pages y le habla a este script por la URL de la aplicación web.
 *
 * PASOS: 1) pega el ID de tu Sheet abajo  2) ejecuta setup() (crea accesos y programa las tareas automáticas)
 *        3) Implementar > Nueva implementación > Aplicación web (ejecutar como: yo / acceso: cualquier persona)
 *        4) copia la URL en config.js de GitHub
 *
 * TODO EN UN SOLO ARCHIVO (Code.gs): motor + ficha de revisión CIERRE/SBS + buzón + deshacer/anular. No debe haber otros archivos .gs en el proyecto.
 * Si falta alguno de esos archivos, el portal sigue funcionando: solo esas opciones responden "no disponible".
 */
const SHEET_ID = '<ID_DEL_SHEET>';   // lo que está entre /d/ y /edit en la URL del Sheet
const HOJA_BM = 'BASE MAESTRA', HOJA_CONTROL = 'CONTROL DE REVISIONES', HOJA_CIERRE = 'CIERRE', HOJA_SBS = 'SBS', HOJA_SUB = 'SUBSANACIONES';
const HOJA_ACCESO = 'ACCESO_WEB', HOJA_LOG = 'LOG_WEB';
const CARPETA_BACKUP = 'BACKUPS REVISIÓN DE VENTAS - MG';
const PRIMERA_FILA = 5, ULTIMA_FILA = 997;
const SESION_SEG = 6 * 3600;
// Columnas de BASE MAESTRA que SÍ se ingresan (C..W). A, B y X son fórmulas: nunca se escriben.
const CAMPOS = ['fechaIngreso','cierre','proyecto','tipo','unidad','cliente','dni','nacionalidad','estadoCivil','sepBienes','dni2','cliente2','empresa','fechaSep','fechaMcv','monto','profesion','financiamiento','banco','funcionario','vendedor'];
const COL_INI = 3; // C

// Hoja del Reporte Interno: vacío = se busca sola (la pestaña que tenga la columna "Info Interna").
const HOJA_INTERNO = '';

const GLOBAL_ = (typeof globalThis !== 'undefined') ? globalThis : this;   // para encontrar las funciones de todos los archivos
const API_PUBLICA = ['login', 'opciones', 'registrarVenta', 'misExpedientes', 'detalle', 'subsanar', 'descargarBackup', 'backupDriveAhora', 'accesos', 'regenerarCodigo', 'agregarUsuario', 'guardarPermisos', 'sincronizarAhora', 'registro', 'yo', 'ventas', 'reporteInterno', 'checklistDatos', 'checklistPdf', 'editarUsuario', 'comisiones', 'quitarUsuario', 'revisar', 'calendario', 'salirSesion',
  'ficha', 'guardarFicha', 'buzonEnviar', 'buzonLista', 'buzonResponder', 'anularVenta', 'reactivarVenta', 'deshacerVenta', 'verComo', 'historial',
  'cargoDatos', 'cargosLista', 'cargoEntregar', 'cargoRecibir', 'cargoImprimir'];
// "Ver como": con esa sesión solo se puede leer
const SOLO_LECTURA = ['opciones', 'misExpedientes', 'detalle', 'ficha', 'yo', 'ventas', 'reporteInterno', 'checklistDatos', 'comisiones', 'calendario', 'registro', 'accesos', 'buzonLista', 'salirSesion', 'historial', 'cargoDatos', 'cargosLista'];
// Si existe la versión mejorada (en otro archivo), se usa esa.
const API_ALIAS = { detalle: 'detalleCompleto_' };
function respuesta_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doGet() { return respuesta_({ ok: true, servicio: 'Ventas MG', hora: new Date().toISOString() }); }
/** Todas las pantallas llaman aquí: POST con texto {fn, args}. */
function doPost(e) {
  try {
    if (!e || !e.postData || String(e.postData.contents).length > 6000000) throw new Error('Solicitud inválida.');   // tope de tamaño (fotos incluidas)
    const p = JSON.parse(e.postData.contents);
    if (!Array.isArray(p.args || [])) throw new Error('Solicitud inválida.');
    if (API_PUBLICA.indexOf(p.fn) < 0) throw new Error('Operación no permitida.');
    if (typeof (p.args || [])[0] === 'string' && SOLO_LECTURA.indexOf(p.fn) < 0) { const s = CacheService.getScriptCache().get('s_' + p.args[0]); if (s && JSON.parse(s).como) throw new Error('👁️ Modo "Ver como": solo lectura, no se guardan cambios.'); }
    const alias = API_ALIAS[p.fn], fn = (alias && typeof GLOBAL_[alias] === 'function') ? GLOBAL_[alias] : GLOBAL_[p.fn];
    if (typeof fn !== 'function') throw new Error('Opción no disponible todavía (falta un archivo en Apps Script).');
    const t0 = Date.now(), data = fn.apply(null, p.args || []);
    if (SOLO_LECTURA.indexOf(p.fn) < 0 && p.fn !== 'login') tocarDatos_();   // hubo un cambio: la caché de expedientes se renueva
    console.log(p.fn + ' · ' + (Date.now() - t0) + ' ms');   // tiempos visibles en Ejecuciones
    return respuesta_({ ok: true, data: data });
  } catch (err) { return respuesta_({ ok: false, error: String(err && err.message || err).replace(/^Error:\s*/, '') }); }
}

/* ---------- utilidades ---------- */
function ss_() { return SpreadsheetApp.openById(SHEET_ID); }
function norm_(s) { return String(s == null ? '' : s).replace(/\s+/g, ' ').trim().toUpperCase(); }
function txt_(v) { return (v === 0 || v == null) ? '' : (v instanceof Date ? Utilities.formatDate(v, 'America/Lima', 'dd/MM/yyyy') : String(v).trim()); }
/** Seguridad: un texto que empieza con = + - @ se guardaría como FÓRMULA (p. ej. =IMPORTXML para sacar datos). Se antepone ' para que sea solo texto. */
const seguro_ = v => (typeof v === 'string' && /^[=+\-@\t\r]/.test(v)) ? "'" + v : v;
const fila_ = a => a.map(seguro_);
function log_(user, accion, detalle, origen) { try { ss_().getSheetByName(HOJA_LOG).appendRow(fila_([new Date(), String(user), accion, String(detalle || '').slice(0, 300), origen || 'WEB'])); } catch (e) {} }
/** Varias filas al registro de una vez (historial detallado de cambios). */
function logMuchos_(user, accion, detalles, origen) { try { if (!detalles.length) return; const l = ss_().getSheetByName(HOJA_LOG), ahora = new Date(); l.getRange(l.getLastRow() + 1, 1, detalles.length, 5).setValues(detalles.map(d => fila_([ahora, String(user), accion, String(d).slice(0, 300), origen || 'WEB']))); } catch (e) {} }
/* ---------- permisos (botones que el administrador habilita o deshabilita por usuario) ---------- */
const PERMISOS = ['ver', 'venta', 'subsanar', 'checklist', 'cargo', 'revisar', 'cierre_ver', 'cierre_edit', 'sbs_ver', 'sbs_edit', 'dashboard', 'interno', 'backup', 'accesos', 'registro'];
const PERMISOS_BASE = {
  // ASISTENTE = equipo de Contratos y Cobranzas · JEFE / VENDEDOR / APOYO = área comercial (solo ver y subsanar)
  ADMIN: PERMISOS.slice(), ASISTENTE: ['ver', 'venta', 'subsanar', 'checklist', 'cargo', 'revisar', 'dashboard', 'registro'], JEFE: ['ver', 'subsanar', 'checklist', 'cargo'],
  VENDEDOR: ['ver', 'subsanar', 'checklist', 'cargo'], APOYO: ['ver', 'subsanar', 'checklist', 'cargo'], GERENTE: ['ver', 'dashboard']
};
const ROLES_MANUALES = ['ADMIN', 'ASISTENTE', 'GERENTE'];   // los que se crean a mano en Accesos (no vienen de CATÁLOGOS)
const ROLES_CATALOGO = ['VENDEDOR', 'JEFE', 'APOYO'];   // estos solo vienen de CATÁLOGOS
/** Rol que se puede crear a mano: los fijos o uno nuevo con nombre propio (ej. LEGAL). Sus permisos se eligen con ⚙️. */
function rolValido_(rol) { return ROLES_MANUALES.indexOf(rol) >= 0 || (/^[A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ0-9 ]{1,19}$/.test(rol) && ROLES_CATALOGO.indexOf(rol) < 0); }
const ROL_ERR = 'Rol inválido: escriba solo letras (máx. 20). Vendedor, Jefe y Apoyo se crean desde CATÁLOGOS.';
function permisosDe_(rol, texto) {
  if (rol === 'ADMIN') return PERMISOS.slice();
  const t = String(texto || '').trim();
  if (!t) return (PERMISOS_BASE[rol] || ['ver']).slice();
  const ps = t.split(',').map(x => x.trim()).filter(x => PERMISOS.indexOf(x) >= 0);
  // Checklist PDF: quien ya tenía permisos elegidos antes de este interruptor lo conserva si puede vender o subsanar
  if (!/checklist/.test(t) && (ps.indexOf('venta') >= 0 || ps.indexOf('subsanar') >= 0)) ps.push('checklist');
  // Cargo de entrega: activo para quien vende, subsana o revisa (salvo que se haya apagado: sin_cargo)
  if (!/cargo/.test(t) && ['venta', 'subsanar', 'revisar'].some(p => ps.indexOf(p) >= 0)) ps.push('cargo');
  return ps;
}
function puede_(u, p) { if (u.permisos.indexOf(p) < 0) throw new Error('Su usuario no tiene habilitada esta opción. Consulte con Contratos y Cobranzas.'); }
/* Áreas de la ficha de revisión (CIERRE / SBS): ver y/o editar, por usuario en ⚙️ */
const AREA_PERMS = ['cierre_ver', 'cierre_edit', 'sbs_ver', 'sbs_edit'];
const sinTilde_ = s => norm_(s).normalize('NFD').replace(/[\u0300-\u036f]/g, '');
/** Si en ⚙️ nunca se eligieron áreas, se toman de CONTACTOS: <PERSONA_1> y <PERSONA_2> → CIERRE, <NOMBRE_ADMIN> → SBS, sin área → ambas. */
function conAreas_(rol, ps, nombre, texto, cont) {
  if (rol === 'ADMIN' || ps.indexOf('revisar') < 0 || /cierre_|sbs_|sin_area/.test(String(texto || ''))) return ps;
  const k = sinTilde_(nombre).split(' ')[0], c = (cont || contactos_()).find(x => sinTilde_(x.nombre).split(' ')[0] === k), a = c ? c.area : 'TODO';
  return ps.concat(a === 'CIERRE' ? ['cierre_ver', 'cierre_edit'] : a === 'SBS' ? ['sbs_ver', 'sbs_edit'] : AREA_PERMS);
}
function puedeArea_(u, area, modo) { const k = area.toLowerCase(); return u.permisos.indexOf(k + '_edit') >= 0 || (modo === 'ver' && u.permisos.indexOf(k + '_ver') >= 0); }

function setup() {
  const ss = ss_();
  let h = ss.getSheetByName(HOJA_ACCESO);
  if (!h) {
    h = ss.insertSheet(HOJA_ACCESO);
    h.appendRow(['Nombre', 'Correo (avisos)', 'Rol', 'Nombre en catálogo', 'Código de acceso', 'Activo (SI/NO)', 'Permisos (vacío = estándar del rol)', 'Bloqueado por admin (SI)']);
    h.appendRow(['<NOMBRE_ADMIN>', '<CORREO_ADMIN>', 'ADMIN', '', generarCodigo_(), 'SI']);
    h.setFrozenRows(1); h.getRange('A1:H1').setFontWeight('bold').setBackground('#0C1A3D').setFontColor('#FFFFFF');
    h.setColumnWidths(1, 8, 200); h.hideSheet();
  }
  if (!ss.getSheetByName(HOJA_LOG)) {
    const l = ss.insertSheet(HOJA_LOG); l.appendRow(['Fecha', 'Usuario', 'Acción', 'Detalle', 'Origen']); l.hideSheet();
  }
  const hh = ss.getSheetByName(HOJA_ACCESO); hh.getRange('G1:H1').setValues([['Permisos (vacío = estándar del rol)', 'Bloqueado por admin (SI)']]);
  const ll = ss.getSheetByName(HOJA_LOG); ll.getRange('E1').setValue('Origen');
  asegurarEquipo_();
  asegurarContactos_();
  asegurarCalendario_();
  sincronizarAccesos_();
  instalarTareas();
  Logger.log('Listo. Accesos creados desde CATÁLOGOS y tareas automáticas programadas. Ahora: Implementar > Aplicación web.');
}
/** Programa: backup semanal en Drive + actualización diaria de accesos (altas y bajas de vendedores/jefes). */
function instalarTareas() {
  ScriptApp.getProjectTriggers().forEach(t => { const f = t.getHandlerFunction(); if (['backupEnDrive', 'sincronizarAccesos_', 'registrarCambio_', 'archivarHistorial', 'registrarEstructura_'].indexOf(f) >= 0) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('registrarCambio_').forSpreadsheet(SHEET_ID).onEdit().create();
  ScriptApp.newTrigger('registrarEstructura_').forSpreadsheet(SHEET_ID).onChange().create();   // filas/columnas insertadas o eliminadas   // quién edita el Excel/Sheet directamente
  ScriptApp.newTrigger('backupEnDrive').timeBased().everyDays(7).atHour(23).create();
  ScriptApp.newTrigger('sincronizarAccesos_').timeBased().everyDays(1).atHour(6).create();
  ScriptApp.newTrigger('archivarHistorial').timeBased().onMonthDay(1).atHour(2).create();   // historial de +6 meses → HISTORIAL - GRUPO MG
}
/** AUTOMÁTICO: cada vendedor ACTIVO, cada jefe vigente y el personal de APOYO tienen acceso; al salir del catálogo se desactiva solo. Los ADMIN no se tocan. */
function sincronizarAccesos_() {
  const ss = ss_(), h = ss.getSheetByName(HOJA_ACCESO), cat = ss.getSheetByName('CATÁLOGOS'), hoy = new Date();
  const deseados = {};   // clave ROL|NOMBRE -> {nombre, rol, ref}
  cat.getRange(91, 6, 109, 3).getValues().forEach(r => {
    const n = String(r[0]).trim(), e = norm_(r[1]); if (!n) return;
    if (e === 'ACTIVO') deseados['VENDEDOR|' + norm_(n)] = { nombre: n, rol: 'VENDEDOR', ref: n };
    if (e === 'APOYO') deseados['APOYO|' + norm_(n)] = { nombre: n, rol: 'APOYO', ref: n };
  });
  cat.getRange(91, 10, 35, 4).getValues().forEach(r => {   // J:M  Jefe | Desde | Hasta | Reemplaza a
    const n = String(r[0]).trim(); if (!n) return;
    const hasta = r[2] instanceof Date ? r[2] : null;
    if (!hasta || hasta >= hoy) deseados['JEFE|' + norm_(n)] = { nombre: n, rol: 'JEFE', ref: n };
  });
  const nf = Math.max(h.getLastRow() - 1, 0), filas = nf ? h.getRange(2, 1, nf, 8).getValues() : [], vistos = {};
  filas.forEach((f, i) => {
    const rol = norm_(f[2]); if (rol === 'ADMIN' || !String(f[3]).trim()) return;   // admins y usuarios creados a mano no dependen del catálogo
    const k = rol + '|' + norm_(f[3]); vistos[k] = true;
    const quiere = (deseados[k] && norm_(f[7]) !== 'SI') ? 'SI' : 'NO';
    if (norm_(f[5]) !== quiere) h.getRange(i + 2, 6).setValue(quiere);
    if (!f[4]) h.getRange(i + 2, 5).setValue(generarCodigo_());
  });
  Object.keys(deseados).forEach(k => { if (!vistos[k]) { const d = deseados[k]; h.appendRow([d.nombre, '', d.rol, d.ref, generarCodigo_(), 'SI', '', '']); } });
}
function sincronizarAhora(token) { const u = sesion_(token); puede_(u, 'accesos'); sincronizarAccesos_(); log_(u.nombre, 'SINCRONIZAR', 'Accesos actualizados desde catálogo'); return accesos(token); }
function accesos(token) {
  const u = sesion_(token); puede_(u, 'accesos');
  const h = ss_().getSheetByName(HOJA_ACCESO), n = h.getLastRow(); if (n < 2) return { permisos: PERMISOS, lista: [] };
  const cont = contactos_();
  return { permisos: PERMISOS, lista: h.getRange(2, 1, n - 1, 8).getValues().map((f, i) => { const rol = norm_(f[2]); return { fila: i + 2, nombre: String(f[0]), correo: String(f[1]), rol: rol, ref: String(f[3]), codigo: String(f[4]), activo: norm_(f[5]) === 'SI', bloqueado: norm_(f[7]) === 'SI', permisos: conAreas_(rol, permisosDe_(rol, f[6]), f[0], f[6], cont), yo: (i + 2) === u.fila }; }) };
}
/** Habilita / deshabilita botones y el acceso de una persona. */
function guardarPermisos(token, fila, permisos, habilitado) {
  const u = sesion_(token); puede_(u, 'accesos');
  const h = ss_().getSheetByName(HOJA_ACCESO); fila = Number(fila);
  if (!(fila >= 2 && fila <= h.getLastRow())) throw new Error('Fila inválida.');
  if (fila === u.fila) throw new Error('No puede cambiar sus propios permisos.');
  const f = h.getRange(fila, 1, 1, 8).getValues()[0], rol = norm_(f[2]);
  if (rol === 'ADMIN') throw new Error('Los administradores siempre tienen todas las opciones.');
  const ps = (permisos || []).filter(x => PERMISOS.indexOf(x) >= 0);
  const txt = ps.join(',') + (ps.indexOf('revisar') >= 0 && !ps.some(p => AREA_PERMS.indexOf(p) >= 0) ? ',sin_area' : '') + (ps.indexOf('checklist') < 0 ? ',sin_checklist' : '') + (ps.indexOf('cargo') < 0 ? ',sin_cargo' : '');   // sin_checklist = apagado a propósito   // sin_area = revisa pero sin CIERRE ni SBS
  h.getRange(fila, 7).setValue(txt || '-');   // '-' = sin permisos (vacío sería el estándar del rol)
  h.getRange(fila, 8).setValue(habilitado ? '' : 'SI');
  if (habilitado && norm_(f[5]) !== 'SI') h.getRange(fila, 6).setValue('SI');
  if (!habilitado) h.getRange(fila, 6).setValue('NO');
  const antes = permisosDe_(rol, f[6]).join(',') + (norm_(f[5]) === 'SI' ? '' : ' (deshabilitado)');
  log_(u.nombre, 'PERMISOS', f[0] + ': [' + antes + '] → [' + ps.join(',') + ']' + (habilitado ? '' : ' (deshabilitado)'));
  return true;
}
function regenerarCodigo(token, fila) {
  const u = sesion_(token); puede_(u, 'accesos');
  const h = ss_().getSheetByName(HOJA_ACCESO); fila = Number(fila); if (!(fila >= 2 && fila <= h.getLastRow())) throw new Error('Fila inválida.');
  const c = generarCodigo_(); h.getRange(fila, 5).setValue(c); log_(u.nombre, 'CODIGO_NUEVO', h.getRange(fila, 1).getValue()); return c;
}
function agregarUsuario(token, nombre, correo, rol) {
  const u = sesion_(token); puede_(u, 'accesos');
  nombre = String(nombre || '').trim(); correo = String(correo || '').trim(); rol = norm_(rol);
  if (!rolValido_(rol)) throw new Error(ROL_ERR);
  if (!nombre) throw new Error('Escriba el nombre.');
  if (correo && !/^\S+@\S+\.\S+$/.test(correo)) throw new Error('Correo inválido.');
  const c = generarCodigo_(); ss_().getSheetByName(HOJA_ACCESO).appendRow(fila_([nombre, correo, rol, '', c, 'SI', '', ''])); log_(u.nombre, 'USUARIO_NUEVO', nombre + ' (' + rol + ')'); return c;
}
/** Rotación de personal: borra un usuario creado a mano (los de CATÁLOGOS se dan de baja allí). */
function quitarUsuario(token, fila) {
  const u = sesion_(token); puede_(u, 'accesos');
  const h = ss_().getSheetByName(HOJA_ACCESO); fila = Number(fila);
  if (!(fila >= 2 && fila <= h.getLastRow())) throw new Error('Fila inválida.');
  if (fila === u.fila) throw new Error('No puede quitarse a sí misma.');
  const f = h.getRange(fila, 1, 1, 8).getValues()[0];
  if (String(f[3]).trim()) throw new Error('Este usuario viene de CATÁLOGOS: dalo de baja allí o deshabilítalo con ⚙️.');
  if (norm_(f[2]) === 'ADMIN' && !h.getRange(2, 1, h.getLastRow() - 1, 6).getValues().some((r, i) => i + 2 !== fila && norm_(r[2]) === 'ADMIN' && norm_(r[5]) === 'SI')) throw new Error('Debe quedar al menos un administrador activo.');
  h.deleteRow(fila);
  log_(u.nombre, 'USUARIO_QUITADO', f[0] + ' (' + norm_(f[2]) + ')');
  return true;
}
/* ---------- calendario comercial: hoja CALENDARIO (Cierre | Fecha | Etapa | Hora límite) — se edita en el Sheet para nuevos meses ---------- */
const HOJA_CAL = 'CALENDARIO';
const ETAPAS = { E: 'ENTREGA DE FILES (COMERCIAL)', R: 'REVISIÓN DE FILES (CONTRATOS)', P: 'CONTRATOS PRESENTA REVISIÓN', S: 'COMERCIAL SUBSANA', X: 'COMERCIAL PRESENTA SUBSANADOS', V: 'VERIFICACIÓN DE SUBSANACIONES', G: 'PRESENTACIÓN A GERENCIA', F: 'FERIADO' };
const ETAPA_HORA = { E: '17:30', P: '17:30', X: '17:30' };
// Calendario recibido (CALENDARIO COMERCIAL.xlsx): cierre → "fecha:etapa" (las fechas sin año/mes heredan las anteriores)
const CAL_INICIAL = {
  'AGOSTO 2026': '2026-09-01:E 02:R 03:R 04:P 07:S 08:S 09:X 10:V 11:V 14:G 15:G',
  'SEPTIEMBRE 2026': '2026-10-01:E 02:R 05:R 06:P 07:S 08:F 09:S 12:X 13:V 14:V 15:G 16:G',
  'OCTUBRE 2026': '2026-11-02:E 03:R 04:R 05:P 06:S 09:S 10:X 11:V 12:V 13:G 16:G',
  'NOVIEMBRE 2026': '2026-12-01:E 02:R 03:R 04:P 07:S 08:F 09:F 10:S 11:X 14:V 15:V 16:G 17:G 25:F',
  'DICIEMBRE 2026': '2027-01-01:F 04:E 05:R 06:R 07:P 08:S 11:S 12:X 13:V 14:V 15:G 18:G'
};
function codEtapa_(t) { t = norm_(t); return /FERIADO/.test(t) ? 'F' : /ENTREGA/.test(t) ? 'E' : /VERIFICA/.test(t) ? 'V' : /PRESENTA SUBSAN|PRESENTA TODOS/.test(t) ? 'X' : /SUBSANA/.test(t) ? 'S' : /CONTRATOS PRESENTA|PRESENTA REVISI/.test(t) ? 'P' : /REVISI/.test(t) ? 'R' : /GERENCIA/.test(t) ? 'G' : ''; }
const MESES_ = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];
// Sheets convierte "AGOSTO 2026" en fecha: se acepta fecha, texto de fecha ("Sat Aug 01 2026 …") o el nombre del mes
function mesTxt_(v) {
  let d = Object.prototype.toString.call(v) === '[object Date]' ? v : null; const s = String(v == null ? '' : v).trim();
  if (!d && /^[A-Z][a-z]{2} [A-Z][a-z]{2} \d{1,2} \d{4}/i.test(s)) { const t = new Date(s.replace(/\s*\(.*\)\s*$/, '')); if (!isNaN(t)) d = t; }
  return d ? MESES_[d.getMonth()] + ' ' + d.getFullYear() : s;
}
function asegurarCalendario_() {
  const ss = ss_(), ya = ss.getSheetByName(HOJA_CAL);
  if (ya) {   // corrige cierres que Sheets volvió fecha
    if (ya.getLastRow() >= 2) {
      const n = ya.getLastRow() - 1, a = ya.getRange(2, 1, n, 1), v = a.getValues(); if (v.some(r => mesTxt_(r[0]) !== String(r[0]).trim())) { a.setNumberFormat('@'); a.setValues(v.map(r => [mesTxt_(r[0])])); }
      const hz = ya.getRange(2, 3, n, 2).getValues();   // horas que Sheets guardó como hora (con desfase): se reescriben como texto según la etapa
      if (hz.some(r => Object.prototype.toString.call(r[1]) === '[object Date]')) { const d = ya.getRange(2, 4, n, 1); d.setNumberFormat('@'); d.setValues(hz.map(r => [Object.prototype.toString.call(r[1]) === '[object Date]' ? (ETAPA_HORA[codEtapa_(r[0])] || '') : String(r[1] || '')])); }
    }
    return;
  }
  const h = ss.insertSheet(HOJA_CAL), filas = [];
  h.getRange(2, 1, 400, 1).setNumberFormat('@'); h.getRange(2, 4, 400, 1).setNumberFormat('@');
  Object.keys(CAL_INICIAL).forEach(c => { let ym = ''; CAL_INICIAL[c].split(' ').forEach(p => { const [f, e] = p.split(':'); if (f.length > 2) ym = f.slice(0, 8); const d = (ym + (f.length > 2 ? f.slice(8) : f)).split('-'); filas.push([c, new Date(+d[0], +d[1] - 1, +d[2]), ETAPAS[e], ETAPA_HORA[e] || '']); }); });
  h.getRange(1, 1, 1, 4).setValues([['Cierre (mes de venta)', 'Fecha', 'Etapa', 'Hora límite']]).setFontWeight('bold').setBackground('#0C1A3D').setFontColor('#FFFFFF');
  h.getRange(2, 1, filas.length, 4).setValues(filas); h.getRange(2, 2, filas.length, 1).setNumberFormat('dd/mm/yyyy');
  h.getRange(2, 3, 400, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(Object.keys(ETAPAS).map(k => ETAPAS[k]), true).build());
  h.setFrozenRows(1); h.setColumnWidths(1, 4, 230);
}
const mesNorm_ = s => norm_(s).replace('SETIEMBRE', 'SEPTIEMBRE');
function cal_() {
  const h = ss_().getSheetByName(HOJA_CAL); if (!h || h.getLastRow() < 2) return [];
  const rg = h.getRange(2, 1, h.getLastRow() - 1, 4), dv = rg.getDisplayValues();   // la hora se lee tal como se ve (evita desfases de zona horaria)
  return rg.getValues().map((r, i) => [r[0], r[1], r[2], (dv[i][3].match(/(\d{1,2}):(\d{2})/) || [''])[0]]).filter(r => r[0] && Object.prototype.toString.call(r[1]) === '[object Date]')
    .map(r => ({ cierre: mesNorm_(mesTxt_(r[0])), fecha: Utilities.formatDate(r[1], 'America/Lima', 'yyyy-MM-dd'), etapa: String(r[2]).trim(), cod: codEtapa_(r[2]), hora: r[3] }));
}
function calendario(token) { sesion_(token); return cal_(); }
/** Fecha y hora límite de una etapa (la última fecha de esa etapa en ese cierre). null si el cierre no está en el calendario. */
function plazo_(cierre, cod) {
  const l = cal_().filter(r => r.cierre === mesNorm_(cierre) && r.cod === cod).sort((a, b) => a.fecha < b.fecha ? 1 : -1)[0]; if (!l) return null;
  const d = l.fecha.split('-'), hm = (l.hora || '23:59').split(':');
  return new Date(+d[0], +d[1] - 1, +d[2], +hm[0] || 23, +hm[1] || 59);
}
const fmtPlazo_ = d => Utilities.formatDate(d, 'America/Lima', 'dd/MM/yyyy') + ' ' + Utilities.formatDate(d, 'America/Lima', 'h:mm a').toLowerCase();

/* ---------- avisos: hoja CONTACTOS (Nombre | WhatsApp | Correo | Área) define a quién avisar según el área observada ---------- */
const HOJA_CONT = 'CONTACTOS';
const AREAS_INICIALES = { '<PERSONA_1>': 'CIERRE', '<PERSONA_2>': 'CIERRE', '<NOMBRE_ADMIN>': 'SBS' };
function asegurarContactos_() {
  const ss = ss_(); let h = ss.getSheetByName(HOJA_CONT);
  if (!h) { h = ss.insertSheet(HOJA_CONT); h.appendRow(['Nombre', 'WhatsApp', 'Correo']); }
  if (!String(h.getRange(1, 4).getValue()).trim()) h.getRange(1, 4).setValue('Área (SBS / CIERRE / TODO)');
  const n = h.getLastRow(), v = n >= 2 ? h.getRange(2, 1, n - 1, 4).getValues() : [];
  v.forEach((r, i) => { if (r[0] && !String(r[3]).trim()) { const a = AREAS_INICIALES[norm_(r[0]).split(' ')[0]]; if (a) h.getRange(i + 2, 4).setValue(a); } });
  if (!v.some(r => /^<PERSONA_2>/.test(norm_(r[0])))) h.appendRow(['<PERSONA_2>', '', '', 'CIERRE']);
  const s = ss.getSheetByName(HOJA_SUB);
  if (!String(s.getRange(1, 6).getValue()).trim()) s.getRange(1, 6, 1, 3).setValues([['Área', 'Avisar a', 'Estado revisión']]).setFontWeight('bold');
  if (!String(s.getRange(1, 9).getValue()).trim()) s.getRange(1, 9).setValue('Plazo (calendario)').setFontWeight('bold');
}
function contactos_() {
  const h = ss_().getSheetByName(HOJA_CONT); if (!h || h.getLastRow() < 2) return [];
  return h.getRange(2, 1, h.getLastRow() - 1, 4).getDisplayValues().filter(r => String(r[0]).trim())
    .map(r => ({ nombre: String(r[0]).trim(), wa: String(r[1]).replace(/\D/g, ''), correo: String(r[2]).trim(), area: norm_(r[3]) || 'TODO' }));
}
/** Área de quien entra (para la campana): por su primer nombre en CONTACTOS. Admin y quien no figure ven todo. */
function miArea_(u) {
  if (u.rol === 'ADMIN') return 'TODO';
  const k = norm_(u.nombre).split(' ')[0], c = contactos_().find(c => norm_(c.nombre).split(' ')[0] === k);
  return c ? c.area : 'TODO';
}
/** Área de la subsanación: la del documento elegido; si es "General", las áreas que tienen observaciones. */
function areasDe_(ex, items, item) {
  const m = String(item || '').match(/^(SBS|CIERRE)\b/i); if (m) return [m[1].toUpperCase()];
  const a = [...new Set(items.map(i => i.area))]; if (a.length) return a;
  const r = []; if (ex.obsSbs) r.push('SBS'); if (ex.obsCtr) r.push('CIERRE');
  return r.length ? r : ['SBS', 'CIERRE'];
}
/** Equipo de Contratos y Cobranzas que setup() deja creado (si ya existen, no se duplican). */
const EQUIPO_INICIAL = [['<PERSONA_1>', 'ASISTENTE'], ['<PERSONA_2>', 'ASISTENTE']];
function asegurarEquipo_() {
  const h = ss_().getSheetByName(HOJA_ACCESO), n = h.getLastRow(), hay = n >= 2 ? h.getRange(2, 1, n - 1, 1).getValues().map(r => norm_(r[0])) : [], nuevos = [];
  EQUIPO_INICIAL.forEach(e => { if (hay.indexOf(norm_(e[0])) < 0) { h.appendRow([e[0], '', e[1], '', generarCodigo_(), 'SI', '', '']); nuevos.push(e[0]); } });
  return nuevos;
}
/** Cambia nombre, correo y (solo en usuarios creados a mano) el rol ADMIN / ASISTENTE. */
function editarUsuario(token, fila, nombre, correo, rol) {
  const u = sesion_(token); puede_(u, 'accesos');
  const h = ss_().getSheetByName(HOJA_ACCESO); fila = Number(fila);
  if (!(fila >= 2 && fila <= h.getLastRow())) throw new Error('Fila inválida.');
  nombre = String(nombre || '').trim(); correo = String(correo || '').trim(); rol = norm_(rol);
  if (!nombre) throw new Error('Escriba el nombre.');
  if (correo && !/^\S+@\S+\.\S+$/.test(correo)) throw new Error('Correo inválido.');
  const f = h.getRange(fila, 1, 1, 8).getValues()[0], antes = norm_(f[2]);
  if (String(f[3]).trim()) rol = antes;   // vendedores y jefes vienen de CATÁLOGOS: su rol no se cambia aquí
  else if (!rolValido_(rol)) throw new Error(ROL_ERR);
  if (fila === u.fila && rol !== antes) throw new Error('No puede cambiar su propio rol.');
  if (antes === 'ADMIN' && rol !== 'ADMIN') {
    const otros = h.getRange(2, 1, h.getLastRow() - 1, 6).getValues().filter((r, i) => i + 2 !== fila && norm_(r[2]) === 'ADMIN' && norm_(r[5]) === 'SI');
    if (!otros.length) throw new Error('Debe quedar al menos un administrador activo.');
  }
  h.getRange(fila, 1, 1, 3).setValues([fila_([nombre, correo, rol])]);
  log_(u.nombre, 'USUARIO_EDITADO', f[0] + ' → ' + nombre + ' · ' + (correo || 'sin correo') + ' · ' + rol);
  return true;
}
function generarCodigo_() { const c = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; let s = ''; for (let i = 0; i < 8; i++) s += c[Math.floor(Math.random() * c.length)]; return s; }
/** Ejecuta esto si quieres códigos nuevos para los usuarios que tengan la columna E vacía. */
function completarCodigos() {
  const h = ss_().getSheetByName(HOJA_ACCESO), n = h.getLastRow();
  for (let r = 2; r <= n; r++) if (h.getRange(r, 1).getValue() && !h.getRange(r, 5).getValue()) h.getRange(r, 5).setValue(generarCodigo_());
}

/* ---------- sesión ---------- */
function login(codigo) {
  const cache = CacheService.getScriptCache();
  const fallos = Number(cache.get('fallos') || 0);
  if (fallos >= 20) throw new Error('Demasiados intentos. Espere 10 minutos.');
  Utilities.sleep(700);
  const h = ss_().getSheetByName(HOJA_ACCESO), n = h.getLastRow();
  const cod = String(codigo || '').trim().toUpperCase();
  if (cod && n >= 2) {
    const v = h.getRange(2, 1, n - 1, 8).getValues();
    for (let i = 0; i < v.length; i++) {
      const f = v[i];
      if (String(f[4]).trim().toUpperCase() === cod) {
        if (norm_(f[5]) !== 'SI') { log_(String(f[0]), 'LOGIN_BLOQUEADO', 'Intento con usuario deshabilitado'); throw new Error('Su usuario está deshabilitado. Consulte con Contratos y Cobranzas.'); }
        const token = Utilities.getUuid();
        cache.put('s_' + token, JSON.stringify({ fila: i + 2, cod: cod }), SESION_SEG);
        const user = sesion_(token);
        log_(user.nombre, 'LOGIN', '');
        return { token: token, user: user };
      }
    }
  }
  cache.put('fallos', String(fallos + 1), 600);
  log_('(desconocido)', 'LOGIN_FALLIDO', 'Código incorrecto');
  if (fallos + 1 === 20) {   // posible ataque de fuerza bruta: se pausa el ingreso 10 min y se avisa a los administradores
    log_('SISTEMA', 'ALERTA_SEGURIDAD', '20 intentos fallidos en 10 minutos: ingreso pausado 10 min');
    try { MailApp.sendEmail(admins_(), '⚠️ Alerta de seguridad · Cuadro de Revisiones', 'Se registraron 20 intentos de ingreso con códigos incorrectos en 10 minutos. El ingreso quedó pausado 10 minutos.\nRevisa la pestaña Registro del portal (Intento fallido). Si se repite, genera códigos nuevos en Accesos.'); } catch (e) {}
  }
  throw new Error('Código incorrecto.');
}
/** Lee SIEMPRE la fila de ACCESO_WEB: si el admin cambia permisos o deshabilita, se aplica de inmediato. */
function sesion_(token) {
  const s = CacheService.getScriptCache().get('s_' + token);
  if (!s) throw new Error('Sesión vencida. Ingrese nuevamente.');
  const ses = JSON.parse(s), f = ss_().getSheetByName(HOJA_ACCESO).getRange(ses.fila, 1, 1, 8).getValues()[0];
  if (String(f[4]).trim().toUpperCase() !== ses.cod || norm_(f[5]) !== 'SI') throw new Error('Sesión vencida. Ingrese nuevamente.');
  const rol = norm_(f[2]);
  return { fila: ses.fila, nombre: String(f[0]), correo: String(f[1]), rol: rol, ref: String(f[3]).trim(), permisos: conAreas_(rol, permisosDe_(rol, f[6]), f[0], f[6]), como: ses.como || '' };
}
/** HISTORIAL de un expediente: todo lo del registro (portal y Excel) que menciona su código. Lo más reciente primero. */
function historial(token, id) {
  const u = sesion_(token); puede_(u, 'registro');
  id = String(id || '').trim(); if (!/^[\w-]+$/.test(id)) throw new Error('Expediente inválido.');
  if (!misExpedientes(token).some(x => x.id === id)) throw new Error('No tiene acceso a este expediente.');
  const re = new RegExp(id + '(?!\\d)'), l = ss_().getSheetByName(HOJA_LOG), n = l.getLastRow(), out = [];
  if (n >= 2) l.getRange(2, 1, n - 1, 5).getValues().forEach(r => {
    if (!(r[0] instanceof Date) || !re.test(String(r[3]))) return;
    out.push({ t: r[0].getTime(), fecha: Utilities.formatDate(r[0], 'America/Lima', 'dd/MM/yyyy HH:mm'), usuario: String(r[1]), accion: String(r[2]), detalle: String(r[3]).replace(re, '').replace(/^\s*·\s*/, '').trim(), origen: String(r[4] || 'WEB') });
  });
  return out.sort((a, b) => b.t - a.t).slice(0, 400);
}

/* ---------- protección de la estructura del Sheet ----------
 * CONTROL DE REVISIONES: hoja completa protegida (es 100% fórmulas).
 * Columna A (códigos EXP-) de BASE MAESTRA y SBS: protegida (son fórmulas; además impide que otros eliminen o inserten filas ahí).
 * Columna A de CIERRE: solo AVISO (a veces se escribe a mano cuando la venta se registra directo en el Excel).
 * La cuenta dueña (Contratos y Cobranzas) puede editar todo; el portal sigue funcionando porque escribe con esa cuenta. */
function protegerRango_(rg, desc, soloAviso) {
  const ya = rg.getSheet().getProtections(SpreadsheetApp.ProtectionType.RANGE).filter(p => p.getDescription() === desc)[0];
  const p = ya || rg.protect(); p.setDescription(desc);
  if (soloAviso) p.setWarningOnly(true); else { p.removeEditors(p.getEditors()); if (p.canDomainEdit()) p.setDomainEdit(false); }
  return p;
}
/** Disparador instalable (al cambiar la estructura): registra filas/columnas/pestañas insertadas o eliminadas. */
function registrarEstructura_(e) {
  tocarDatos_();
  try {
    const t = e && e.changeType, N = { INSERT_ROW: 'insertó fila(s)', REMOVE_ROW: 'eliminó fila(s)', INSERT_COLUMN: 'insertó columna(s)', REMOVE_COLUMN: 'eliminó columna(s)', INSERT_GRID: 'creó una pestaña', REMOVE_GRID: 'eliminó una pestaña' };
    if (!N[t]) return;
    let quien = ''; try { quien = e.user ? e.user.getEmail() : ''; } catch (x) {}
    const h = ss_().getSheetByName(HOJA_ACCESO), v = h.getRange(2, 1, Math.max(h.getLastRow() - 1, 1), 2).getValues();
    const f = v.filter(r => r[1] && quien && String(r[1]).toLowerCase() === String(quien).toLowerCase())[0];
    let donde = ''; try { const s = e.source.getActiveSheet(); donde = s.getName() + '!' + s.getActiveRange().getA1Notation() + ' · '; } catch (x) {}
    log_(f ? f[0] + ' (' + quien + ')' : (quien || '(editor no identificado)'), 'ESTRUCTURA', donde + N[t], 'EXCEL');
  } catch (err) {}
}
/** Ejecutar UNA vez: aplica las protecciones y activa el registro de filas/columnas. */
function instalarProteccion() {
  const ss = ss_(), c = ss.getSheetByName(HOJA_CONTROL), pc = c.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0] || c.protect();
  pc.setDescription('CONTROL: solo fórmulas, no se edita a mano'); pc.removeEditors(pc.getEditors()); if (pc.canDomainEdit()) pc.setDomainEdit(false);
  protegerRango_(ss.getSheetByName(HOJA_BM).getRange('A' + PRIMERA_FILA + ':A' + ULTIMA_FILA), 'Códigos de expediente (fórmula)');
  protegerRango_(ss.getSheetByName(HOJA_SBS).getRange('A6:A' + ULTIMA_FILA), 'Códigos de expediente (fórmula)');
  protegerRango_(ss.getSheetByName(HOJA_CIERRE).getRange('A5:A' + ULTIMA_FILA), 'Códigos de expediente: cambiar con cuidado', true);
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'registrarEstructura_').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('registrarEstructura_').forSpreadsheet(SHEET_ID).onChange().create();
  Logger.log('Listo: CONTROL protegida · columna A protegida en BASE MAESTRA y SBS · aviso en CIERRE · registro de filas/columnas activo.');
}

/* ---------- respaldo del historial (LOG_WEB) ----------
 * LOG_WEB queda protegida: solo la escribe el sistema (los editores del Sheet no pueden cambiarla ni borrarla).
 * Cada día 1, lo que tiene más de HIST_MESES meses se MUEVE al Sheet aparte "HISTORIAL - GRUPO MG" (una pestaña por año,
 * también protegida, en la carpeta de backups). Así el Sheet principal no crece sin límite y el archivo nunca se borra solo. */
const HIST_ARCHIVO = 'HISTORIAL - GRUPO MG', HIST_MESES = 6;
function protegerHoja_(h) {
  const p = h.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0] || h.protect();
  p.setDescription('Historial: solo lo escribe el sistema'); p.removeEditors(p.getEditors()); if (p.canDomainEdit()) p.setDomainEdit(false);
  return p;
}
function archivoHistorial_() {
  const pr = PropertiesService.getScriptProperties(), id = pr.getProperty('HIST_ID');
  if (id) try { return SpreadsheetApp.openById(id); } catch (e) {}
  const it = DriveApp.getFoldersByName(CARPETA_BACKUP), carpeta = it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA_BACKUP);
  const ss = SpreadsheetApp.create(HIST_ARCHIVO); DriveApp.getFileById(ss.getId()).moveTo(carpeta);
  const h = ss.getSheets()[0].setName('LEEME'); h.getRange('A1:A3').setValues([['Historial archivado del portal Cuadro de Revisiones (Grupo MG).'], ['Cada día 1 se agregan aquí los movimientos con más de ' + HIST_MESES + ' meses: una pestaña por año.'], ['Solo lo escribe el sistema. No editar.']]);
  protegerHoja_(h); pr.setProperty('HIST_ID', ss.getId());
  return ss;
}
/** AUTOMÁTICO (día 1 de cada mes): mueve al archivo el historial con más de HIST_MESES meses. Copia primero y solo borra si la copia quedó completa. */
function archivarHistorial() {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const l = ss_().getSheetByName(HOJA_LOG), n = l.getLastRow(); if (n < 2) return 0;
    const lim = new Date(); lim.setMonth(lim.getMonth() - HIST_MESES);
    const v = l.getRange(2, 1, n - 1, 5).getValues();
    let k = 0; while (k < v.length && v[k][0] instanceof Date && v[k][0] < lim) k++;   // el registro está en orden: se mueven las primeras filas
    if (!k) return 0;
    const dst = archivoHistorial_(), porAnio = {};
    v.slice(0, k).forEach(r => { const a = String(r[0].getFullYear()); (porAnio[a] = porAnio[a] || []).push(r); });
    let copiadas = 0;
    Object.keys(porAnio).sort().forEach(a => {
      let h = dst.getSheetByName(a);
      if (!h) { h = dst.insertSheet(a); h.getRange(1, 1, 1, 5).setValues([['Fecha', 'Usuario', 'Acción', 'Detalle', 'Origen']]).setFontWeight('bold'); h.setFrozenRows(1); protegerHoja_(h); }
      const ini = h.getLastRow() + 1; h.getRange(ini, 1, porAnio[a].length, 5).setValues(porAnio[a]); copiadas += porAnio[a].length;
    });
    SpreadsheetApp.flush();
    if (copiadas !== k) throw new Error('Copia incompleta: no se borró nada de LOG_WEB.');
    l.deleteRows(2, k);
    log_('SISTEMA', 'HISTORIAL_ARCHIVADO', k + ' movimientos anteriores al ' + Utilities.formatDate(lim, 'America/Lima', 'dd/MM/yyyy') + ' → ' + HIST_ARCHIVO, 'AUTO');
    return k;
  } finally { lock.releaseLock(); }
}
/** Ejecutar UNA vez: protege LOG_WEB, crea el Sheet de archivo y programa la tarea del día 1 de cada mes. */
function instalarHistorial() {
  protegerHoja_(ss_().getSheetByName(HOJA_LOG));
  const dst = archivoHistorial_();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'archivarHistorial').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('archivarHistorial').timeBased().onMonthDay(1).atHour(2).create();
  Logger.log('Listo. LOG_WEB protegida · archivo: ' + dst.getUrl() + ' · tarea mensual programada (día 1, 2 a. m.).');
}

/** VER COMO (solo administradores): sesión de SOLO LECTURA con los permisos de otra persona, para revisar qué ve y qué puede hacer. */
function verComo(token, fila) {
  const u = sesion_(token); if (u.rol !== 'ADMIN' || u.como) throw new Error('Solo un administrador puede usar "Ver como".');
  const h = ss_().getSheetByName(HOJA_ACCESO); fila = Number(fila);
  if (!(fila >= 2 && fila <= h.getLastRow()) || fila === u.fila) throw new Error('Fila inválida.');
  const f = h.getRange(fila, 1, 1, 8).getValues()[0];
  if (norm_(f[5]) !== 'SI') throw new Error('Ese usuario está deshabilitado: habilítalo con ⚙️ para verlo.');
  const t = Utilities.getUuid();
  CacheService.getScriptCache().put('s_' + t, JSON.stringify({ fila: fila, cod: String(f[4]).trim().toUpperCase(), como: u.nombre }), 3600);
  log_(u.nombre, 'VER_COMO', String(f[0]));
  return { token: t, user: sesion_(t) };
}
/** Cierra la sesión también en el servidor (el token deja de servir aunque alguien lo haya copiado). */
function salirSesion(token) { try { CacheService.getScriptCache().remove('s_' + token); } catch (e) {} return true; }
function yo(token) { const u = sesion_(token); u.area = (u.rol === 'ADMIN' || u.rol === 'ASISTENTE') ? miArea_(u) : ''; return u; }

/* ---------- registro de actividad ---------- */
function registro(token, desdeISO) {
  const u = sesion_(token); puede_(u, 'registro');
  const l = ss_().getSheetByName(HOJA_LOG), n = l.getLastRow(); if (n < 2) return [];
  const ini = Math.max(2, n - 1999), v = l.getRange(ini, 1, n - ini + 1, 5).getValues(), out = [];
  for (let i = v.length - 1; i >= 0; i--) {
    const r = v[i]; if (!(r[0] instanceof Date)) continue;
    out.push({ fecha: Utilities.formatDate(r[0], 'America/Lima', 'dd/MM/yyyy HH:mm'), iso: Utilities.formatDate(r[0], 'America/Lima', 'yyyy-MM-dd'), usuario: String(r[1]), accion: String(r[2]), detalle: String(r[3]), origen: String(r[4] || 'WEB') });
  }
  return out;
}
/** Disparador instalable: registra quién edita directamente el Excel/Sheet (hoja, celda, valor anterior → nuevo). */
function registrarCambio_(e) {
  tocarDatos_();   // edición directa en el Excel: la caché de expedientes se renueva
  try {
    const sh = e.range.getSheet(), nombre = sh.getName();
    if ([HOJA_LOG, HOJA_ACCESO].indexOf(nombre) >= 0) return;
    let quien = '';
    try { quien = e.user ? e.user.getEmail() : ''; } catch (x) {}
    if (!quien) try { quien = Session.getActiveUser().getEmail(); } catch (x) {}
    const h = ss_().getSheetByName(HOJA_ACCESO), v = h.getRange(2, 1, Math.max(h.getLastRow() - 1, 1), 2).getValues();
    const fila = v.filter(r => r[1] && String(r[1]).toLowerCase() === String(quien).toLowerCase())[0];
    const usuario = fila ? fila[0] + ' (' + quien + ')' : (quien || '(editor no identificado)');
    const celdas = e.range.getNumRows() * e.range.getNumColumns();
    let det = nombre + '!' + e.range.getA1Notation();
    if ([HOJA_BM, HOJA_CIERRE, HOJA_SBS].indexOf(nombre) >= 0 && e.range.getNumRows() === 1) {   // para el historial: código del expediente y nombre de la columna
      const r = e.range.getRow(), c = e.range.getColumn(), idx = String(sh.getRange(r, 1).getDisplayValue()).trim();
      const tit = String(sh.getRange(nombre === HOJA_SBS ? 5 : 4, c).getDisplayValue() || sh.getRange(4, c).getDisplayValue()).replace(/\s+/g, ' ').trim();
      if (/^EXP-/i.test(idx)) det = idx + ' · ' + det + (tit && celdas === 1 ? ' (' + tit + ')' : '');
    }
    if (celdas === 1) det += ': "' + String(e.oldValue == null ? '' : e.oldValue).slice(0, 60) + '" → "' + String(e.value == null ? '' : e.value).slice(0, 60) + '"';
    else det += ' (' + celdas + ' celdas)';
    log_(usuario, 'EDICION_EXCEL', det, 'EXCEL');
  } catch (err) {}
}

/* ---------- listas desplegables (se leen de las validaciones de BASE MAESTRA: siempre al día) ---------- */
function lista_(sh, a1, respaldo) {
  try {
    const dv = sh.getRange(a1).getDataValidation(); if (!dv) return respaldo;
    const t = dv.getCriteriaType(), v = dv.getCriteriaValues();
    let out = [];
    if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) out = v[0];
    else if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) out = v[0].getValues().map(r => r[0]);
    out = out.map(x => String(x).trim()).filter(x => x);
    return out.length ? out : respaldo;
  } catch (e) { return respaldo; }
}
/** Proyectos: lista "Espejo Proyectos" de BASE MAESTRA (AB2 hacia abajo, hasta la primera celda vacía). */
function proyectos_(bm) {
  const v = bm.getRange(2, 28, 80, 1).getDisplayValues().map(r => String(r[0]).trim()), i = v.indexOf('');
  return (i < 0 ? v : v.slice(0, i)).filter(x => x && !/^ESPEJO/i.test(x));
}
/** Cierres: 3 meses antes y 3 después del actual + los ya usados en la columna D + "NO APLICA". */
function cierres_(bm) {
  const M = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];
  const key = s => { const p = String(s).toUpperCase().split(/\s+/), i = M.indexOf(p[0]); return i < 0 || !(+p[1]) ? -1 : +p[1] * 12 + i; };
  const usados = bm.getRange(PRIMERA_FILA, 4, ULTIMA_FILA - PRIMERA_FILA + 1, 1).getDisplayValues().map(r => String(r[0]).trim()).filter(String);
  const n = new Date(), gen = [];
  for (let k = -3; k <= 3; k++) { const d = new Date(n.getFullYear(), n.getMonth() + k, 1); gen.push(M[d.getMonth()] + ' ' + d.getFullYear()); }
  const meses = [...new Set(gen.concat(usados.map(x => x.toUpperCase())))].filter(x => key(x) >= 0).sort((a, b) => key(b) - key(a));
  return meses.concat([usados.find(x => /^NO APLICA/i.test(x)) || 'NO APLICA (Estac./Depósito)']);
}
function vendedoresActivos_() {
  const c = ss_().getSheetByName('CATÁLOGOS'), v = c.getRange(91, 6, 109, 3).getValues();   // F:H filas 91-199
  return v.filter(r => r[0] && ['ACTIVO', 'APOYO'].indexOf(norm_(r[1])) >= 0).map(r => ({ nombre: String(r[0]).trim(), jefe: String(r[2]).trim(), apoyo: norm_(r[1]) === 'APOYO' }));
}
function opciones(token) {
  const u = sesion_(token); puede_(u, 'venta');
  const bm = ss_().getSheetByName(HOJA_BM);
  let vend = vendedoresActivos_();
  if (u.rol === 'VENDEDOR' || u.rol === 'APOYO') vend = vend.filter(x => norm_(x.nombre) === norm_(u.ref));
  else if (u.rol === 'JEFE') vend = vend.filter(x => !x.apoyo && (norm_(x.jefe) === norm_(u.ref) || norm_(x.nombre) === norm_(u.ref)));
  let tipos = lista_(bm, 'F5', ['DEPARTAMENTO', 'ESTACIONAMIENTO', 'DEPOSITO']);
  if (u.rol === 'APOYO') tipos = tipos.filter(t => !/^DEPARTAMENTO/i.test(t));
  return {
    proyecto: lista_(bm, 'E5', proyectos_(bm)), tipo: tipos, apoyo: vend.filter(x => x.apoyo).map(x => x.nombre),
    nacionalidad: lista_(bm, 'J9', ['PERUANA', 'VENEZOLANA', 'COLOMBIANA', 'CHILENA', 'ARGENTINA', 'ESPAÑOLA', 'OTRA']),
    estadoCivil: lista_(bm, 'K5', ['Soltero', 'Casado', 'Viudo', 'Divorciado']), sepBienes: lista_(bm, 'L5', ['Sí', 'No']),
    empresa: lista_(bm, 'O5', []), financiamiento: lista_(bm, 'T5', []), banco: lista_(bm, 'U5', []),
    cierre: lista_(bm, 'D5', cierres_(bm)), vendedor: vend.map(x => x.nombre)
  };
}

/* ---------- registrar venta ---------- */
function parseFecha_(s) { if (!s) return ''; const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})$/); if (!m) throw new Error('Fecha inválida: ' + s); return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])); }
function registrarVenta(token, d) {
  const u = sesion_(token); puede_(u, 'venta');
  const vend = opciones(token).vendedor;
  const req = { fechaIngreso: 'Fecha de ingreso', cierre: 'Venta corresponde al cierre de', proyecto: 'Proyecto', tipo: 'Tipo de inmueble', unidad: 'N° de unidad', cliente: 'Cliente', dni: 'DNI / CE / RUC', vendedor: 'Vendedor' };
  for (const k in req) if (!String(d[k] == null ? '' : d[k]).trim()) throw new Error('Falta: ' + req[k]);
  if (!vend.some(x => norm_(x) === norm_(d.vendedor))) throw new Error('Vendedor no permitido para su usuario.');
  const op = opciones(token);
  if (op.apoyo.some(x => norm_(x) === norm_(d.vendedor)) && /^DEPARTAMENTO/i.test(String(d.tipo))) throw new Error('El personal de apoyo solo registra estacionamientos o depósitos.');
  const dni = String(d.dni).replace(/'/g, '').trim();
  if ([8, 9, 11].indexOf(dni.length) < 0) throw new Error('DNI / CE / RUC debe tener 8, 9 u 11 caracteres.');
  const monto = d.monto === '' || d.monto == null ? '' : Number(d.monto);
  if (monto !== '' && !(monto >= 0)) throw new Error('Monto inválido.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ss = ss_(), bm = ss.getSheetByName(HOJA_BM);
    // siguiente fila = después de la última fila con cliente/vendedor/fecha (no rellena huecos para no pisar expedientes de CIERRE)
    const n = ULTIMA_FILA - PRIMERA_FILA + 1;
    const cols = bm.getRange(PRIMERA_FILA, 3, n, 21).getValues();   // C:W
    let ultima = PRIMERA_FILA - 1;
    cols.forEach((r, i) => { if (r[0] !== '' || r[5] !== '' || r[20] !== '') ultima = PRIMERA_FILA + i; });
    const fila = ultima + 1; if (fila > ULTIMA_FILA) throw new Error('BASE MAESTRA llegó a su límite (fila ' + ULTIMA_FILA + ').');
    // escribe SOLO C:W (A, B y X son fórmulas)
    const vals = CAMPOS.map(k => {
      if (k === 'fechaIngreso' || k === 'fechaSep' || k === 'fechaMcv') return parseFecha_(d[k]);
      if (k === 'monto') return monto;
      if (k === 'dni') return dni;
      return seguro_(String(d[k] == null ? '' : d[k]).trim());
    });
    // respeta la fórmula de NACIONALIDAD si el formulario no trae un valor
    const rango = bm.getRange(fila, COL_INI, 1, vals.length);
    rango.setValues([vals]);
    SpreadsheetApp.flush();
    const id = String(bm.getRange(fila, 1).getValue());
    // CIERRE: su ID es un valor (no fórmula): se registra en la misma fila
    const ci = ss.getSheetByName(HOJA_CIERRE);
    if (id && !ci.getRange(fila, 1).getValue()) ci.getRange(fila, 1).setValue(id);
    const lim = plazo_(d.cierre, 'E'), fuera = !!(lim && new Date() > lim);   // calendario: "Entrega de files" 5:30 pm
    log_(u.nombre, 'VENTA', id + ' fila ' + fila + ' · ' + d.cliente + (fuera ? ' · FUERA DE PLAZO (entrega ' + d.cierre + ' cerró ' + fmtPlazo_(lim) + ')' : ''));
    try { MailApp.sendEmail(admins_(), 'Nueva venta registrada: ' + id, u.nombre + ' registró ' + id + '\nCliente: ' + d.cliente + '\nProyecto: ' + d.proyecto + ' · Unidad ' + d.unidad + '\nVendedor: ' + d.vendedor); } catch (e) {}
    return { id: id, fila: fila, jefe: String(bm.getRange(fila, 24).getValue()), fueraPlazo: fuera, limite: lim ? fmtPlazo_(lim) : '' };
  } finally { lock.releaseLock(); }
}
function admins_() {
  const h = ss_().getSheetByName(HOJA_ACCESO), v = h.getRange(2, 1, Math.max(h.getLastRow() - 1, 1), 6).getValues();
  const c = v.filter(r => norm_(r[2]) === 'ADMIN' && norm_(r[5]) === 'SI' && r[1]).map(r => r[1]);
  if (!c.length) throw new Error('sin admin');
  return c.join(',');
}

/* ---------- expedientes y subsanación ---------- */
function visible_(u, vendedor, jefe) {
  if (u.rol === 'JEFE') return norm_(jefe) === norm_(u.ref);
  if (u.rol === 'VENDEDOR' || u.rol === 'APOYO') return norm_(vendedor) === norm_(u.ref);
  return true;   // ADMIN, ASISTENTE, GERENTE y roles nuevos: lo que pueden ver lo decide su permiso "ver"
}
/* ---------- caché de expedientes: leer CONTROL DE REVISIONES es lo más lento (fórmulas). Se guarda 2 min por usuario
 * y se renueva al instante con cualquier cambio (portal o edición directa del Excel). ---------- */
function verDatos_() { return CacheService.getScriptCache().get('verDatos') || '0'; }
function tocarDatos_() { try { CacheService.getScriptCache().put('verDatos', String(Date.now()), 21600); } catch (e) {} }
function misExpedientes(token) {
  const u = sesion_(token), c = CacheService.getScriptCache(), k = 'exp_' + verDatos_() + '_' + u.fila + '_' + u.permisos.join('');
  const z = c.get(k);
  if (z) try { return JSON.parse(Utilities.unzip(Utilities.newBlob(Utilities.base64Decode(z), 'application/x-gzip')).getDataAsString()); } catch (e) {}
  const out = misExpedientesSinCache_(token);
  try { const b = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(JSON.stringify(out))).getBytes()); if (b.length < 95000) c.put(k, b, 120); } catch (e) {}
  return out;
}
function misExpedientesSinCache_(token) {
  const u = sesion_(token); puede_(u, 'ver');
  const ss = ss_(), sh = ss.getSheetByName(HOJA_CONTROL);
  const v = sh.getRange(5, 1, 996, 17).getValues(), out = [];
  // Última respuesta por expediente en SUBSANACIONES: si es de Ventas, el expediente queda "SUBSANADO" (esperando revisión)
  // hasta que Contratos y Cobranzas lo apruebe en el Sheet o responda en el portal.
  const ult = {}, su = ss.getSheetByName(HOJA_SUB), ns = su.getLastRow();
  if (ns >= 2) su.getRange(2, 1, ns - 1, 6).getValues().forEach(r => { const id = String(r[1]).trim(); if (id) ult[id] = r; });
  const anu = typeof anulados_ === 'function' ? anulados_() : {};   // ventas anuladas (Anular.gs)
  v.forEach(r => {
    if (!r[0] || String(r[0]).charAt(0) === '#') return;
    if (!visible_(u, r[4], r[5])) return;
    const x = { id: String(r[0]), cliente: txt_(r[1]), proyecto: txt_(r[2]), unidad: txt_(r[3]), vendedor: txt_(r[4]), jefe: txt_(r[5]),
      sbs: txt_(r[6]), obsSbs: txt_(r[7]), ctr: txt_(r[8]), obsCtr: txt_(r[9]), revisor: txt_(r[10]), estado: txt_(r[11]), mes: txt_(r[16]) };
    // Si "Estado SBS" o "Estado CONTRACTUAL" dicen OBSERVADO, el expediente está POR CORREGIR aunque la columna L diga COMPLETO o PENDIENTE
    if (x.estado !== 'OBSERVADO' && (norm_(x.sbs) === 'OBSERVADO' || norm_(x.ctr) === 'OBSERVADO')) { x.estadoSheet = x.estado; x.estado = 'OBSERVADO'; }
    const s = ult[x.id];
    if (s && x.estado === 'OBSERVADO' && !/Contratos y Cobranzas/i.test(String(s[4]))) {
      x.estado = 'SUBSANADO';
      x.sub = { fecha: s[0] instanceof Date ? Utilities.formatDate(s[0], 'America/Lima', 'dd/MM/yyyy HH:mm') : txt_(s[0]), quien: String(s[4]).replace(/\s*\(Ventas\)\s*$/, ''), area: String(s[5] || '').trim() };
    }
    if (anu[x.id]) { x.anulado = true; x.anu = anu[x.id]; }
    out.push(x);
  });
  return out;
}
function detalle(token, id) {
  const u = sesion_(token), ss = ss_();
  const ex = misExpedientes(token).find(x => x.id === id);
  if (!ex) throw new Error('No tiene acceso a este expediente.');
  const items = [];
  // CIERRE: documentos I..X observados / pendientes
  const ci = ss.getSheetByName(HOJA_CIERRE), ids = ci.getRange(5, 1, 993, 1).getValues().map(r => String(r[0]));
  const i = ids.indexOf(id);
  if (i >= 0) {
    const hdr = ci.getRange(4, 9, 1, 16).getValues()[0], fila = ci.getRange(5 + i, 9, 1, 16).getValues()[0];
    fila.forEach((x, k) => { const e = norm_(x); if (e === 'OBSERVADO' || e === 'PENDIENTE') items.push({ area: 'CIERRE', item: String(hdr[k]).trim(), estado: e, col: 9 + k }); });
    const y = ci.getRange(5 + i, 25).getValue(); if (txt_(y)) ex.obsCtr = ex.obsCtr || txt_(y);
  }
  // SBS: ítems H..AI observados / pendientes (la fila de SBS = fila de BASE MAESTRA + 1)
  const sb = ss.getSheetByName(HOJA_SBS), sid = sb.getRange(6, 1, 995, 1).getValues().map(r => String(r[0])), j = sid.indexOf(id);
  if (j >= 0) {
    const h4 = sb.getRange(4, 8, 1, 28).getValues()[0], h5 = sb.getRange(5, 8, 1, 28).getValues()[0], fila = sb.getRange(6 + j, 8, 1, 28).getValues()[0];
    let grupo = '';
    fila.forEach((x, k) => {
      if (String(h4[k]).trim()) grupo = String(h4[k]).trim();
      const e = norm_(x);
      if (e === 'OBSERVADO' || e === 'PENDIENTE') items.push({ area: 'SBS', item: ((grupo && String(h5[k]).trim() ? grupo + ' · ' : grupo) + String(h5[k]).trim()).trim() || ('Columna ' + (k + 8)), estado: e, col: 8 + k });
    });
  }
  // historial de subsanaciones
  const su = ss.getSheetByName(HOJA_SUB), n = su.getLastRow(), hist = [];
  if (n >= 2) su.getRange(2, 1, n - 1, 8).getValues().forEach(r => { if (String(r[1]).trim() === id) hist.push({ fecha: txt_(r[0]), item: txt_(r[2]), texto: txt_(r[3]), quien: txt_(r[4]), estado: txt_(r[7]) }); });
  return { exp: ex, items: items, historial: hist.reverse() };
}
function subsanar(token, id, item, texto, adjunto) {
  const u = sesion_(token); puede_(u, 'subsanar');
  texto = String(texto || '').trim(); if (texto.length < 5) throw new Error('Escriba el detalle de la subsanación.');
  const det = detalle(token, id), cc = u.rol === 'ASISTENTE' || u.rol === 'ADMIN';
  const areas = areasDe_(det.exp, det.items, item), para = cc ? [] : contactos_().filter(c => c.area === 'TODO' || areas.indexOf(c.area) >= 0);
  const lim = cc ? null : plazo_(det.exp.mes, 'X'), fuera = !!(lim && new Date() > lim);   // calendario: "Comercial presenta subsanados" 5:30 pm
  const txtPlazo = lim ? (fuera ? '⏰ Fuera de plazo (límite ' + fmtPlazo_(lim) + ')' : '✅ En plazo (límite ' + fmtPlazo_(lim) + ')') : '';
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (adjunto && adjunto.base64) { const url = guardarAdjunto_(id, adjunto); texto += '\n' + url; }
    const sh = ss_().getSheetByName(HOJA_SUB), ahora = new Date();
    // Contratos responde → las subsanaciones pendientes de ese expediente quedan "Revisado" (y el expediente vuelve a 🔴 si sigue observado)
    if (cc && sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues().forEach((r, i) => {
      if (String(r[1]).trim() === id && /Pendiente/i.test(String(r[7]))) sh.getRange(i + 2, 8).setValue('✅ Revisado · ' + u.nombre + ' · ' + Utilities.formatDate(ahora, 'America/Lima', 'dd/MM/yyyy HH:mm')).setBackground('#DDF3E8');
    });
    sh.appendRow(fila_([ahora, id, String(item || 'General'), texto, u.nombre + (cc ? ' (Contratos y Cobranzas)' : ' (Ventas)'), areas.join(' + '), para.map(c => c.nombre).join(', '), cc ? '' : '🟠 Pendiente de revisión', txtPlazo]));
    if (!cc) { sh.getRange(sh.getLastRow(), 6, 1, 3).setBackground('#FDE7D6'); if (fuera) sh.getRange(sh.getLastRow(), 9).setBackground('#FBE0DC'); }
    log_(u.nombre, 'SUBSANACION', id + ' · ' + item + ' · ' + areas.join('+') + (fuera ? ' · FUERA DE PLAZO' : ''));
  } finally { lock.releaseLock(); }
  const correos = para.map(c => c.correo).filter(c => /@/.test(c));
  const ref = typeof expRef_ === 'function' ? expRef_(det.exp) : id;
  try { if (!cc) MailApp.sendEmail(correos.length ? correos.join(',') : admins_(), 'Subsanación recibida (' + areas.join(' + ') + '): ' + ref, u.nombre + ' subsanó ' + ref + ' (' + item + '):\n\n' + texto); } catch (e) {}
  return { areas: areas, contactos: para.filter(c => c.wa || c.correo).map(c => ({ nombre: c.nombre, wa: c.wa, correo: c.correo })), fueraPlazo: fuera, plazo: txtPlazo };
}

/**
 * Revisión de Contratos y Cobranzas: marca cada documento observado como OK u OBSERVADO en CIERRE / SBS
 * (las mismas celdas que se llenan a mano; CONTROL DE REVISIONES se recalcula solo).
 * Si algo sigue observado, la subsanación del vendedor queda "❌ Observado nuevamente" y el expediente vuelve a 🔴.
 */
function revisar(token, id, decisiones, comentario) {
  const u = sesion_(token); puede_(u, 'revisar');
  id = String(id || '').trim(); comentario = String(comentario || '').trim();
  const det = detalle(token, id), ss = ss_(), ahora = new Date(), fh = Utilities.formatDate(ahora, 'America/Lima', 'dd/MM/yyyy HH:mm');
  // solo documentos que hoy están observados/pendientes en este expediente
  const dec = (decisiones || []).map(d => ({ it: det.items.find(i => i.area === d.area && i.col === Number(d.col)), v: d.estado === 'OK' ? 'OK' : 'OBSERVADO' })).filter(d => d.it);
  const ok = dec.filter(d => d.v === 'OK').map(d => d.it.item), obs = dec.filter(d => d.v !== 'OK').map(d => d.it.item);
  const sigue = obs.length + det.items.filter(i => i.estado === 'OBSERVADO' && !dec.some(d => d.it === i)).length;
  if (sigue && !comentario) throw new Error('Escribe al vendedor qué falta corregir.');
  if (!dec.length && !comentario) throw new Error('Marca algún documento o escribe un comentario.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    dec.forEach(d => {
      const sh = ss.getSheetByName(d.it.area === 'CIERRE' ? HOJA_CIERRE : HOJA_SBS), f0 = d.it.area === 'CIERRE' ? 5 : 6;
      const i = sh.getRange(f0, 1, Math.max(sh.getLastRow() - f0 + 1, 1), 1).getDisplayValues().findIndex(r => String(r[0]).trim() === id);
      if (i >= 0) sh.getRange(f0 + i, d.it.col).setValue(d.v);
    });
    const su = ss.getSheetByName(HOJA_SUB), est = sigue ? '❌ Observado nuevamente · ' : '✅ Conforme · ';
    if (su.getLastRow() >= 2) su.getRange(2, 1, su.getLastRow() - 1, 8).getValues().forEach((r, k) => {
      if (String(r[1]).trim() === id && /Pendiente/i.test(String(r[7]))) su.getRange(k + 2, 8).setValue(est + u.nombre + ' · ' + fh).setBackground(sigue ? '#FBE0DC' : '#DDF3E8');
    });
    const txt = (ok.length ? '✅ Conforme: ' + ok.join(', ') + '. ' : '') + (obs.length ? '🔴 Sigue observado: ' + obs.join(', ') + '. ' : '') + comentario;
    su.appendRow([ahora, id, 'Revisión', seguro_(txt || 'Revisado'), u.nombre + ' (Contratos y Cobranzas)', [...new Set(det.items.map(i => i.area))].join(' + '), '', sigue ? '🔴 Devuelto al vendedor' : '✅ Conforme']);
    log_(u.nombre, 'REVISION', id + ' · ' + ok.length + ' OK · ' + obs.length + ' observados');
  } finally { lock.releaseLock(); }
  return { conformes: ok, observados: obs };
}

/** Foto o documento de la subsanación: se guarda PRIVADO en Drive (carpeta del expediente) y el enlace queda en el historial. */
function guardarAdjunto_(id, adj) {
  const mime = String(adj.mime || 'image/jpeg'); if (!/^image\/(jpeg|png|webp)$/.test(mime)) throw new Error('Solo se aceptan fotos.');
  if (String(adj.base64).length > 4000000) throw new Error('La foto es muy pesada.');
  const raiz = DriveApp.getFoldersByName('ADJUNTOS SUBSANACIONES - MG'), r = raiz.hasNext() ? raiz.next() : DriveApp.createFolder('ADJUNTOS SUBSANACIONES - MG');
  const sub = r.getFoldersByName(id), f = sub.hasNext() ? sub.next() : r.createFolder(id);
  const nombre = id + ' ' + Utilities.formatDate(new Date(), 'America/Lima', 'yyyy-MM-dd HHmmss') + '.jpg';
  return f.createFile(Utilities.newBlob(Utilities.base64Decode(adj.base64), mime, nombre)).getUrl();
}

/* ---------- dashboard de ventas (sale de BASE MAESTRA: D cierre · E proyecto · F tipo · R monto · W vendedor · X jefe) ---------- */
function ventas(token) {
  const u = sesion_(token); puede_(u, 'dashboard');
  const v = ss_().getSheetByName(HOJA_BM).getRange(PRIMERA_FILA, 1, ULTIMA_FILA - PRIMERA_FILA + 1, 24).getValues(), out = [];
  const anu = typeof anulados_ === 'function' ? anulados_() : {};
  v.forEach(r => {
    const vend = txt_(r[22]); if (!vend || anu[String(r[0]).trim()]) return;
    const m = typeof r[17] === 'number' ? r[17] : Number(String(r[17]).replace(/[^\d.-]/g, ''));
    const mes = r[3] instanceof Date ? ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'][r[3].getMonth()] + ' ' + r[3].getFullYear() : txt_(r[3]).toUpperCase();
    out.push({ mes: mes, proyecto: txt_(r[4]), tipo: txt_(r[5]), monto: isFinite(m) ? m : 0, vendedor: vend, jefe: txt_(r[23]) });
  });
  return out;
}

/* ---------- reporte interno (solo usuarios con el permiso "interno") ---------- */
function hojaInterno_() {
  const ss = ss_();
  if (HOJA_INTERNO) { const h = ss.getSheetByName(HOJA_INTERNO); if (!h) throw new Error('No existe la hoja ' + HOJA_INTERNO + '.'); return h; }
  const shs = ss.getSheets();
  for (let i = 0; i < shs.length; i++) {
    const h = shs[i], n = Math.min(h.getLastRow(), 15), c = h.getLastColumn(); if (n < 1 || c < 1) continue;
    const v = h.getRange(1, 1, n, c).getDisplayValues();
    for (let r = 0; r < v.length; r++) if (v[r].some(x => norm_(x) === 'INFO INTERNA')) return h;
  }
  throw new Error('No encontré la hoja del Reporte Interno (columna "Info Interna").');
}
function reporteInterno(token) {
  const u = sesion_(token); puede_(u, 'interno');
  const h = hojaInterno_(), v = h.getDataRange().getDisplayValues();
  const hi = v.findIndex(r => r.some(x => norm_(x) === 'INFO INTERNA')), hd = v[hi].map(norm_);
  const col = (...ks) => { for (const k of ks) { const i = hd.indexOf(k); if (i >= 0) return i; } return -1; };
  const c = { exp: col('EXPEDIENTE', 'ID', 'N° EXPEDIENTE'), pro: col('PROYECTO'), uni: col('N° UNIDAD', 'UNIDAD', 'N° DE UNIDAD'), cli: col('CLIENTE'), mes: col('MES DE CIERRE', 'MES CIERRE', 'CIERRE'), inf: col('INFO INTERNA', 'INTERNO') };
  if (c.exp < 0) c.exp = 0;
  const g = (r, i) => i >= 0 ? String(r[i]).trim() : '', out = [];
  for (let i = hi + 1; i < v.length; i++) {
    const r = v[i], id = g(r, c.exp); if (!/^EXP-\d/i.test(id)) continue;
    out.push({ expediente: id, proyecto: g(r, c.pro), unidad: g(r, c.uni), cliente: g(r, c.cli), mes: g(r, c.mes), interno: g(r, c.inf) });
  }
  log_(u.nombre, 'REPORTE_INTERNO', out.length + ' filas · hoja ' + h.getName());
  return { hoja: h.getName(), filas: out };
}

/* ---------- escala de comisiones (ref. 01 set 2025): cada rol ve SOLO la suya; vive aquí porque la web es pública ---------- */
const COMISIONES = {
  VENDEDOR: { titulo: 'Asesores de venta (ref. 01 set 2025)', tablas: [
    { t: 'Regular · Plan de ahorros', c: ['Tramo', 'Total', '50% Minuta', '50% Escritura'], f: [['1 depa', 1000, 500, 500], ['2 a 3 depas', 1500, 750, 750], ['4 a 5 depas', 1800, 900, 900], ['6 a 8 depas', 2000, 1000, 1000], ['9 depas a más', 3000, 1500, 1500]] },
    { t: 'Regular · Crédito hipotecario o directo MG', c: ['Tramo', 'Total', '70% Minuta', '30% Escritura'], f: [['1 depa', 1000, 700, 300], ['2 a 3 depas', 1500, 1050, 450], ['4 a 5 depas', 1800, 1260, 540], ['6 a 8 depas', 2000, 1400, 600], ['9 depas a más', 3000, 2100, 900]] },
    { t: 'Regular · Contado 100%', c: ['Tramo', '100% Minuta'], f: [['1 depa', 1000], ['2 a 3 depas', 1500], ['4 a 5 depas', 1800], ['6 a 8 depas', 2000], ['9 depas a más', 3000]] }],
    notas: ['El tramo se determina por el total de ventas del mes del asesor: todas sus ventas del mes reciben el valor de ese tramo, según la modalidad.', 'No se paga comisión si hubo un descuento adicional al precio vigente.'] },
  JEFE: { titulo: 'Jefes de ventas', tablas: [
    { t: 'Jefes de ventas', c: ['Unidades (equipo)', 'Total', '70% 1er pago', '30% 2do pago'], f: [['1 a 6', 400, 280, 120], ['7 a 13', 600, 420, 180], ['14 a 18', 800, 560, 240], ['19 a más', 1000, 700, 300]] }],
    notas: ['El tramo se determina por el total de unidades vendidas por todo tu equipo en el mes. Monto de minuta = unidades × valor 70% del tramo (igual para todas las unidades, sin distinción de modalidad).', 'No se paga comisión si hubo un descuento adicional al precio vigente.'] },
  GERENTE: { titulo: 'Gerente comercial', tablas: [
    { t: 'Gerente comercial', c: ['Concepto', 'Monto'], f: [['Por unidad (empresa)', 200]] }],
    notas: ['S/ 200 por unidad vendida a nivel empresa.'] }
};
function comisiones(token) {
  const u = sesion_(token), k = { ADMIN: ['VENDEDOR', 'JEFE', 'GERENTE'], VENDEDOR: ['VENDEDOR'], JEFE: ['JEFE'], GERENTE: ['GERENTE'] }[u.rol] || [];
  return k.map(x => COMISIONES[x]);
}

/* ---------- checklist de presentación (PDF que el jefe de ventas firma a mano) ---------- */
const HOJA_CHECK = 'CHECKLIST', CARPETA_CHECK = 'CHECKLIST VENTAS - MG';
const CHECK_TIPOS = { NATURAL: 'Persona natural · solo un titular', CASADO: 'Persona natural · casado / conviviente / copropietario', JURIDICA: 'Persona jurídica' };
const CHECK_FIN = { DIRECTO: 'Crédito directo / Hipotecario', AHORRO: 'Ahorro Casa', CONTADO: 'Contado' };   // hipotecario = crédito directo (mismos documentos)
// Editable: cada grupo aparece solo si el tipo de cliente o el financiamiento está en "si" (vacío = siempre).
const CHECK_GRUPOS = [
  { t: 'Titular', si: ['NATURAL', 'CASADO'], items: ['DNI / CE / PASAPORTE TITULAR', 'INTERPOL TITULAR', 'MIGRACIONES TITULAR', 'OFAC TITULAR', 'GOOGLE TITULAR', 'SUNAT TITULAR', 'ESSALUD TITULAR', 'DJ TITULAR', 'ROP', 'PARTIDAS Y OTROS SEGÚN SEA EL CASO', 'PERMISO EN CASO DE PASAPORTE PARA FIRMA DE CONTRATO'] },
  { t: 'Esposo(a) · conviviente · copropietario', si: ['CASADO'], items: ['DNI / CE / PASAPORTE', 'INTERPOL', 'MIGRACIONES', 'OFAC', 'GOOGLE', 'SUNAT', 'ESSALUD', 'DJ'].map(x => x + ' DE ESPOSO-CONVIVIENTE O COPROPIETARIO') },
  { t: 'Empresa', si: ['JURIDICA'], items: ['FICHA RUC', 'VIGENCIA PODER (30 DÍAS)', 'DNI REPRESENTANTE LEGAL', 'DECLARACIÓN JURADA DE EMPRESA', 'ROP', 'PARTIDAS Y OTROS SEGÚN SEA EL CASO', 'PERMISO EN CASO DE PASAPORTE PARA FIRMA DE CONTRATO'] },
  { t: 'Accionistas con más del 25% de acciones', si: ['JURIDICA'], items: ['DNI', 'SUNAT', 'ESSALUD', 'INTERPOL', 'OFAC', 'GOOGLE', 'DJ'].map(x => x + ' DE CADA ACCIONISTA') },
  { t: 'Financiamiento · Ahorro Casa', si: ['AHORRO'], items: ['CONSTANCIA DE PRIMER ABONO DE AHORRO', 'VERIFICACIÓN DE BIENES', 'CONSTANCIA DE BONIFICACIÓN'] },
  { t: 'Financiamiento · Crédito directo / Hipotecario', si: ['DIRECTO', 'HIPOTECARIO'], items: ['CARTA DE APROBACIÓN DEL BANCO', 'RECIBOS DE PAGO'] },
  { t: 'Documentos de la venta', si: [], items: ['CONSTANCIA DE SEPARACIÓN', 'PROFORMA', 'LIQUIDACIÓN DE VENTA ACTUALIZADA', 'MINUTA DE COMPRA VENTA', 'DJ DE EJECUTIVO DE VENTAS', 'CLÁUSULA ADICIONAL O ACLARATORIAS', 'CONFORMIDAD POR PRECIO ESPECIAL', 'CONFORMIDAD POR FECHA DE ENTREGA ANTICIPADA', 'CONSTANCIA DE BONIFICACIÓN POR REFERIDO'] }
];
function checkExp_(u, id, perm) {
  puede_(u, perm || 'checklist');
  id = String(id || '').trim();
  const bm = ss_().getSheetByName(HOJA_BM), ids = bm.getRange(PRIMERA_FILA, 1, ULTIMA_FILA - PRIMERA_FILA + 1, 1).getDisplayValues();
  const i = ids.findIndex(r => String(r[0]).trim() === id); if (!id || i < 0) throw new Error('No encontré el expediente ' + id + ' en BASE MAESTRA.');
  const r = bm.getRange(PRIMERA_FILA + i, 1, 1, 24).getValues()[0];
  if (!visible_(u, txt_(r[22]), txt_(r[23]))) throw new Error('No tiene acceso a este expediente.');
  return { id: id, cliente: txt_(r[7]), dni: txt_(r[8]).replace(/'/g, ''), dniConyuge: txt_(r[12]).replace(/'/g, ''), conyuge: txt_(r[13]), proyecto: txt_(r[4]), tipo: txt_(r[5]), unidad: txt_(r[6]), estadoCivil: txt_(r[10]), financiamiento: txt_(r[19]), vendedor: txt_(r[22]), jefe: txt_(r[23]), cierre: txt_(r[3]) };
}
function checklistDatos(token, id) { const x = checkExp_(sesion_(token), id); x.tipos = CHECK_TIPOS; x.fin = CHECK_FIN; x.grupos = CHECK_GRUPOS; return x; }
function checklistPdf(token, id, d) {
  const u = sesion_(token), x = checkExp_(u, id); d = d || {};
  if (!CHECK_TIPOS[d.tipo]) throw new Error('Elija el tipo de cliente.');
  if (!CHECK_FIN[d.fin]) throw new Error('Elija el financiamiento.');
  const gr = CHECK_GRUPOS.filter(g => !g.si.length || g.si.indexOf(d.tipo) >= 0 || g.si.indexOf(d.fin) >= 0);
  const marc = {}; (d.marcados || []).forEach(k => marc[k] = true);
  const elegidos = []; gr.forEach(g => g.items.forEach(it => { if (marc[g.t + '|' + it]) elegidos.push(it); }));
  if (!elegidos.length) throw new Error('Marque al menos un documento.');
  const ahora = new Date(), fh = Utilities.formatDate(ahora, 'America/Lima', 'dd/MM/yyyy HH:mm:ss');
  const lim = plazo_(x.cierre, 'E'), fuera = !!(lim && ahora > lim);
  const plazoTxt = lim ? 'Cierre ' + x.cierre + ' · entrega de files hasta ' + fmtPlazo_(lim) + ' · ' + (fuera ? '<b style="color:#B3261E">FUERA DE PLAZO</b>' : '<b style="color:#1E8E5A">DENTRO DEL PLAZO</b>') : (x.cierre ? 'Cierre ' + x.cierre : '');
  const H = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const bx = on => '<span style="display:inline-block;width:10px;height:10px;border:1.2px solid #0C1A3D;text-align:center;font-size:9px;line-height:10px;font-weight:bold;margin-right:5px">' + (on ? 'X' : '&nbsp;') + '</span>';
  const campo = (k, v) => '<td style="padding:3px 6px;color:#555;width:23%">' + k + '</td><td style="padding:3px 6px;border-bottom:1px solid #999;font-weight:bold">' + H(v) + '</td>';
  const opc = (o, sel) => Object.keys(o).map(k => '<span style="margin-right:14px">' + bx(k === sel) + H(o[k]) + '</span>').join('');
  let cuerpo = '';
  gr.forEach(g => {
    const it = g.items, mitad = Math.ceil(it.length / 2), cel = j => j < it.length ? bx(marc[g.t + '|' + it[j]]) + H(it[j]) : '';
    let filas = ''; for (let j = 0; j < mitad; j++) filas += '<tr><td style="padding:2px 4px;width:50%">' + cel(j) + '</td><td style="padding:2px 4px">' + cel(j + mitad) + '</td></tr>';
    cuerpo += '<div style="background:#0C1A3D;color:#fff;font-weight:bold;padding:3px 6px;margin-top:8px">' + H(g.t.toUpperCase()) + '</div><table style="width:100%;border-collapse:collapse">' + filas + '</table>';
  });
  const html = '<html><head><meta charset="utf-8"></head><body style="font-family:Arial,sans-serif;font-size:9.5px;color:#111;margin:18px">' +
    '<table style="width:100%"><tr><td style="font-size:17px;font-weight:bold;color:#0C1A3D">CHECK LIST DE EXPEDIENTE DE VENTA</td><td style="text-align:right;font-size:11px"><b>' + H(x.id) + '</b></td></tr></table>' +
    '<table style="width:100%;border-collapse:collapse;margin-top:6px">' +
    '<tr>' + campo('Nombre y apellidos / Razón social', x.cliente) + campo('DNI / RUC', x.dni) + '</tr>' +
    '<tr>' + campo('Cónyuge o copropietario', x.conyuge) + campo('DNI', x.dniConyuge) + '</tr>' +
    '<tr>' + campo('Proyecto', x.proyecto) + campo('Ejecutivo de ventas', x.vendedor) + '</tr>' +
    '<tr>' + campo('Departamento', d.depa) + campo('Estacionamiento', d.estac) + '</tr>' +
    '<tr>' + campo('Depósito', d.deposito) + campo('Otros (detallar)', d.otros) + '</tr></table>' +
    '<div style="margin-top:8px"><b>Tipo de cliente:</b> ' + opc(CHECK_TIPOS, d.tipo) + '</div>' +
    '<div style="margin-top:4px"><b>Financiamiento:</b> ' + opc(CHECK_FIN, d.fin) + '</div>' + cuerpo +
    '<div style="margin-top:10px;padding:6px;border:1px solid #0C1A3D"><b>Presentado por:</b> ' + H(u.nombre) + ' &nbsp;·&nbsp; <b>Fecha y hora:</b> ' + fh + ' &nbsp;·&nbsp; <b>Documentos marcados:</b> ' + elegidos.length + (plazoTxt ? '<br>' + plazoTxt : '') + '</div>' +
    '<table style="width:100%;margin-top:42px"><tr><td style="width:50%;text-align:center"><div style="border-top:1px solid #111;width:80%;margin:0 auto;padding-top:3px">Firma ejecutivo de ventas<br><b>' + H(x.vendedor) + '</b></div></td>' +
    '<td style="text-align:center"><div style="border-top:1px solid #111;width:80%;margin:0 auto;padding-top:3px">V°B° y firma jefe de ventas<br><b>' + H(x.jefe) + '</b></div></td></tr></table>' +
    '<div style="margin-top:14px;text-align:center;color:#777;font-size:8px">Cuadro de Revisiones · Diseñado por el área de Contratos y Cobranzas</div></body></html>';
  const nombre = 'CHECKLIST ' + x.id + ' ' + Utilities.formatDate(ahora, 'America/Lima', 'yyyy-MM-dd HHmm') + '.pdf';
  const pdf = Utilities.newBlob(html, 'text/html', 'c.html').getAs('application/pdf').setName(nombre);
  const it = DriveApp.getFoldersByName(CARPETA_CHECK), raiz = it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA_CHECK);
  const url = raiz.createFile(pdf).getUrl();
  const ss = ss_(); let h = ss.getSheetByName(HOJA_CHECK);
  if (!h) { h = ss.insertSheet(HOJA_CHECK); h.appendRow(['Fecha y hora', 'ID expediente', 'Cliente', 'Proyecto', 'Unidad', 'Tipo de cliente', 'Financiamiento', 'Documentos marcados', 'N° marcados', 'Presentado por', 'Jefe de ventas', 'PDF']); h.setFrozenRows(1); h.getRange('A1:L1').setFontWeight('bold').setBackground('#0C1A3D').setFontColor('#FFFFFF'); }
  h.appendRow(fila_([ahora, x.id, x.cliente, x.proyecto, x.unidad, CHECK_TIPOS[d.tipo], CHECK_FIN[d.fin], elegidos.join(' · '), elegidos.length, u.nombre, x.jefe, url]));
  log_(u.nombre, 'CHECKLIST', x.id + ' · ' + elegidos.length + ' documentos');
  return { nombre: nombre, base64: Utilities.base64Encode(pdf.getBytes()) };
}

/* ---------- backup ---------- */
function exportarXlsx_() {
  const url = 'https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/export?format=xlsx';
  const r = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('No se pudo exportar (' + r.getResponseCode() + ').');
  return r.getBlob().setName('REVISIÓN DE VENTAS - MG ' + Utilities.formatDate(new Date(), 'America/Lima', 'yyyy-MM-dd HHmm') + '.xlsx');
}
function descargarBackup(token) {
  const u = sesion_(token); puede_(u, 'backup');
  const b = exportarXlsx_(); log_(u.nombre, 'BACKUP_DESCARGA', b.getName());
  return { nombre: b.getName(), base64: Utilities.base64Encode(b.getBytes()) };
}
/** Guarda una copia en Drive y conserva solo las últimas 12. La programa instalarBackupAutomatico(). */
function backupEnDrive() {
  const it = DriveApp.getFoldersByName(CARPETA_BACKUP), carpeta = it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA_BACKUP);
  carpeta.createFile(exportarXlsx_());
  const archivos = []; const fi = carpeta.getFiles(); while (fi.hasNext()) archivos.push(fi.next());
  archivos.sort((a, b) => b.getDateCreated() - a.getDateCreated()).slice(12).forEach(f => f.setTrashed(true));
  log_('SISTEMA', 'BACKUP_DRIVE', 'Copia automática: ' + carpeta.getName(), 'AUTO');
}
/** Ejecutar UNA vez. dias = 7 (semanal) o 15 (quincenal). */
function instalarBackupAutomatico() {
  instalarTareas(); return;
  const dias = 7;
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'backupEnDrive').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('backupEnDrive').timeBased().everyDays(dias).atHour(23).create();
}
function backupDriveAhora(token) {
  const u = sesion_(token); puede_(u, 'backup'); backupEnDrive(); log_(u.nombre, 'BACKUP_DRIVE', 'Manual'); return true;
}


/* =====================================================================
 * Ficha
 * ===================================================================== */
/**
 * FICHA DE REVISIÓN — Contratos y Cobranzas revisa desde el portal, igual que en el Excel:
 *   CIERRE: documentos I..X (OK / OBSERVADO / NO APLICA) · Y Observación contractual · Z Estado revisión contractual
 *           · AB Revisado por (<PERSONA_1> / <PERSONA_2>) · AC N° Revisión
 *   SBS:    ítems F..AI (Check list, Carta banco, DNI titular/cónyuge, …) · AJ Estado revisión SBS · AK Observación SBS · AM N° Revisión
 * Escribe en las MISMAS celdas; CONTROL DE REVISIONES se recalcula solo con sus fórmulas.
 * Las columnas se ubican por su encabezado (filas 4 y 5); si alguien las mueve, se siguen encontrando.
 */
const FICHA_OPC = ['OK', 'OBSERVADO', 'NO APLICA'];
const FICHA_HOJAS = {
  CIERRE: { hoja: HOJA_CIERRE, f0: 5, c0: 9, n: 16 },    // I..X
  SBS: { hoja: HOJA_SBS, f0: 6, c0: 6, n: 30 }           // F..AI
};
const FICHA_EXTRAS = { obs: /^OBSERVACI/, estado: /^ESTADO REVISI/, revisor: /^REVISADO POR/, nrev: /^N\S*\s*REVISI/ };

function opcionesDv_(dv, respaldo) {
  try {
    if (!dv) return respaldo;
    const t = dv.getCriteriaType(), v = dv.getCriteriaValues(); let l = [];
    if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) l = v[0];
    else if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) l = v[0].getValues().map(r => r[0]);
    l = l.map(x => String(x).trim()).filter(x => x && x !== 'XXX');
    return l.length ? l : respaldo;
  } catch (e) { return respaldo; }
}

/** Lee la fila del expediente en CIERRE o SBS con todo lo que se revisa. null si no tiene fila. */
function leerFicha_(area, id) {
  const cfg = FICHA_HOJAS[area], sh = ss_().getSheetByName(cfg.hoja);
  if (!sh || sh.getLastRow() < cfg.f0) return null;
  const ids = sh.getRange(cfg.f0, 1, sh.getLastRow() - cfg.f0 + 1, 1).getDisplayValues();
  const i = ids.findIndex(r => String(r[0]).trim() === id); if (i < 0) return null;
  const fila = cfg.f0 + i, lc = sh.getLastColumn(), L = s => String(s || '').replace(/\s+/g, ' ').trim();
  const h4 = sh.getRange(4, 1, 1, lc).getDisplayValues()[0].map(L), h5 = area === 'SBS' ? sh.getRange(5, 1, 1, lc).getDisplayValues()[0].map(L) : h4.map(() => '');   // en CIERRE la fila 5 ya es data
  // grupos de la fila 4 (celdas combinadas: "DNI/CE/PAS/RUC" cubre TITU. y CÓNY.)
  const g4 = h4.slice();
  sh.getRange(4, 1, 1, lc).getMergedRanges().forEach(m => { const v = L(m.getDisplayValue()); for (let c = m.getColumn(); c <= m.getLastColumn(); c++) g4[c - 1] = v; });
  const filaRg = sh.getRange(fila, 1, 1, lc), vals = filaRg.getDisplayValues()[0], dvs = filaRg.getDataValidations()[0], fxs = filaRg.getFormulas()[0];
  const nombre = t => t.replace(/^TITU\.?$/i, 'Titular').replace(/^C[OÓ]NY\.?$/i, 'Cónyuge');
  const campos = [];
  for (let col = cfg.c0; col < cfg.c0 + cfg.n; col++) {
    const k = col - 1, sub = area === 'SBS' ? h5[k] : '', grupo = area === 'SBS' ? g4[k] : '';
    const titulo = nombre(sub || h4[k] || h5[k] || ('Columna ' + col)), valor = String(vals[k]).trim(), op = opcionesDv_(dvs[k], FICHA_OPC.slice());
    if (valor && op.indexOf(valor) < 0) op.push(valor);
    campos.push({ col: col, titulo: titulo, grupo: grupo && grupo !== titulo ? grupo : '', valor: valor, opciones: op, fx: !!fxs[k] });
  }
  const out = { area: area, fila: fila, campos: campos };
  Object.keys(FICHA_EXTRAS).forEach(key => {
    for (let col = cfg.c0 + cfg.n; col <= lc; col++) {
      const t = h5[col - 1] || h4[col - 1];
      if (FICHA_EXTRAS[key].test(t.toUpperCase())) {
        const k = col - 1, resp = key === 'estado' ? ['APROBADO', 'OBSERVADO'] : null;
        out[key] = { col: col, titulo: t, valor: String(vals[k]).trim(), fx: !!fxs[k], opciones: resp || key === 'revisor' ? opcionesDv_(dvs[k], resp || []) : null };
        break;
      }
    }
  });
  return out;
}

/** Ficha completa (CIERRE + SBS) de un expediente, para quien revisa. */
function ficha(token, id) {
  const u = sesion_(token); puede_(u, 'revisar');
  id = String(id || '').trim();
  if (!misExpedientes(token).some(x => x.id === id)) throw new Error('No tiene acceso a este expediente.');
  const r = {};
  Object.keys(FICHA_HOJAS).forEach(a => {
    if (!puedeArea_(u, a, 'ver')) { r[a] = { oculto: true }; return; }
    const f = leerFicha_(a, id); if (f) f.soloLectura = !puedeArea_(u, a, 'edit'); r[a] = f;
  });
  // Seguimiento.gs: quién revisó primero cada área y N° de revisión calculado del historial (SBS no tiene columna "Revisado por")
  if (typeof seguimiento_ === 'function') try {
    const ex = misExpedientes(token).find(x => x.id === id) || {};
    const sg = seguimiento_(id, ex.estado, { CIERRE: ex.ctr, SBS: ex.sbs, revisorCierre: ex.revisor });
    Object.keys(FICHA_HOJAS).forEach(a => {
      const f = r[a], s = sg.areas[a]; if (!f || f.oculto || !s) return;
      if (!f.revisor) f.revisor = { col: 0, titulo: 'Revisado por', valor: s.revisor, fx: true, auto: true };
      else if (!f.revisor.valor && s.revisor) f.revisor.valor = s.revisor;
      if (f.nrev && !(Number(f.nrev.valor) > 0) && s.rondas) f.nrev.valor = String(s.rondas);
      if (!f.nrev && s.rondas) f.nrev = { col: 0, titulo: 'N° Revisión', valor: String(s.rondas), fx: true, auto: true };
      f.seg = s;
    });
    r.seguimiento = sg;
  } catch (e) {}
  return r;
}

/**
 * Guarda la ficha.
 *  cambios = [{area, col, valor}]  (documentos)
 *  extras  = {CIERRE: {obs, estado}, SBS: {obs, estado}}  (solo lo que cambió)
 *  comentario = mensaje al vendedor (queda en SUBSANACIONES)
 * "Revisado por" y "N° Revisión" son automáticos: quien guarda, y 1 la primera vez / +1 cuando Comercial
 * subsanó esa área desde la última revisión. Cada persona solo edita las áreas habilitadas en ⚙️.
 */
function guardarFicha(token, id, cambios, extras, comentario) {
  const u = sesion_(token); puede_(u, 'revisar');
  id = String(id || '').trim(); comentario = String(comentario || '').trim(); extras = extras || {};
  if (!misExpedientes(token).some(x => x.id === id)) throw new Error('No tiene acceso a este expediente.');
  const ss = ss_(), ahora = new Date(), fh = Utilities.formatDate(ahora, 'America/Lima', 'dd/MM/yyyy HH:mm');
  const resumen = [], observados = [], tocadas = [], obsArea = {};
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    // subsanaciones de Comercial pendientes de revisión en este expediente (col F = área: SBS / CIERRE / SBS + CIERRE)
    const su = ss.getSheetByName(HOJA_SUB), subs = su.getLastRow() >= 2 ? su.getRange(2, 1, su.getLastRow() - 1, 8).getValues() : [];
    const pend = subs.map((r, k) => ({ k: k, area: String(r[5] || ''), ok: String(r[1]).trim() === id && /Pendiente/i.test(String(r[7])) })).filter(p => p.ok);
    const deArea = (p, a) => !p.area.trim() || p.area.toUpperCase().indexOf(a) >= 0;
    Object.keys(FICHA_HOJAS).forEach(area => {
      const f = leerFicha_(area, id); if (!f) return;
      const sh = ss.getSheetByName(FICHA_HOJAS[area].hoja), ex = extras[area] || {}, mios = (cambios || []).filter(c => c.area === area), antes = resumen.length;
      if ((mios.length || ['obs', 'estado'].some(k => ex[k] != null)) && !puedeArea_(u, area, 'edit')) throw new Error('Su usuario no tiene habilitado editar ' + area + '. Consulte con el administrador.');
      mios.forEach(c => {
        const campo = f.campos.find(x => x.col === Number(c.col)), v = String(c.valor == null ? '' : c.valor).trim();
        if (!campo || campo.fx || v === campo.valor) return;
        if (v && campo.opciones.indexOf(v) < 0) throw new Error('Valor no permitido en ' + campo.titulo + ': ' + v);
        const ant = campo.valor; sh.getRange(f.fila, campo.col).setValue(v); campo.valor = v;
        resumen.push(area + ' · ' + (campo.grupo ? campo.grupo + ' ' : '') + campo.titulo + ': ' + (ant || '—') + ' → ' + (v || '(vacío)'));
      });
      ['obs', 'estado'].forEach(key => {   // revisor y N° revisión ya no vienen del portal: se calculan abajo
        const c = f[key]; if (!c || c.fx || ex[key] == null) return;
        const v = String(ex[key]).trim(); if (v === c.valor) return;
        if (c.opciones && c.opciones.length && v && c.opciones.indexOf(v) < 0) throw new Error('Valor no permitido en ' + c.titulo + ': ' + v);
        const ant = String(c.valor || '').slice(0, 80); sh.getRange(f.fila, c.col).setValue(key === 'obs' ? seguro_(v) : v); c.valor = v;
        resumen.push(area + ' · ' + c.titulo + ': ' + (ant || '—') + ' → ' + (v.slice(0, 120) || '(vacío)'));
      });
      if (resumen.length > antes) {
        tocadas.push(area);
        const rv = f.revisor;   // Revisado por = quien guarda (el nombre de la lista que coincide con su primer nombre)
        if (rv && !rv.fx) {
          const k = sinTilde_(u.nombre).split(' ')[0], o = rv.opciones && rv.opciones.length ? rv.opciones.find(x => sinTilde_(x).split(' ')[0] === k) : u.nombre;
          if (o && o !== rv.valor) { sh.getRange(f.fila, rv.col).setValue(o); resumen.push(area + ' · ' + rv.titulo + ': ' + (rv.valor || '—') + ' → ' + o); }
        }
        const nr = f.nrev;   // N° Revisión
        if (nr && !nr.fx) {
          // con Seguimiento.gs se calcula del historial (sirve aunque la celda esté vacía en expedientes antiguos)
          const cur = Number(nr.valor) || 0, nuevo = typeof segRondaNueva_ === 'function' ? segRondaNueva_(id, area, cur) : (cur === 0 ? 1 : (pend.some(p => deArea(p, area)) ? cur + 1 : cur));
          if (nuevo !== cur) { sh.getRange(f.fila, nr.col).setValue(nuevo); resumen.push(area + ' · ' + nr.titulo + ': ' + (cur || '—') + ' → ' + nuevo); }
        }
      }
      const ob = f.campos.filter(x => norm_(x.valor) === 'OBSERVADO').map(x => area + ' · ' + (x.grupo ? x.grupo + ' ' : '') + x.titulo);
      if (f.estado && norm_(f.estado.valor) === 'OBSERVADO' && !ob.length) ob.push(area + ' (estado OBSERVADO)');
      obsArea[area] = ob.length > 0; ob.forEach(x => observados.push(x));
    });
    if (!resumen.length && !comentario) throw new Error('No hay cambios para guardar.');
    const sigue = observados.length > 0;
    // cierra solo las subsanaciones de las áreas que se revisaron ahora (cada equipo cierra lo suyo)
    pend.filter(p => !tocadas.length || tocadas.some(a => deArea(p, a))).forEach(p => {
      const s = tocadas.length ? tocadas.some(a => deArea(p, a) && obsArea[a]) : sigue;
      su.getRange(p.k + 2, 8).setValue((s ? '❌ Observado nuevamente · ' : '✅ Conforme · ') + u.nombre + ' · ' + fh).setBackground(s ? '#FBE0DC' : '#DDF3E8');
    });
    const txt = (sigue ? '🔴 Observado: ' + observados.join(', ') + '. ' : '✅ Sin observaciones. ') + comentario;
    su.appendRow([ahora, id, 'Revisión', seguro_(txt), u.nombre + ' (Contratos y Cobranzas)', tocadas.length ? tocadas.join(' + ') : 'SBS + CIERRE', '', sigue ? '🔴 Devuelto al vendedor' : '✅ Conforme']);
    log_(u.nombre, 'REVISION_FICHA', id + ' · ' + resumen.length + ' cambios · ' + observados.length + ' observados · ' + resumen.join(' | ').slice(0, 200));
    logMuchos_(u.nombre, 'FICHA_CAMBIO', resumen.map(r => id + ' · ' + r));   // una fila por cambio (historial del expediente)
  } finally { lock.releaseLock(); }
  return { cambios: resumen, observados: observados };
}

/** Detalle para el portal: igual que detalle(), pero "Qué falta" incluye TODAS las columnas (también Check list y Carta banco de SBS). */
function detalleCompleto_(token, id) {
  const d = detalle(token, id);
  try {
    const it = [];
    ['CIERRE', 'SBS'].forEach(a => {
      const f = leerFicha_(a, String(id).trim()); if (!f) return;
      f.campos.forEach(c => { const e = norm_(c.valor); if (e === 'OBSERVADO' || e === 'PENDIENTE') it.push({ area: a, item: (c.grupo ? c.grupo + ' · ' : '') + c.titulo, estado: e, col: c.col }); });
      if (a === 'CIERRE' && f.revisor) d.exp.revisor = d.exp.revisor || f.revisor.valor;
    });
    d.items = it;
  } catch (e) {}
  if (typeof seguimiento_ === 'function') try { d.seguimiento = seguimiento_(String(id).trim(), d.exp.estado, { CIERRE: d.exp.ctr, SBS: d.exp.sbs, revisorCierre: d.exp.revisor }); } catch (e) {}
  return d;
}


/* =====================================================================
 * Buzon
 * ===================================================================== */
/**
 * BUZÓN DEL PORTAL — cualquier usuario reporta errores, incidencias o sugerencias (texto + hasta 3 capturas).
 * Todo queda en la pestaña "BUZÓN" del Sheet (se crea sola) y las capturas en Drive (carpeta privada "BUZÓN PORTAL - MG").
 * Cada persona ve sus envíos y la respuesta; quien tiene el permiso "accesos" (administración) ve todo y responde.
 */
const HOJA_BUZON = 'BUZÓN', CARPETA_BUZON = 'BUZÓN PORTAL - MG';
const BUZON_TIPOS = ['ERROR', 'INCIDENCIA', 'SUGERENCIA'];
const BUZON_ESTADOS = ['NUEVO', 'EN REVISIÓN', 'RESUELTO', 'DESCARTADO'];

function hojaBuzon_() {
  const ss = ss_(); let h = ss.getSheetByName(HOJA_BUZON);
  if (!h) {
    h = ss.insertSheet(HOJA_BUZON);
    h.appendRow(['Fecha', 'Usuario', 'Rol', 'Tipo', 'Mensaje', 'Capturas (Drive)', 'Pantalla', 'Estado', 'Respuesta', 'Atendido por', 'Fecha respuesta']);
    h.setFrozenRows(1); h.getRange('A1:K1').setFontWeight('bold').setBackground('#0C1A3D').setFontColor('#FFFFFF');
    h.setColumnWidth(5, 420); h.setColumnWidth(6, 260); h.setColumnWidth(9, 320);
    h.getRange(2, 8, 998, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(BUZON_ESTADOS, true).build());
  }
  return h;
}

/** Envía un reporte. capturas = [{mime, base64}] (máx. 3). */
function buzonEnviar(token, tipo, texto, capturas, pantalla) {
  const u = sesion_(token);
  tipo = norm_(tipo); if (BUZON_TIPOS.indexOf(tipo) < 0) throw new Error('Elige si es error, incidencia o sugerencia.');
  texto = String(texto || '').trim(); if (texto.length < 5) throw new Error('Cuéntanos un poco más (mínimo 5 letras).');
  texto = texto.slice(0, 3000);
  const caps = (capturas || []).slice(0, 3), urls = [];
  if (caps.length) {
    const it = DriveApp.getFoldersByName(CARPETA_BUZON), carpeta = it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA_BUZON);
    caps.forEach((c, i) => {
      const mime = String(c && c.mime || ''); if (!/^image\/(jpeg|png|webp)$/.test(mime)) throw new Error('Solo se aceptan imágenes (capturas o fotos).');
      if (String(c.base64 || '').length > 4000000) throw new Error('Una captura es muy pesada.');
      const nombre = 'BUZON ' + Utilities.formatDate(new Date(), 'America/Lima', 'yyyy-MM-dd HHmmss') + ' ' + u.nombre + ' ' + (i + 1) + '.jpg';
      urls.push(carpeta.createFile(Utilities.newBlob(Utilities.base64Decode(c.base64), mime, nombre)).getUrl());
    });
  }
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { hojaBuzon_().appendRow(fila_([new Date(), u.nombre, u.rol, tipo, texto, urls.join('\n'), String(pantalla || '').slice(0, 200), 'NUEVO', '', '', ''])); }
  finally { lock.releaseLock(); }
  log_(u.nombre, 'BUZON', tipo + ' · ' + texto.slice(0, 120) + (urls.length ? ' · ' + urls.length + ' captura(s)' : ''));
  try { MailApp.sendEmail(admins_(), '📮 Buzón del portal · ' + tipo + ' de ' + u.nombre, u.nombre + ' (' + u.rol + ') escribió:\n\n' + texto + (urls.length ? '\n\nCapturas:\n' + urls.join('\n') : '') + '\n\nRespóndelo en el portal (Ayuda → Buzón) o en la pestaña BUZÓN del Sheet.'); } catch (e) {}
  return true;
}

/** Lista: cada usuario ve lo suyo; administración (permiso "accesos") ve todo. */
function buzonLista(token) {
  const u = sesion_(token), admin = u.permisos.indexOf('accesos') >= 0, h = hojaBuzon_(), n = h.getLastRow();
  if (n < 2) return { admin: admin, estados: BUZON_ESTADOS, lista: [] };
  const f = d => d instanceof Date ? Utilities.formatDate(d, 'America/Lima', 'dd/MM/yyyy HH:mm') : txt_(d);
  const lista = h.getRange(2, 1, n - 1, 11).getValues().map((r, i) => ({
    fila: i + 2, fecha: f(r[0]), usuario: String(r[1]), rol: String(r[2]), tipo: String(r[3]), mensaje: String(r[4]),
    capturas: String(r[5] || '').split('\n').filter(x => /^https:\/\//.test(x)), pantalla: String(r[6]), estado: String(r[7] || 'NUEVO'),
    respuesta: String(r[8]), atendio: String(r[9]), fresp: f(r[10])
  })).filter(x => admin || norm_(x.usuario) === norm_(u.nombre)).reverse().slice(0, 300);
  return { admin: admin, estados: BUZON_ESTADOS, lista: lista };
}

/** Administración cambia el estado y responde. */
function buzonResponder(token, fila, estado, respuesta) {
  const u = sesion_(token); puede_(u, 'accesos');
  const h = hojaBuzon_(); fila = Number(fila);
  if (!(fila >= 2 && fila <= h.getLastRow())) throw new Error('Fila inválida.');
  estado = norm_(estado); if (BUZON_ESTADOS.indexOf(estado) < 0) throw new Error('Estado inválido.');
  respuesta = String(respuesta || '').trim().slice(0, 2000);
  h.getRange(fila, 8, 1, 4).setValues([fila_([estado, respuesta, u.nombre, new Date()])]);
  log_(u.nombre, 'BUZON_RESPUESTA', 'Fila ' + fila + ' · ' + estado);
  return true;
}


/* =====================================================================
 * Anular
 * ===================================================================== */
/**
 * DESHACER Y ANULAR VENTAS — sin borrar filas (SBS y CIERRE están amarradas a la fila de BASE MAESTRA).
 *  · Deshacer registro: solo la ÚLTIMA venta registrada, sin revisión en CIERRE/SBS; quien la registró tiene 30 min,
 *    administración (permiso "accesos") en cualquier momento. Limpia C:W de BASE MAESTRA y el ID de CIERRE:
 *    la siguiente venta vuelve a usar esa misma fila (no quedan huecos).
 *  · Anular venta: NO borra nada. Registra el ID en la pestaña ANULADOS (motivo, quién, cuándo) y el portal la oculta
 *    de expedientes, dashboard y duplicados. Se puede reactivar.
 */
const HOJA_ANU = 'ANULADOS', DESHACER_MIN = 30;

function hojaAnu_() {
  const ss = ss_(); let h = ss.getSheetByName(HOJA_ANU);
  if (!h) {
    h = ss.insertSheet(HOJA_ANU);
    h.appendRow(['Fecha', 'ID expediente', 'Cliente', 'Proyecto', 'Unidad', 'Motivo', 'Anulado por', 'Estado']);
    h.setFrozenRows(1); h.getRange('A1:H1').setFontWeight('bold').setBackground('#0C1A3D').setFontColor('#FFFFFF'); h.setColumnWidth(6, 360);
  }
  return h;
}
/** { ID: {motivo, quien, fecha} } de las ventas anuladas vigentes (el último registro de cada ID manda). */
function anulados_() {
  const h = ss_().getSheetByName(HOJA_ANU); if (!h || h.getLastRow() < 2) return {};
  const m = {};
  h.getRange(2, 1, h.getLastRow() - 1, 8).getValues().forEach(r => {
    const id = String(r[1]).trim(); if (!id) return;
    if (norm_(r[7]) === 'ANULADO') m[id] = { motivo: String(r[5]), quien: String(r[6]), fecha: txt_(r[0]) }; else delete m[id];
  });
  return m;
}
function filaBM_(id) {
  const bm = ss_().getSheetByName(HOJA_BM), ids = bm.getRange(PRIMERA_FILA, 1, ULTIMA_FILA - PRIMERA_FILA + 1, 1).getDisplayValues();
  const i = ids.findIndex(r => String(r[0]).trim() === id); if (i < 0) throw new Error('No encontré ' + id + ' en BASE MAESTRA.');
  return { bm: bm, fila: PRIMERA_FILA + i };
}

function anularVenta(token, id, motivo) {
  const u = sesion_(token); if (u.permisos.indexOf('accesos') < 0) puede_(u, 'revisar');
  id = String(id || '').trim(); motivo = String(motivo || '').trim();
  if (motivo.length < 5) throw new Error('Escribe el motivo de la anulación.');
  const ex = misExpedientes(token).find(x => x.id === id); if (!ex) throw new Error('No tiene acceso a este expediente.');
  if (ex.anulado) throw new Error('Este expediente ya está anulado.');
  hojaAnu_().appendRow(fila_([new Date(), id, ex.cliente, ex.proyecto, ex.unidad, motivo, u.nombre, 'ANULADO']));
  log_(u.nombre, 'ANULAR', id + ' · ' + ex.cliente + ' · ' + motivo);
  return true;
}
function reactivarVenta(token, id) {
  const u = sesion_(token); if (u.permisos.indexOf('accesos') < 0) puede_(u, 'revisar');
  id = String(id || '').trim();
  if (!anulados_()[id]) throw new Error('Este expediente no está anulado.');
  hojaAnu_().appendRow(fila_([new Date(), id, '', '', '', 'Reactivado', u.nombre, 'REACTIVADO']));
  log_(u.nombre, 'REACTIVAR', id);
  return true;
}

function deshacerVenta(token, id) {
  const u = sesion_(token); puede_(u, 'venta');
  id = String(id || '').trim(); if (!id) throw new Error('Falta el expediente.');
  const admin = u.permisos.indexOf('accesos') >= 0, ss = ss_();
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const { bm, fila } = filaBM_(id);
    // 1) debe ser la última venta (si no, quedaría un hueco: para eso está Anular)
    const n = ULTIMA_FILA - fila;
    if (n > 0 && bm.getRange(fila + 1, 3, n, 21).getValues().some(r => r[0] !== '' || r[5] !== '' || r[20] !== ''))
      throw new Error('Solo se puede deshacer la última venta registrada. Para esta usa "Anular venta".');
    // 2) sin revisión en CIERRE ni SBS
    ['CIERRE', 'SBS'].forEach(a => { const f = leerFicha_(a, id); if (f && f.campos.some(c => c.valor && !c.fx)) throw new Error('Ya tiene revisión en ' + a + '. Usa "Anular venta".'); });
    // 3) quien la registró, dentro de 30 min (administración: siempre)
    if (!admin) {
      const l = ss.getSheetByName(HOJA_LOG), nl = l.getLastRow(), ini = Math.max(2, nl - 999);
      const v = nl >= 2 ? l.getRange(ini, 1, nl - ini + 1, 4).getValues() : [];
      const reg = v.reverse().find(r => String(r[2]) === 'VENTA' && String(r[3]).indexOf(id + ' fila ') === 0);
      if (!reg || norm_(reg[1]) !== norm_(u.nombre)) throw new Error('Solo quien registró la venta (o administración) puede deshacerla.');
      if (!(reg[0] instanceof Date) || (new Date() - reg[0]) > DESHACER_MIN * 60000) throw new Error('Pasaron más de ' + DESHACER_MIN + ' minutos. Pide a Contratos y Cobranzas que la anule.');
    }
    const cliente = bm.getRange(fila, 8).getDisplayValue();
    bm.getRange(fila, COL_INI, 1, CAMPOS.length).clearContent();   // C:W (A, B y X son fórmulas: no se tocan)
    const ci = ss.getSheetByName(HOJA_CIERRE); if (String(ci.getRange(fila, 1).getDisplayValue()).trim() === id) ci.getRange(fila, 1).clearContent();
    log_(u.nombre, 'VENTA_DESHECHA', id + ' fila ' + fila + ' · ' + cliente);
  } finally { lock.releaseLock(); }
  return true;
}

/* ---------- CARGO DE ENTREGA: comercial deja el expediente físico en Contratos y Cobranzas ----------
 * Hoja CARGOS: solo se agregan filas (al confirmar la recepción se completan sus columnas de recepción).
 * Cualquiera con el permiso 'cargo' genera el cargo; el equipo de Contratos (roles que no son de comercial) confirma la recepción.
 * Cada PDF lleva quién lo generó y la fecha y hora. */
const HOJA_CARGO = 'CARGOS', CARPETA_CARGO = 'CARGOS DE ENTREGA - MG';
// Versión resumida del checklist (para una entrega rápida). Editable.
const CARGO_DOCS = ['Checklist firmado por el jefe de ventas', 'DNI / CE / pasaporte (titular y cónyuge)', 'Consultas SBS (Interpol, OFAC, Migraciones, SUNAT, EsSalud)', 'DJ de origen de fondos', 'Documentos de empresa (ficha RUC, vigencia de poder)', 'Constancia de separación', 'Proforma y liquidación de venta', 'Minuta de compraventa', 'DJ del ejecutivo de ventas', 'Documentos del financiamiento (banco / Ahorro Casa)', 'Cláusulas, conformidades y otros'];
const CARGO_COLS = ['N° cargo', 'Fecha entrega', 'ID expediente', 'Cliente', 'Proyecto', 'Unidad', 'Documentos', 'N° documentos', 'Folios', 'Observación', 'Entregado por', 'Estado', 'Fecha recepción', 'Recibido por', 'Obs. recepción', 'PDF'];
function hojaCargo_() {
  const ss = ss_(); let h = ss.getSheetByName(HOJA_CARGO);
  if (!h) { h = ss.insertSheet(HOJA_CARGO); h.appendRow(CARGO_COLS); h.setFrozenRows(1); h.getRange(1, 1, 1, CARGO_COLS.length).setFontWeight('bold').setBackground('#0C1A3D').setFontColor('#FFFFFF'); }
  return h;
}
const esContratos_ = u => ROLES_CATALOGO.indexOf(u.rol) < 0;
function cargos_() {
  const h = hojaCargo_(), n = h.getLastRow(); if (n < 2) return [];
  const f = v => v instanceof Date ? Utilities.formatDate(v, 'America/Lima', 'dd/MM/yyyy HH:mm') : String(v || '');
  return h.getRange(2, 1, n - 1, CARGO_COLS.length).getValues().map((r, i) => ({ fila: i + 2, numero: String(r[0]).trim(), fecha: f(r[1]), id: String(r[2]), cliente: String(r[3]), proyecto: String(r[4]), unidad: String(r[5]), docs: String(r[6]).split(' · ').filter(String), n: Number(r[7]) || 0, folios: String(r[8] || ''), obs: String(r[9] || ''), entrega: String(r[10]), estado: String(r[11] || 'PENDIENTE'), fechaRec: f(r[12]), recibe: String(r[13] || ''), obsRec: String(r[14] || '') })).filter(c => c.numero);
}
function cargoPorNumero_(numero) { const c = cargos_().filter(k => k.numero === String(numero || '').trim())[0]; if (!c) throw new Error('No encontré el cargo ' + numero + '.'); return c; }
/** Formulario del cargo de un expediente + sus entregas anteriores. */
function cargoDatos(token, id) {
  const u = sesion_(token), x = checkExp_(u, id, 'cargo');
  x.docs = CARGO_DOCS; x.cargos = cargos_().filter(c => c.id === x.id).reverse(); x.puedeRecibir = esContratos_(u);
  return x;
}
/** Bandeja: Contratos ve todas las entregas; comercial solo las de sus expedientes. Lo más reciente primero. */
function cargosLista(token) {
  const u = sesion_(token); puede_(u, 'cargo');
  const ids = {}; if (!esContratos_(u)) misExpedientes(token).forEach(x => ids[x.id] = 1);
  return { puedeRecibir: esContratos_(u), lista: cargos_().filter(c => esContratos_(u) || ids[c.id]).reverse().slice(0, 200) };
}
/** Genera el cargo (N° CE-AAAA-0001…) con lo que se entrega. */
function cargoEntregar(token, id, d) {
  const u = sesion_(token), x = checkExp_(u, id, 'cargo'); d = d || {};
  const docs = (d.docs || []).filter(t => CARGO_DOCS.indexOf(t) >= 0);
  if (!docs.length) throw new Error('Marque al menos un documento.');
  const folios = String(d.folios || '').replace(/\D/g, '').slice(0, 4), obs = String(d.obs || '').trim().slice(0, 300);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const h = hojaCargo_(), ahora = new Date(), pref = 'CE-' + Utilities.formatDate(ahora, 'America/Lima', 'yyyy') + '-';
    const ult = cargos_().filter(c => c.numero.indexOf(pref) === 0).reduce((m, c) => Math.max(m, Number(c.numero.slice(pref.length)) || 0), 0);
    const numero = pref + ('000' + (ult + 1)).slice(-4);
    h.appendRow(fila_([numero, ahora, x.id, x.cliente, x.proyecto, x.unidad, docs.join(' · '), docs.length, folios, obs, u.nombre, 'PENDIENTE', '', '', '', '']));
    SpreadsheetApp.flush();
    log_(u.nombre, 'CARGO_ENTREGA', x.id + ' · ' + numero + ' · ' + docs.length + ' documentos' + (folios ? ' · ' + folios + ' folios' : '') + (obs ? ' · ' + obs : ''));
    return cargoPdf_(u, cargoPorNumero_(numero), true);
  } finally { lock.releaseLock(); }
}
/** Contratos confirma la recepción (conforme) u observa la entrega (falta algo / no coincide). */
function cargoRecibir(token, numero, conforme, obs) {
  const u = sesion_(token); puede_(u, 'cargo');
  if (!esContratos_(u)) throw new Error('La recepción la confirma el equipo de Contratos y Cobranzas.');
  const c = cargoPorNumero_(numero);
  if (c.estado === 'RECIBIDO') throw new Error('Este cargo ya fue recibido por ' + c.recibe + ' (' + c.fechaRec + ').');
  obs = String(obs || '').trim().slice(0, 300);
  if (!conforme && obs.length < 3) throw new Error('Escriba qué falta o qué no coincide.');
  const est = conforme ? 'RECIBIDO' : 'OBSERVADO', ahora = new Date();
  hojaCargo_().getRange(c.fila, 12, 1, 4).setValues([fila_([est, ahora, u.nombre, obs])]);
  log_(u.nombre, conforme ? 'CARGO_RECIBIDO' : 'CARGO_OBSERVADO', c.id + ' · ' + c.numero + (obs ? ' · ' + obs : ''));
  c.estado = est; c.fechaRec = Utilities.formatDate(ahora, 'America/Lima', 'dd/MM/yyyy HH:mm'); c.recibe = u.nombre; c.obsRec = obs;
  return cargoPdf_(u, c, true);
}
/** Vuelve a generar el PDF de un cargo (para imprimir o reenviar), con quien lo genera ahora. */
function cargoImprimir(token, numero) {
  const u = sesion_(token), c = cargoPorNumero_(numero);
  checkExp_(u, c.id, 'cargo');   // valida permiso y que vea ese expediente
  log_(u.nombre, 'CARGO_IMPRESO', c.id + ' · ' + c.numero);
  return cargoPdf_(u, c, false);
}
function cargoPdf_(u, c, guardar) {
  const H = s => String(s == null ? '' : s).replace(/[&<>"]/g, k => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[k]));
  const gen = Utilities.formatDate(new Date(), 'America/Lima', 'dd/MM/yyyy HH:mm:ss');
  const E = { PENDIENTE: ['#8A5A00', '#FFF4DC', 'PENDIENTE DE RECEPCIÓN'], RECIBIDO: ['#1E7A4C', '#E3F5EC', 'RECIBIDO CONFORME'], OBSERVADO: ['#A32D2D', '#FCEBEB', 'RECIBIDO CON OBSERVACIÓN'] }[c.estado] || ['#555555', '#EEEEEE', c.estado];
  const fila = (k, v) => '<tr><td style="padding:5px 0;color:#667;width:32%">' + k + '</td><td style="padding:5px 0;font-weight:bold">' + H(v) + '</td></tr>';
  const bx = '<span style="display:inline-block;width:10px;height:10px;border:1.2px solid #0C1A3D;text-align:center;font-size:9px;line-height:10px;font-weight:bold;margin-right:6px">X</span>';
  const html = '<html><head><meta charset="utf-8"></head><body style="font-family:Arial,sans-serif;font-size:11px;color:#1d2433;margin:28px">' +
    '<table style="width:100%"><tr><td><div style="font-size:19px;font-weight:bold;color:#0C1A3D">CARGO DE ENTREGA DE EXPEDIENTE</div><div style="color:#667;margin-top:2px">Grupo MG · Contratos y Cobranzas</div></td>' +
    '<td style="text-align:right"><div style="color:#667;font-size:10px">N° de cargo</div><div style="font-size:16px;font-weight:bold;color:#0C1A3D">' + H(c.numero) + '</div></td></tr></table>' +
    '<div style="margin:12px 0;padding:7px 10px;background:' + E[1] + ';color:' + E[0] + ';font-weight:bold;text-align:center">' + E[2] + '</div>' +
    '<table style="width:100%;border-collapse:collapse">' + fila('Expediente', c.id) + fila('Cliente', c.cliente) + fila('Proyecto / unidad', c.proyecto + ' · ' + c.unidad) +
    fila('Entregado por', c.entrega + ' · ' + c.fecha) + fila('Recibido por', c.recibe ? c.recibe + ' · ' + c.fechaRec : '—') + (c.folios ? fila('Folios', c.folios) : '') + '</table>' +
    '<div style="margin-top:12px;font-weight:bold;color:#0C1A3D">Documentos entregados (' + c.docs.length + ')</div>' +
    '<div style="margin-top:4px;line-height:1.9">' + c.docs.map(t => bx + H(t)).join('<br>') + '</div>' +
    (c.obs ? '<div style="margin-top:10px"><b>Observación de entrega:</b> ' + H(c.obs) + '</div>' : '') +
    (c.obsRec ? '<div style="margin-top:6px"><b>Observación de recepción:</b> ' + H(c.obsRec) + '</div>' : '') +
    '<table style="width:100%;margin-top:48px"><tr><td style="width:50%;text-align:center"><div style="border-top:1px solid #1d2433;width:80%;margin:0 auto;padding-top:3px">Entregó<br><b>' + H(c.entrega) + '</b></div></td>' +
    '<td style="text-align:center"><div style="border-top:1px solid #1d2433;width:80%;margin:0 auto;padding-top:3px">Recibió · Contratos y Cobranzas<br><b>' + H(c.recibe) + '</b></div></td></tr></table>' +
    '<div style="margin-top:22px;padding-top:6px;border-top:1px solid #dddde8;color:#889;font-size:9px">Generado por ' + H(u.nombre) + ' el ' + gen + ' · Cuadro de Revisiones</div></body></html>';
  const nombre = 'CARGO ' + c.numero + ' ' + c.id + '.pdf';
  const pdf = Utilities.newBlob(html.replace(/<table style="/g, '<table style="font-size:11px;'), 'text/html', 'c.html').getAs('application/pdf').setName(nombre);
  if (guardar) { try { const it = DriveApp.getFoldersByName(CARPETA_CARGO), raiz = it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA_CARGO); hojaCargo_().getRange(c.fila, 16).setValue(raiz.createFile(pdf).getUrl()); } catch (e) {} }
  const texto = '📥 *Cargo ' + c.numero + '* · ' + E[2].toLowerCase() + '\nExpediente: ' + c.id + ' · ' + c.cliente + '\n' + c.proyecto + ' · ' + c.unidad +
    '\nEntregó: ' + c.entrega + ' (' + c.fecha + ')' + (c.recibe ? '\nRecibió: ' + c.recibe + ' (' + c.fechaRec + ')' : '') +
    '\nDocumentos: ' + c.docs.length + (c.folios ? ' · ' + c.folios + ' folios' : '') + (c.obsRec ? '\nObservación: ' + c.obsRec : '') + '\nGenerado por ' + u.nombre + ' · ' + gen;
  return { numero: c.numero, estado: c.estado, nombre: nombre, base64: Utilities.base64Encode(pdf.getBytes()), texto: texto };
}
