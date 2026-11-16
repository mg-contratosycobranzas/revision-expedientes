/**
 * 📍 ¿DÓNDE ESTÁ EL FILE? — custodia del expediente físico después de validado (archivo independiente).
 * Cada movimiento queda en la hoja FILE_UBICACION (una fila por movimiento, no se borra nada):
 *   RETIRO     → alguien de Comercial se lleva el file (para subsanar, firmar, etc.).
 *                Si lo registra Contratos queda CONFIRMADO; si lo registra Comercial queda POR CONFIRMAR hasta que Contratos lo confirme.
 *   DEVOLUCION → el file vuelve a Contratos. Si lo registra Contratos queda RECIBIDO; si lo registra Comercial, POR RECIBIR.
 *   Un movimiento que Contratos rechaza ("no se llevó el file" / "no llegó") queda ANULADO y no cuenta.
 * La ubicación actual de cada EXP = su último movimiento válido (sin movimientos = en Contratos).
 * Se conecta solo con Code.gs: agrega sus funciones a API_PUBLICA.
 */
const HOJA_FUB = 'FILE_UBICACION';
const FUB_HDR = ['Fecha', 'EXP', 'Cliente', 'Proyecto', 'Unidad', 'Movimiento', 'Tiene el file', 'Registró', 'Motivo / nota', 'Estado', 'Confirmó', 'Fecha confirmación'];
const FUB_COMERCIAL = ['VENDEDOR', 'JEFE', 'APOYO', 'GERENTE'];
try {
  ['fileUbic', 'filesFuera', 'fileRetirar', 'fileRetirarVarios', 'fileDevolver', 'fileConfirmar'].forEach(f => { if (API_PUBLICA.indexOf(f) < 0) API_PUBLICA.push(f); });
  ['fileUbic', 'filesFuera'].forEach(f => { if (SOLO_LECTURA.indexOf(f) < 0) SOLO_LECTURA.push(f); });
} catch (e) {}

/** Contratos y Cobranzas: puede revisar y no es de Comercial. */
function fubCC_(u) { return u.permisos.indexOf('revisar') >= 0 && FUB_COMERCIAL.indexOf(u.rol) < 0; }
function fubHoja_() {
  const ss = ss_(); let h = ss.getSheetByName(HOJA_FUB);
  if (!h) {
    h = ss.insertSheet(HOJA_FUB); h.appendRow(FUB_HDR); h.setFrozenRows(1);
    h.getRange(1, 1, 1, FUB_HDR.length).setFontWeight('bold').setBackground('#1F2A44').setFontColor('#FFFFFF');
  }
  return h;
}
function fubFilas_() {
  const h = ss_().getSheetByName(HOJA_FUB); if (!h || h.getLastRow() < 2) return [];
  return h.getRange(2, 1, h.getLastRow() - 1, FUB_HDR.length).getValues().map((r, i) => ({ fila: i + 2, r: r }));
}
const fubFecha_ = d => d instanceof Date ? Utilities.formatDate(d, 'America/Lima', 'dd/MM/yyyy HH:mm') : txt_(d);
const fubMov_ = x => ({ fecha: fubFecha_(x.r[0]), movimiento: txt_(x.r[5]), persona: txt_(x.r[6]), registro: txt_(x.r[7]), nota: txt_(x.r[8]), estado: txt_(x.r[9]), confirmo: txt_(x.r[10]), fConfirmo: fubFecha_(x.r[11]) });
/** Ubicación actual por EXP (solo los que tienen algún movimiento válido). */
function fubEstados_() {
  const o = {};
  fubFilas_().forEach(x => {
    const id = String(x.r[1]).trim(), est = norm_(x.r[9]); if (!id || est === 'ANULADO') return;
    const ret = norm_(x.r[5]) === 'RETIRO', ms = x.r[0] instanceof Date ? x.r[0].getTime() : 0;
    o[id] = { id: id, cliente: txt_(x.r[2]), proyecto: txt_(x.r[3]), unidad: txt_(x.r[4]),
      estado: ret ? 'FUERA' : (est === 'RECIBIDO' ? 'EN CONTRATOS' : 'DEVOLVIENDO'),
      persona: txt_(x.r[6]), registro: txt_(x.r[7]), nota: txt_(x.r[8]), porConfirmar: est === 'POR CONFIRMAR' || est === 'POR RECIBIR',
      desde: fubFecha_(x.r[0]), dias: ms ? Math.floor((Date.now() - ms) / 864e5) : null, fila: x.fila };
  });
  return o;
}
function fubAcceso_(token, id) {
  const ex = misExpedientes(token).find(x => x.id === String(id || '').trim());
  if (!ex) throw new Error('No tienes acceso al expediente ' + id + '.');
  return ex;
}
const fubYo_ = (u, nombre) => [norm_(u.nombre), norm_(u.ref)].filter(Boolean).indexOf(norm_(nombre)) >= 0;

/** Ubicación e historial de un expediente. */
function fileUbic(token, id) {
  const u = sesion_(token); fubAcceso_(token, id); id = String(id).trim();
  const e = fubEstados_()[id] || { id: id, estado: 'EN CONTRATOS' };
  e.movimientos = fubFilas_().filter(x => String(x.r[1]).trim() === id).map(fubMov_).reverse();
  e.cc = fubCC_(u); e.mio = !!e.persona && fubYo_(u, e.persona);
  return e;
}
/** Files que no están en Contratos. Contratos ve todos; Comercial, los suyos y los de sus expedientes. */
function filesFuera(token) {
  const u = sesion_(token), cc = fubCC_(u), est = fubEstados_();
  const mios = cc ? null : misExpedientes(token).reduce((a, x) => { a[x.id] = 1; return a; }, {});
  return Object.keys(est).map(k => est[k]).filter(e => e.estado !== 'EN CONTRATOS' && (cc || mios[e.id] || fubYo_(u, e.persona)))
    .map(e => Object.assign({ mio: fubYo_(u, e.persona) }, e)).sort((a, b) => (b.dias || 0) - (a.dias || 0));
}
function fubRetiro_(u, ex, persona, motivo, cc, ahora) {
  return [ahora, ex.id, seguro_(ex.cliente), seguro_(ex.proyecto), seguro_(ex.unidad), 'RETIRO', seguro_(persona), u.nombre, seguro_(motivo),
    cc ? 'CONFIRMADO' : 'POR CONFIRMAR', cc ? u.nombre : '', cc ? ahora : ''];
}
/** Registrar que alguien se lleva el file. Comercial solo puede registrarse a sí mismo; Contratos elige la persona. */
function fileRetirar(token, id, persona, motivo) {
  const u = sesion_(token), ex = fubAcceso_(token, id), cc = fubCC_(u);
  persona = cc ? String(persona || '').trim() : u.nombre; motivo = String(motivo || '').trim() || 'Subsanar observación';
  if (persona.length < 3) throw new Error('Indica quién se lleva el file.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const e = fubEstados_()[ex.id];
    if (e && e.estado !== 'EN CONTRATOS') {
      if (norm_(e.persona) === norm_(persona)) throw new Error('El file ya figura en poder de ' + e.persona + ' desde el ' + e.desde + '.');
      if (!cc) throw new Error('El file figura en poder de ' + e.persona + '. Pide a Contratos y Cobranzas que lo corrija.');
    }
    const ahora = new Date(); fubHoja_().appendRow(fubRetiro_(u, ex, persona, motivo, cc, ahora));
    log_(u.nombre, 'FILE_RETIRO', ex.id + ' · ' + persona + ' · ' + motivo);
    return { id: ex.id, persona: persona, confirmado: cc };
  } finally { lock.releaseLock(); }
}
/** Contratos: registrar de una vez varios files en poder de una persona (p. ej. la carga inicial). */
function fileRetirarVarios(token, ids, persona, motivo) {
  const u = sesion_(token); if (!fubCC_(u)) throw new Error('Solo Contratos y Cobranzas puede registrar varios files a la vez.');
  persona = String(persona || '').trim(); motivo = String(motivo || '').trim() || 'Subsanar observación';
  if (persona.length < 3) throw new Error('Indica quién tiene los files.');
  ids = (ids || []).map(x => String(x).trim()).filter(Boolean); if (!ids.length) throw new Error('Elige al menos un expediente.');
  if (ids.length > 100) throw new Error('Máximo 100 a la vez.');
  const exps = misExpedientes(token), lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const est = fubEstados_(), ahora = new Date(), filas = [], omitidos = [];
    ids.forEach(id => {
      const ex = exps.find(x => x.id === id); if (!ex) { omitidos.push(id + ' (no existe)'); return; }
      const e = est[id]; if (e && e.estado !== 'EN CONTRATOS' && norm_(e.persona) === norm_(persona)) { omitidos.push(id + ' (ya registrado)'); return; }
      filas.push(fubRetiro_(u, ex, persona, motivo, true, ahora));
    });
    if (filas.length) { const h = fubHoja_(); h.getRange(h.getLastRow() + 1, 1, filas.length, FUB_HDR.length).setValues(filas); }
    log_(u.nombre, 'FILE_RETIRO', filas.length + ' files · ' + persona + ' · ' + motivo);
    return { registrados: filas.length, omitidos: omitidos };
  } finally { lock.releaseLock(); }
}
/** Registrar que el file volvió a Contratos. Si lo registra Comercial, queda por recibir hasta que Contratos confirme. */
function fileDevolver(token, id, nota) {
  const u = sesion_(token), ex = fubAcceso_(token, id), cc = fubCC_(u);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const e = fubEstados_()[ex.id];
    if (!e || e.estado === 'EN CONTRATOS') throw new Error('Este file ya figura en Contratos.');
    if (e.estado === 'DEVOLVIENDO' && !cc) throw new Error('Ya registraste la devolución: falta que Contratos confirme la recepción.');
    const ahora = new Date();
    if (e.estado === 'DEVOLVIENDO') {   // Contratos confirma la devolución que avisó Comercial
      fubHoja_().getRange(e.fila, 10, 1, 3).setValues([['RECIBIDO', u.nombre, ahora]]);
    } else {
      fubHoja_().appendRow([ahora, ex.id, seguro_(ex.cliente), seguro_(ex.proyecto), seguro_(ex.unidad), 'DEVOLUCION', seguro_(e.persona), u.nombre, seguro_(String(nota || '').trim()),
        cc ? 'RECIBIDO' : 'POR RECIBIR', cc ? u.nombre : '', cc ? ahora : '']);
    }
    log_(u.nombre, 'FILE_DEVOLUCION', ex.id + ' · ' + e.persona + (cc ? ' · recibido' : ' · por recibir'));
    return { id: ex.id, recibido: cc };
  } finally { lock.releaseLock(); }
}
/** Contratos: confirma (ok=true) o rechaza (ok=false) el último movimiento que registró Comercial. */
function fileConfirmar(token, id, ok) {
  const u = sesion_(token); if (!fubCC_(u)) throw new Error('Solo Contratos y Cobranzas confirma los movimientos del file.');
  const ex = fubAcceso_(token, id), lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const e = fubEstados_()[ex.id]; if (!e || !e.porConfirmar) throw new Error('No hay nada por confirmar en este file.');
    const ahora = new Date(), nuevo = !ok ? 'ANULADO' : e.estado === 'FUERA' ? 'CONFIRMADO' : 'RECIBIDO';
    fubHoja_().getRange(e.fila, 10, 1, 3).setValues([[nuevo, u.nombre, ahora]]);
    log_(u.nombre, 'FILE_CONFIRMA', ex.id + ' · ' + nuevo);
    return { id: ex.id, estado: nuevo };
  } finally { lock.releaseLock(); }
}
