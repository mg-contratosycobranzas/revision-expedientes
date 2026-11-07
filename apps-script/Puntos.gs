/**
 * SUBSANACIÓN POR PUNTO — archivo independiente (se conecta solo con Code.gs, igual que Comisiones.gs).
 * Los puntos salen de:
 *   - las observaciones de SBS y de Contractual, separadas con "/"  (ej.: "ROP SIN NACIONALIDAD / DJ SIN FECHA")
 *   - los documentos marcados OBSERVADO en CIERRE y SBS
 * El portal arma la lista con lo que ya trae "detalle"; aquí solo se guarda el estado de cada punto (hoja SUBSANACION_PUNTOS):
 *   PENDIENTE → (Comercial) POR VERIFICAR → (Contratos) CONFORME  ó  vuelve a PENDIENTE con su comentario
 * Comercial envía VARIOS puntos en un solo envío (una sola espera) y queda el historial normal en SUBSANACIONES.
 */
const HOJA_PUN = 'SUBSANACION_PUNTOS';
const PUN_HDR = ['ID expediente', 'Área', 'Punto', 'Estado', 'Subsanó', 'Fecha subsanación', 'Comentario de Comercial', 'Verificó', 'Fecha verificación', 'Comentario de Contratos'];
try {
  ['puntosEstado', 'puntosSubsanar', 'puntosVerificar'].forEach(f => { if (API_PUBLICA.indexOf(f) < 0) API_PUBLICA.push(f); });
  if (SOLO_LECTURA.indexOf('puntosEstado') < 0) SOLO_LECTURA.push('puntosEstado');
} catch (e) {}

const punClave_ = (area, punto) => norm_(area) + '|' + norm_(punto).replace(/\s+/g, ' ');
const punFecha_ = d => d instanceof Date ? Utilities.formatDate(d, 'America/Lima', 'dd/MM/yyyy HH:mm') : txt_(d);
function punHoja_() {
  let h = ss_().getSheetByName(HOJA_PUN);
  if (!h) {
    h = ss_().insertSheet(HOJA_PUN);
    h.getRange(1, 1, 1, PUN_HDR.length).setValues([PUN_HDR]).setFontWeight('bold').setBackground('#0C1A3D').setFontColor('#FFFFFF');
    h.setFrozenRows(1); h.setColumnWidth(3, 320); h.setColumnWidth(7, 260); h.setColumnWidth(10, 260);
  }
  return h;
}
/** Solo quien puede ver el expediente (misExpedientes ya filtra por vendedor / jefe y está en caché). */
function punAcceso_(token, id) {
  id = String(id || '').trim();
  const ex = misExpedientes(token).find(x => x.id === id);
  if (!ex) throw new Error('No tiene acceso a este expediente.');
  return ex;
}
function punFilas_(id) {
  const h = punHoja_(), n = h.getLastRow();
  if (n < 2) return [];
  return h.getRange(2, 1, n - 1, PUN_HDR.length).getValues().map((r, i) => ({ fila: i + 2, r: r })).filter(x => String(x.r[0]).trim() === id);
}
/** Guarda (o actualiza) el estado de varios puntos de un expediente. cambios: [{area, punto, valores:{col: valor}}] */
function punGuardar_(id, cambios) {
  const h = punHoja_(), ya = {};
  punFilas_(id).forEach(x => { ya[punClave_(x.r[1], x.r[2])] = x.fila; });
  const nuevas = [];
  cambios.forEach(c => {
    const k = punClave_(c.area, c.punto), f = ya[k];
    if (f) Object.keys(c.valores).forEach(col => h.getRange(f, Number(col)).setValue(seguro_(c.valores[col])));
    else { const row = [id, c.area, c.punto, '', '', '', '', '', '', '']; Object.keys(c.valores).forEach(col => row[Number(col) - 1] = c.valores[col]); nuevas.push(row); }
  });
  if (nuevas.length) h.getRange(h.getLastRow() + 1, 1, nuevas.length, PUN_HDR.length).setValues(nuevas.map(fila_));
}
const punLista_ = lista => (Array.isArray(lista) ? lista : []).map(p => ({ area: String(p.area || '').trim().toUpperCase(), punto: String(p.punto || '').trim().slice(0, 300), estado: String(p.estado || '').toUpperCase(), coment: String(p.coment || '').trim().slice(0, 500) }))
  .filter(p => p.punto && (p.area === 'SBS' || p.area === 'CIERRE'));

/** Estado guardado de los puntos del expediente (el portal lo cruza con las observaciones que ya tiene). */
function puntosEstado(token, id) {
  sesion_(token); punAcceso_(token, id);
  return punFilas_(String(id).trim()).map(x => ({ area: txt_(x.r[1]), punto: txt_(x.r[2]), estado: txt_(x.r[3]) || 'PENDIENTE', subsano: txt_(x.r[4]), fechaSub: punFecha_(x.r[5]),
    comentario: txt_(x.r[6]), verifico: txt_(x.r[7]), fechaVer: punFecha_(x.r[8]), comentarioCC: txt_(x.r[9]) }));
}

/** Comercial: marca uno o varios puntos como subsanados en UN solo envío (queda el historial normal y el aviso a Contratos). */
function puntosSubsanar(token, id, lista, texto, adjunto) {
  const u = sesion_(token); puede_(u, 'subsanar'); punAcceso_(token, id);
  const ps = punLista_(lista); if (!ps.length) throw new Error('Marca al menos un punto que ya subsanaste.');
  texto = String(texto || '').trim();
  const item = ps.length === 1 ? ps[0].area + ' · ' + ps[0].punto : 'Varios puntos (' + ps.length + ')';
  const det = 'Subsané: ' + ps.map(p => p.punto).join(' / ') + (texto ? '\n✍️ ' + texto : '');
  const r = subsanar(token, id, item, det, adjunto);   // historial en SUBSANACIONES, estado 🟠, plazo y correo a Contratos (Code.gs)
  const ahora = new Date();
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { punGuardar_(String(id).trim(), ps.map(p => ({ area: p.area, punto: p.punto, valores: { 4: 'POR VERIFICAR', 5: u.nombre, 6: ahora, 7: texto, 8: '', 9: '', 10: '' } }))); }
  finally { lock.releaseLock(); }
  r.puntos = ps.length;
  return r;
}

/** Contratos y Cobranzas: CONFORME (sale de pendientes) u OBSERVADO (vuelve a Comercial con el comentario). */
function puntosVerificar(token, id, lista, comentario) {
  const u = sesion_(token); puede_(u, 'revisar'); punAcceso_(token, id);
  const ps = punLista_(lista).filter(p => p.estado === 'CONFORME' || p.estado === 'OBSERVADO');
  if (!ps.length) throw new Error('Marca Conforme o Sigue observado en al menos un punto.');
  comentario = String(comentario || '').trim();
  const ahora = new Date(), ok = ps.filter(p => p.estado === 'CONFORME'), obs = ps.filter(p => p.estado === 'OBSERVADO');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    punGuardar_(String(id).trim(), ps.map(p => ({ area: p.area, punto: p.punto,
      valores: { 4: p.estado === 'CONFORME' ? 'CONFORME' : 'PENDIENTE', 8: u.nombre, 9: ahora, 10: p.coment || (p.estado === 'OBSERVADO' ? comentario : '') } })));
  } finally { lock.releaseLock(); }
  // mensaje en la conversación: cierra la subsanación pendiente y, si algo sigue observado, el expediente vuelve a Comercial
  const partes = [];
  if (ok.length) partes.push('✅ Conforme: ' + ok.map(p => p.punto).join(', ') + '.');
  if (obs.length) partes.push('🔴 Sigue observado: ' + obs.map(p => p.punto + (p.coment ? ' (' + p.coment + ')' : '')).join(', ') + '.');
  if (comentario) partes.push(comentario);
  try { subsanar(token, id, 'Revisión de puntos', partes.join(' '), null); } catch (e) { log_(u.nombre, 'PUNTOS_AVISO', id + ' · ' + e.message); }
  log_(u.nombre, 'PUNTOS_VERIFICADOS', id + ' · ' + ok.length + ' conformes · ' + obs.length + ' siguen observados');
  return { conformes: ok.map(p => p.punto), observados: obs.map(p => p.punto) };
}
