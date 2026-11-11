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
/* Historial para reportes: SUBSANACION_PUNTOS solo guarda el último estado; esta hoja solo agrega filas y conserva cada ronda. */
const HOJA_EVT = 'SUBSANACION_EVENTOS';
const EVT_HDR = ['Fecha', 'ID expediente', 'Mes de cierre', 'Área', 'Punto', 'Evento', 'Ronda', 'Usuario', 'Plazo', 'Comentario'];
try {
  ['puntosEstado', 'puntosSubsanar', 'puntosVerificar', 'puntosReglas', 'puntosTablero'].forEach(f => { if (API_PUBLICA.indexOf(f) < 0) API_PUBLICA.push(f); });
  ['puntosEstado', 'puntosReglas', 'puntosTablero'].forEach(f => { if (SOLO_LECTURA.indexOf(f) < 0) SOLO_LECTURA.push(f); });
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
/** Crea la hoja de historial y le copia lo que ya estaba en SUBSANACION_PUNTOS, para no perder lo enviado antes. */
function punEvtHoja_() {
  let h = ss_().getSheetByName(HOJA_EVT);
  if (h) return h;
  h = ss_().insertSheet(HOJA_EVT);
  h.getRange(1, 1, 1, EVT_HDR.length).setValues([EVT_HDR]).setFontWeight('bold').setBackground('#0C1A3D').setFontColor('#FFFFFF');
  h.setFrozenRows(1); h.setColumnWidth(5, 320); h.setColumnWidth(10, 260);
  const p = ss_().getSheetByName(HOJA_PUN), n = p ? p.getLastRow() : 0, previas = [];
  if (n > 1) p.getRange(2, 1, n - 1, PUN_HDR.length).getValues().forEach(r => {
    if (r[5] instanceof Date) previas.push([r[5], r[0], '', r[1], r[2], 'SUBSANADO', 1, r[4], 'NO REGISTRADO', r[6]]);
    if (r[8] instanceof Date && norm_(r[3]) === 'CONFORME') previas.push([r[8], r[0], '', r[1], r[2], 'CONFORME', 1, r[7], '', r[9]]);
  });
  if (previas.length) h.getRange(2, 1, previas.length, EVT_HDR.length).setValues(previas.map(fila_));
  return h;
}
/** Agrega una fila por punto. Si falla, solo queda en el registro: nunca debe impedir subsanar o verificar. Llamar dentro del lock. */
function punHistorial_(id, mes, eventos) {
  try {
    const h = punEvtHoja_(), n = h.getLastRow(), veces = {}, ahora = new Date();
    if (n > 1) h.getRange(2, 2, n - 1, 5).getValues().forEach(r => {
      if (String(r[0]).trim() === id && norm_(r[4]) === 'SUBSANADO') { const k = punClave_(r[2], r[3]); veces[k] = (veces[k] || 0) + 1; }
    });
    const filas = eventos.map(e => {
      const k = punClave_(e.area, e.punto);
      if (e.evento === 'SUBSANADO') veces[k] = (veces[k] || 0) + 1;
      return [ahora, id, mes || '', e.area, e.punto, e.evento, veces[k] || 0, e.usuario, e.plazo || '', e.coment || ''];
    });
    if (filas.length) h.getRange(n + 1, 1, filas.length, EVT_HDR.length).setValues(filas.map(fila_));
  } catch (e) { try { log_('SISTEMA', 'PUNTOS_HISTORIAL_ERROR', id + ' · ' + e.message); } catch (z) {} }
}
const punLista_ = lista => (Array.isArray(lista) ? lista : []).map(p => ({ area: String(p.area || '').trim().toUpperCase(), punto: String(p.punto || '').trim().slice(0, 300), estado: String(p.estado || '').toUpperCase(), coment: String(p.coment || '').trim().slice(0, 500) }))
  .filter(p => p.punto && (p.area === 'SBS' || p.area === 'CIERRE'));

/** Qué reglas tiene este motor (el portal solo habilita lo que el motor instalado permite). */
function puntosReglas() { return { esperaPorPunto: true, tablero: true }; }

/** TABLERO: para el resumen de cada tarjeta (🔴 CIERRE 4 obs. · 🟠 SBS 2 por verificar…). Solo lectura.
 *  Por expediente observado / subsanado: los documentos en OBSERVADO de la ficha (mismo nombre que en "Subsanar por punto")
 *  y la etapa guardada de cada punto. Las observaciones escritas ya vienen en la lista de expedientes. */
function puntosTablero(token) {
  sesion_(token);
  const out = {}, L = s => String(s || '').replace(/\s+/g, ' ').trim();
  misExpedientes(token).forEach(x => { if (x.estado === 'OBSERVADO' || x.estado === 'SUBSANADO') out[x.id] = { it: [], st: [] }; });
  if (!Object.keys(out).length) return out;
  if (typeof FICHA_HOJAS !== 'undefined') Object.keys(FICHA_HOJAS).forEach(area => {
    try {
      const cfg = FICHA_HOJAS[area], sh = ss_().getSheetByName(cfg.hoja); if (!sh || sh.getLastRow() < cfg.f0) return;
      const lc = sh.getLastColumn(), h4 = sh.getRange(4, 1, 1, lc).getDisplayValues()[0].map(L), h5 = area === 'SBS' ? sh.getRange(5, 1, 1, lc).getDisplayValues()[0].map(L) : h4.map(() => '');
      const g4 = h4.slice(); sh.getRange(4, 1, 1, lc).getMergedRanges().forEach(m => { const v = L(m.getDisplayValue()); for (let c = m.getColumn(); c <= m.getLastColumn(); c++) g4[c - 1] = v; });
      const nombre = t => t.replace(/^TITU\.?$/i, 'Titular').replace(/^C[OÓ]NY\.?$/i, 'Cónyuge'), tit = [];
      for (let col = cfg.c0; col < cfg.c0 + cfg.n; col++) {   // mismo nombre que arma la ficha (leerFicha_ / detalleCompleto_)
        const k = col - 1, sub = area === 'SBS' ? h5[k] : '', grupo = area === 'SBS' ? g4[k] : '', t = nombre(sub || h4[k] || h5[k] || ('Columna ' + col));
        tit.push((grupo && grupo !== t ? grupo + ' · ' : '') + t);
      }
      const n = sh.getLastRow() - cfg.f0 + 1, ids = sh.getRange(cfg.f0, 1, n, 1).getDisplayValues(), vals = sh.getRange(cfg.f0, cfg.c0, n, cfg.n).getDisplayValues();
      ids.forEach((r, i) => { const o = out[String(r[0]).trim()]; if (!o) return; vals[i].forEach((v, k) => { if (norm_(v) === 'OBSERVADO') o.it.push([area, tit[k]]); }); });
    } catch (e) {}
  });
  const h = ss_().getSheetByName(HOJA_PUN), np = h ? h.getLastRow() : 0;
  if (np > 1) h.getRange(2, 1, np - 1, 4).getValues().forEach(r => { const o = out[String(r[0]).trim()]; if (o) o.st.push([txt_(r[1]), txt_(r[2]), txt_(r[3]) || 'PENDIENTE']); });
  return out;
}

/** Estado guardado de los puntos del expediente (el portal lo cruza con las observaciones que ya tiene). */
function puntosEstado(token, id) {
  sesion_(token); punAcceso_(token, id);
  return punFilas_(String(id).trim()).map(x => ({ area: txt_(x.r[1]), punto: txt_(x.r[2]), estado: txt_(x.r[3]) || 'PENDIENTE', subsano: txt_(x.r[4]), fechaSub: punFecha_(x.r[5]),
    comentario: txt_(x.r[6]), verifico: txt_(x.r[7]), fechaVer: punFecha_(x.r[8]), comentarioCC: txt_(x.r[9]) }));
}

/** Comercial: marca uno o varios puntos como subsanados en UN solo envío (queda el historial normal y el aviso a Contratos). */
function puntosSubsanar(token, id, lista, texto, adjunto) {
  const u = sesion_(token); puede_(u, 'subsanar'); const ex = punAcceso_(token, id);
  const ps = punLista_(lista); if (!ps.length) throw new Error('Marca al menos un punto que ya subsanaste.');
  // 🔒 Bloqueo.gs: espera de revisión o plazo vencido. Si la espera es por OTROS puntos ya enviados, se pueden enviar los puntos nuevos
  //    (los que aún no están "POR VERIFICAR"); lo que no se puede es reenviar un punto que ya espera revisión.
  const yaEst = {}; punFilas_(String(id).trim()).forEach(x => { yaEst[punClave_(x.r[1], x.r[2])] = norm_(x.r[3]); });
  const repetidos = ps.filter(p => yaEst[punClave_(p.area, p.punto)] === 'POR VERIFICAR');
  if (repetidos.length) throw new Error('Ya enviaste y está por verificar: ' + repetidos.map(p => p.punto).join(', ') + '. Espera la revisión de Contratos.');
  const eHab = typeof habCalc_ === 'function' ? habCalc_(u, punAcceso_(token, id)) : null;
  if (eHab && eHab.bloqueado && eHab.motivo === 'ESPERA' && eHab.vencido && !eHab.excepcion && !(eHab.habilitacion && eHab.habilitacion.tipo === 'HABILITADO FUERA DE FECHA'))
    throw new Error('⏰ Venció el plazo para presentar subsanados (' + eHab.limite + '). Para enviar más puntos usa "Pedir habilitación".');
  const hb = eHab && eHab.bloqueado && eHab.motivo === 'ESPERA' ? null : (typeof habAntesDeSubsanar_ === 'function' ? habAntesDeSubsanar_(token, id) : null);
  texto = String(texto || '').trim();
  const item = ps.length === 1 ? ps[0].area + ' · ' + ps[0].punto : 'Varios puntos (' + ps.length + ')';
  const det = 'Subsané: ' + ps.map(p => p.punto).join(' / ') + (texto ? '\n✍️ ' + texto : '') + (hb && !habEsCC_(u) ? habNota_(hb.hab) : '');
  const r = subsanar(token, id, item, det, adjunto);   // historial en SUBSANACIONES, estado 🟠, plazo y correo a Contratos (Code.gs)
  if (hb && !habEsCC_(u)) habUsar_(String(id).trim(), hb.hab);
  const ahora = new Date();
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    punGuardar_(String(id).trim(), ps.map(p => ({ area: p.area, punto: p.punto, valores: { 4: 'POR VERIFICAR', 5: u.nombre, 6: ahora, 7: texto, 8: '', 9: '', 10: '' } })));
    const plazo = eHab ? (eHab.vencido ? 'FUERA DE PLAZO' : 'EN PLAZO') : '';
    punHistorial_(String(id).trim(), ex.mes, ps.map(p => ({ area: p.area, punto: p.punto, evento: 'SUBSANADO', usuario: u.nombre, plazo: plazo, coment: texto })));
  }
  finally { lock.releaseLock(); }
  r.puntos = ps.length;
  return r;
}

/** Contratos y Cobranzas: CONFORME (sale de pendientes) u OBSERVADO (vuelve a Comercial con el comentario). */
function puntosVerificar(token, id, lista, comentario) {
  const u = sesion_(token); puede_(u, 'revisar'); const ex = punAcceso_(token, id);
  const ps = punLista_(lista).filter(p => p.estado === 'CONFORME' || p.estado === 'OBSERVADO');
  if (!ps.length) throw new Error('Marca Conforme o Sigue observado en al menos un punto.');
  comentario = String(comentario || '').trim();
  const ahora = new Date(), ok = ps.filter(p => p.estado === 'CONFORME'), obs = ps.filter(p => p.estado === 'OBSERVADO');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    punGuardar_(String(id).trim(), ps.map(p => ({ area: p.area, punto: p.punto,
      valores: { 4: p.estado === 'CONFORME' ? 'CONFORME' : 'PENDIENTE', 8: u.nombre, 9: ahora, 10: p.coment || (p.estado === 'OBSERVADO' ? comentario : '') } })));
    punHistorial_(String(id).trim(), ex.mes, ps.map(p => ({ area: p.area, punto: p.punto, evento: p.estado === 'CONFORME' ? 'CONFORME' : 'OBSERVADO DE NUEVO', usuario: u.nombre,
      coment: p.coment || (p.estado === 'OBSERVADO' ? comentario : '') })));
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
