/**
 * COMISIONES DEL CIERRE — archivo independiente (se conecta solo con Code.gs, igual que Ruta.gs).
 * 1) Administración: "Generar" lee BASE MAESTRA del cierre elegido y arma:
 *    - COMISIONES_DETALLE: una línea por departamento vendido (monto a la minuta y pendiente para la escritura)
 *    - COMISIONES_RESUMEN: una fila por vendedor, jefe de ventas y gerente comercial
 *    Reglas: la escala es la de "💰 Escala de comisiones" (COMISIONES en Code.gs). Solo cuentan DEPARTAMENTOS para el tramo.
 *    Financiamiento: AHORRO = plan de ahorros (50/50) · DIRECTO / HIPOTECARIO / CRÉDITO = 70/30 · CONTADO = 100%.
 *    Escrituras, bonos, adelantos y descuentos se llenan A MANO en las hojas antes de publicar.
 * 2) "Exportar" arma el formato del CIERRE DE MES (una pestaña por jefe, un bloque por vendedor) y lo descarga en Excel o PDF.
 * 3) "Publicar": cada persona ve en el portal SOLO sus montos (vendedor: sus líneas; jefe: su comisión de equipo; gerente: la suya).
 */
const HOJA_COM_DET = 'COMISIONES_DETALLE', HOJA_COM_RES = 'COMISIONES_RESUMEN', CARPETA_COM = 'COMISIONES - MG';
const COM_DET_HDR = ['Cierre', 'Jefe de ventas', 'Vendedor', 'Dpto', 'Est / Dep', 'Proyecto', 'Cliente', 'Fecha minuta', 'Fecha escritura', 'Tipo financ.', 'Monto comisión S/', 'Pendiente S/', 'Concepto', 'ID expediente', 'Origen', 'Subsanación'];
// Ventas que pasan al mes siguiente porque se subsanaron fuera de fecha (o no se subsanaron): se pagan en el cierre "Pasa a"
const HOJA_COM_DIF = 'COMISIONES_DIFERIDOS';
// Decisión por venta (PAGA / PASA / AUTO), monto manual y ventas de meses anteriores agregadas (Subsanación = AGREGADA)
const COM_DIF_HDR = ['Cierre origen', 'ID expediente', 'Vendedor', 'Jefe', 'Decisión', 'Pasa a', 'Monto S/', 'Pendiente S/', 'Dpto', 'Proyecto', 'Cliente', 'Fecha minuta', 'Tipo financ.', 'Subsanación', 'Decidió', 'Fecha', 'Monto manual S/'];
const COM_RES_HDR = ['Cierre', 'Rol', 'Persona', 'Jefe de ventas', 'Unidades', 'Tramo', 'Comisión S/', 'Bono S/', 'Adelanto S/', 'Dscto S/', 'T. Pagar S/', 'Pendiente S/', 'Publicado', 'Generado', 'Pagado (fecha)', 'Pago (medio / N° op.)', 'Marcó el pago'];
const COM_FN = ['comEstado', 'comPrevisualizar', 'comDecidir', 'comAjustar', 'comConcepto', 'comQuitarConcepto', 'comAnteriores', 'comAgregarAnterior', 'comPagar', 'comGenerar', 'comPublicar', 'comExportar', 'misComisiones'];
try {
  COM_FN.forEach(f => { if (API_PUBLICA.indexOf(f) < 0) API_PUBLICA.push(f); });
  ['comEstado', 'comPrevisualizar', 'comAnteriores', 'misComisiones'].forEach(f => { if (SOLO_LECTURA.indexOf(f) < 0) SOLO_LECTURA.push(f); });
} catch (e) {}

/* ---------- escala (sale de la misma tabla que ve cada rol en Ayuda) ---------- */
function comRango_(etq) {   // "1 depa" → [1,1] · "2 a 3 depas" → [2,3] · "9 depas a más" / "19 a más" → [9, ∞]
  const n = String(etq).match(/\d+/g) || [];
  if (/m[aá]s/i.test(etq)) return [Number(n[0]), Infinity];
  return [Number(n[0]), Number(n[1] || n[0])];
}
function comTramo_(tabla, unidades) {
  const f = (tabla.f || []).find(r => { const g = comRango_(r[0]); return unidades >= g[0] && unidades <= g[1]; });
  return f || null;
}
function comEscala_() {
  if (typeof COMISIONES === 'undefined' || !COMISIONES.VENDEDOR) throw new Error('No encontré la escala de comisiones en Code.gs.');
  const t = COMISIONES.VENDEDOR.tablas, busca = re => t.find(x => re.test(x.t));
  return { ahorro: busca(/ahorro/i), credito: busca(/cr[eé]dito|directo|hipotec/i), contado: busca(/contado/i), jefe: COMISIONES.JEFE && COMISIONES.JEFE.tablas[0], gerente: COMISIONES.GERENTE && COMISIONES.GERENTE.tablas[0] };
}
/** Modalidad según TIPO FINANCIAMIENTO de BASE MAESTRA. */
function comModalidad_(fin) {
  const f = norm_(fin);
  if (/AHORRO/.test(f)) return 'ahorro';
  if (/CONTADO/.test(f)) return 'contado';
  if (/DIRECTO|HIPOTEC|CR[EÉ]DITO|CHIP/.test(f)) return 'credito';
  return '';
}

/* ---------- hojas ---------- */
function comHojas_() {
  const ss = ss_(); let d = ss.getSheetByName(HOJA_COM_DET), r = ss.getSheetByName(HOJA_COM_RES);
  const hdr = (h, cab) => { h.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground('#0C1A3D').setFontColor('#FFFFFF'); h.setFrozenRows(1); };
  if (!d) { d = ss.insertSheet(HOJA_COM_DET); hdr(d, COM_DET_HDR); d.getRange('H:I').setNumberFormat('dd/mm/yyyy'); d.getRange('K:L').setNumberFormat('#,##0.00'); }
  if (!r) { r = ss.insertSheet(HOJA_COM_RES); hdr(r, COM_RES_HDR); r.getRange('G:L').setNumberFormat('#,##0.00'); }
  if (String(d.getRange(1, COM_DET_HDR.length).getValue()) !== COM_DET_HDR[COM_DET_HDR.length - 1]) hdr(d, COM_DET_HDR);   // columna nueva "Subsanación"
  let f = ss.getSheetByName(HOJA_COM_DIF);
  if (!f) { f = ss.insertSheet(HOJA_COM_DIF); hdr(f, COM_DIF_HDR); f.getRange('G:H').setNumberFormat('#,##0.00'); }
  [[r, COM_RES_HDR], [f, COM_DIF_HDR]].forEach(([h, cab]) => { if (h.getMaxColumns() < cab.length) h.insertColumnsAfter(h.getMaxColumns(), cab.length - h.getMaxColumns()); if (String(h.getRange(1, cab.length).getValue()) !== cab[cab.length - 1]) hdr(h, cab); });   // columnas nuevas (pago / monto manual)
  return { d: d, r: r, f: f };
}
const comSig_ = c => { const p = c.split(' '), m = MESES_.indexOf(p[0]); return m === 11 ? MESES_[0] + ' ' + (Number(p[1]) + 1) : MESES_[m + 1] + ' ' + p[1]; };
const comDM_ = d => d instanceof Date ? Utilities.formatDate(d, 'America/Lima', 'dd/MM') : '';
/** Subsanación de cada expediente del cierre frente al plazo "Comercial presenta subsanados" del CALENDARIO.
 *  SIN OBS. · EN FECHA dd/mm · FUERA dd/mm · SIN SUBSANAR (plazo vencido y sigue observado) · POR SUBSANAR (aún en plazo) */
function comSubs_(token, cierre) {
  const out = {}, su = ss_().getSheetByName(HOJA_SUB), lim = typeof plazo_ === 'function' ? plazo_(cierre, 'X') : null, vencido = !!(lim && new Date() > lim);
  const filas = su && su.getLastRow() >= 2 ? su.getRange(2, 1, su.getLastRow() - 1, 9).getValues() : [];
  const ven = {}; filas.forEach(r => { if (!(r[0] instanceof Date) || !/\(Ventas\)/.test(String(r[4]))) return; (ven[String(r[1]).trim()] = ven[String(r[1]).trim()] || []).push(r); });
  let exps = []; try { exps = misExpedientes(token); } catch (e) {}
  const est = {}; exps.forEach(x => est[x.id] = x);
  return id => {
    if (out[id]) return out[id];
    const l = ven[id] || [], fuera = l.find(r => /Fuera de plazo/i.test(String(r[8])) || (lim && r[0] > lim)), x = est[id] || {};
    let o;
    if (fuera) o = { tipo: 'FUERA', txt: 'FUERA ' + comDM_(fuera[0]) };
    else if (l.length) o = { tipo: 'EN FECHA', txt: 'EN FECHA ' + comDM_(l[l.length - 1][0]) };
    else if (x.estado === 'OBSERVADO') o = vencido ? { tipo: 'SIN SUBSANAR', txt: 'SIN SUBSANAR' } : { tipo: 'POR SUBSANAR', txt: 'POR SUBSANAR' };
    else o = { tipo: 'SIN OBS.', txt: 'SIN OBS.' };
    // REGLA AUTOMÁTICA DE PAGO: se paga si está APROBADO (SBS y Contractual) y no se subsanó fuera de fecha; si no, pasa al siguiente mes.
    const aprobado = x.estado === 'COMPLETO' || (norm_(x.sbs) === 'APROBADO' && norm_(x.ctr) === 'APROBADO');
    o.estado = aprobado ? 'APROBADO' : x.estado === 'SUBSANADO' ? 'SUBSANADO' : x.estado === 'OBSERVADO' ? 'OBSERVADO' : x.estado ? 'PENDIENTE' : 'SIN REVISAR';
    o.pasa = !aprobado || o.tipo === 'FUERA';
    o.razon = !aprobado ? (o.estado === 'SIN REVISAR' ? 'aún sin revisar' : o.estado.toLowerCase()) : o.tipo === 'FUERA' ? 'aprobado, pero subsanó fuera de fecha' : 'aprobado';
    return (out[id] = o);
  };
}
function comDifFilas_(h) { return comFilas_(h.f).map((r, i) => ({ fila: i + 2, r: r })); }

const comAuto_ = q => /^\(autom/i.test(String(q || ''));   // filas que escribió el sistema (no son decisión tuya)
/** Ajuste de UNA venta del cierre: decisión (PAGA = se paga · PASA = pasa al siguiente mes · AUTO = vuelve a la regla) y/o monto manual ('' = quitar). */
function comAjustar(token, cierre, id, cambios) {
  const u = sesion_(token); puede_(u, 'accesos');
  cierre = comCierreOk_(cierre); id = String(id || '').trim(); cambios = cambios || {};
  const h = comHojas_(), x = comDifFilas_(h).find(z => mesNorm_(mesTxt_(z.r[0])) === cierre && String(z.r[1]).trim() === id && norm_(z.r[13]) !== 'AGREGADA');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    let fila = x ? x.fila : 0;
    if (!fila) { fila = h.f.getLastRow() + 1; const row = new Array(COM_DIF_HDR.length).fill(''); row[0] = cierre; row[1] = id; row[4] = 'AUTO'; h.f.getRange(fila, 1).setNumberFormat('@'); h.f.getRange(fila, 6).setNumberFormat('@'); h.f.getRange(fila, 1, 1, row.length).setValues([fila_(row)]); }
    if (!cambios.decision && x && comAuto_(x.r[14])) { h.f.getRange(fila, 5, 1, 2).setValues([['AUTO', '']]); }   // fila automática: al poner solo el monto, la decisión sigue siendo la regla
    if (cambios.decision) { const d = ['PAGA', 'PASA', 'AUTO'].indexOf(cambios.decision) >= 0 ? cambios.decision : 'AUTO'; h.f.getRange(fila, 6).setNumberFormat('@'); h.f.getRange(fila, 5, 1, 2).setValues([[d, d === 'PASA' ? comSig_(cierre) : '']]); }
    if ('monto' in cambios) { const m = cambios.monto === '' || cambios.monto == null ? '' : Number(cambios.monto); if (m !== '' && (isNaN(m) || m < 0)) throw new Error('Escribe un monto válido.'); h.f.getRange(fila, 17).setValue(m); }
    h.f.getRange(fila, 15, 1, 2).setValues([[u.nombre, new Date()]]);
  } finally { lock.releaseLock(); }
  log_(u.nombre, 'COMISION_AJUSTE', cierre + ' · ' + id + ' · ' + JSON.stringify(cambios));
  return { id: id };
}
function comDecidir(token, cierre, id, decision) { return comAjustar(token, cierre, id, { decision: decision === 'PASA' ? 'PASA' : decision === 'AUTO' ? 'AUTO' : 'PAGA' }); }

/** Concepto manual de un vendedor (bono, ajuste; negativo = descuento). Queda como línea MANUAL en COMISIONES_DETALLE. */
function comConcepto(token, cierre, vendedor, jefe, concepto, monto, dpto) {
  const u = sesion_(token); puede_(u, 'accesos'); cierre = comCierreOk_(cierre);
  concepto = String(concepto || '').trim().toUpperCase(); monto = Number(monto);
  if (concepto.length < 3) throw new Error('Escribe el concepto (ej.: BONO 3 DORM).'); if (!monto || isNaN(monto)) throw new Error('Escribe el monto.');
  const h = comHojas_(), cid = 'MAN-' + Date.now(), i = h.d.getLastRow() + 1;
  h.d.getRange(i, 1).setNumberFormat('@'); h.d.getRange(i, 4).setNumberFormat('@');
  h.d.getRange(i, 1, 1, COM_DET_HDR.length).setValues([fila_([cierre, String(jefe || '').trim(), String(vendedor || '').trim(), String(dpto || '').trim(), '', '', '', '', '', '', monto, 0, concepto, cid, 'MANUAL', ''])]);
  log_(u.nombre, 'COMISION_CONCEPTO', cierre + ' · ' + vendedor + ' · ' + concepto + ' · ' + monto);
  return { id: cid };
}
function comQuitarConcepto(token, cierre, cid) {
  const u = sesion_(token); puede_(u, 'accesos'); cierre = comCierreOk_(cierre); cid = String(cid || '');
  if (!/^MAN-\d+$/.test(cid)) throw new Error('Solo se pueden quitar conceptos agregados a mano.');
  comBorrar_(comHojas_().d, r => mesNorm_(mesTxt_(r[0])) === cierre && String(r[13]) === cid);
  log_(u.nombre, 'COMISION_CONCEPTO_QUITADO', cierre + ' · ' + cid);
  return { ok: true };
}

/** Ventas de meses anteriores del vendedor que NUNCA se pagaron (sin línea con monto en ningún cierre y sin pasar a otro mes). */
function comAnteriores(token, cierre, vendedor) {
  const u = sesion_(token); puede_(u, 'accesos'); cierre = comCierreOk_(cierre);
  const k = c => { const p = String(c).split(' '); return Number(p[1]) * 12 + MESES_.indexOf(p[0]); }, hoy = k(cierre), esc = comEscala_(), h = comHojas_();
  const pagadas = {}; comFilas_(h.d).forEach(r => { if ((Number(r[10]) || 0) > 0) pagadas[String(r[13]).trim()] = 1; });
  const movidas = {}; comDifFilas_(h).forEach(z => { if (norm_(z.r[4]) === 'PASA' && String(z.r[5]).trim()) movidas[String(z.r[1]).trim()] = 1; });
  const v = ss_().getSheetByName(HOJA_BM).getRange(PRIMERA_FILA, 1, ULTIMA_FILA - PRIMERA_FILA + 1, 24).getValues();
  const mias = v.filter(r => r[0] && norm_(r[22]) === norm_(vendedor) && /DEPART/i.test(String(r[5])) && /^[A-Z]+ \d{4}$/.test(mesNorm_(mesTxt_(r[3]))));
  return mias.filter(r => { const c = k(mesNorm_(mesTxt_(r[3]))); return c < hoy && c >= hoy - 6; })
    .filter(r => !pagadas[String(r[0]).trim()] && !movidas[String(r[0]).trim()])
    .map(r => { const ci = mesNorm_(mesTxt_(r[3])), n = mias.filter(z => mesNorm_(mesTxt_(z[3])) === ci).length, mod = comModalidad_(r[19]), t = esc[mod], f = t ? comTramo_(t, n) : null;
      const monto = !f ? 0 : mod === 'contado' ? Number(f[1]) || 0 : Number(f[2]) || 0, pend = !f || mod === 'contado' ? 0 : Number(f[3]) || 0;
      return { id: String(r[0]).trim(), cierre: ci, proyecto: txt_(r[4]), unidad: txt_(r[6]), cliente: txt_(r[7]), fin: txt_(r[19]), monto: monto, pend: pend, tramo: comTramoTxt_(n) }; });
}
/** Agrega al cierre una venta anterior no pagada (queda en COMISIONES_DIFERIDOS como AGREGADA: no se puede pagar dos veces). */
function comAgregarAnterior(token, cierre, id, monto) {
  const u = sesion_(token); puede_(u, 'accesos'); cierre = comCierreOk_(cierre); id = String(id || '').trim(); monto = Number(monto);
  if (isNaN(monto) || monto <= 0) throw new Error('Escribe el monto a pagar.');
  const v = ss_().getSheetByName(HOJA_BM).getRange(PRIMERA_FILA, 1, ULTIMA_FILA - PRIMERA_FILA + 1, 24).getValues(), r = v.find(z => String(z[0]).trim() === id);
  if (!r) throw new Error('No encontré ' + id + ' en BASE MAESTRA.');
  const h = comHojas_(), ori = mesNorm_(mesTxt_(r[3]));
  if (comDifFilas_(h).some(z => String(z.r[1]).trim() === id && norm_(z.r[4]) === 'PASA' && String(z.r[5]).trim())) throw new Error(id + ' ya se pasó o agregó a otro cierre.');
  const i = h.f.getLastRow() + 1; h.f.getRange(i, 1).setNumberFormat('@'); h.f.getRange(i, 6).setNumberFormat('@'); h.f.getRange(i, 9).setNumberFormat('@');
  h.f.getRange(i, 1, 1, COM_DIF_HDR.length).setValues([fila_([ori, id, String(r[22]).trim(), String(r[23] || '').trim(), 'PASA', cierre, monto, 0, txt_(r[6]), txt_(r[4]), txt_(r[7]), r[16] instanceof Date ? r[16] : '', txt_(r[19]), 'AGREGADA', u.nombre, new Date(), ''])]);
  log_(u.nombre, 'COMISION_AGREGADA', cierre + ' · ' + id + ' (de ' + ori + ') · ' + monto);
  return { id: id, origen: ori };
}

/** Marca (o quita) el pago de una o varias personas del cierre. personas: [{rol, nombre}] */
function comPagar(token, cierre, personas, fecha, medio, quitar) {
  const u = sesion_(token); puede_(u, 'accesos'); cierre = comCierreOk_(cierre);
  const h = comHojas_(), ks = (personas || []).map(p => norm_(p.rol) + '|' + norm_(p.nombre)); if (!ks.length) throw new Error('Elige a quién marcar.');
  let f = null; if (!quitar) { const m = String(fecha || '').match(/^(\d{4})-(\d{2})-(\d{2})$/); f = m ? new Date(+m[1], +m[2] - 1, +m[3]) : new Date(); }
  let n = 0;
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    comFilas_(h.r).forEach((r, i) => { if (mesNorm_(mesTxt_(r[0])) !== cierre || ks.indexOf(norm_(r[1]) + '|' + norm_(r[2])) < 0) return;
      h.r.getRange(i + 2, 15, 1, 3).setValues([quitar ? ['', '', ''] : [f, seguro_(String(medio || '').trim().slice(0, 120)), u.nombre]]); if (!quitar) h.r.getRange(i + 2, 15).setNumberFormat('dd/mm/yyyy'); n++; });
  } finally { lock.releaseLock(); }
  if (!n) throw new Error('Primero genera las comisiones de ' + cierre + ' (confirma la vista previa).');
  log_(u.nombre, quitar ? 'COMISION_PAGO_QUITADO' : 'COMISION_PAGADA', cierre + ' · ' + n + ' persona(s)' + (medio ? ' · ' + medio : ''));
  return { n: n };
}
const comFilas_ = h => h.getLastRow() < 2 ? [] : h.getRange(2, 1, h.getLastRow() - 1, h.getLastColumn()).getValues();
function comBorrar_(h, pred) {   // borra de abajo hacia arriba las filas que cumplen pred
  const v = comFilas_(h);
  for (let i = v.length - 1; i >= 0; i--) if (pred(v[i])) h.deleteRow(i + 2);
}
/** Cierres disponibles en BASE MAESTRA (los de meses, sin "NO APLICA"). */
function comCierres_() {
  const bm = ss_().getSheetByName(HOJA_BM), v = bm.getRange(PRIMERA_FILA, 4, ULTIMA_FILA - PRIMERA_FILA + 1, 1).getValues();
  const s = {}; v.forEach(r => { const c = mesNorm_(mesTxt_(r[0])); if (/^[A-Z]+ \d{4}$/.test(c)) s[c] = 1; });
  const k = c => { const p = c.split(' '); return Number(p[1]) * 12 + MESES_.indexOf(p[0]); };
  return Object.keys(s).sort((a, b) => k(b) - k(a));
}

/* ---------- administración ---------- */
function comEstado(token) {
  const u = sesion_(token); puede_(u, 'accesos');
  const h = comHojas_(), res = comFilas_(h.r), est = {};
  res.forEach(r => { const c = mesNorm_(mesTxt_(r[0])); if (!c) return; const e = est[c] = est[c] || { cierre: c, personas: 0, total: 0, publicado: false, generado: '' };
    e.personas++; e.total += Number(r[10]) || 0; if (norm_(r[12]) === 'SI') e.publicado = true; if (r[13]) e.generado = txt_(r[13]); });
  return { cierres: comCierres_(), estado: Object.keys(est).map(k => est[k]), hoja: ss_().getUrl() + '#gid=' + h.r.getSheetId(), hojaDetalle: ss_().getUrl() + '#gid=' + h.d.getSheetId() };
}

/** Cálculo del cierre SIN escribir nada (lo usan la vista previa y Generar). */
function comCalcular_(cierre, h, token) {
  const esc = comEscala_(), SUB = comSubs_(token, cierre), sig = comSig_(cierre), info = {};
  // decisiones tuyas por venta (las filas automáticas no cuentan como decisión) y ventas agregadas de meses anteriores
  const difs = comDifFilas_(h), decide = {};
  difs.forEach(z => { if (mesNorm_(mesTxt_(z.r[0])) !== cierre) return; const id = String(z.r[1]).trim(), mm = z.r[16];
    decide[id] = { dec: comAuto_(z.r[14]) ? 'AUTO' : norm_(z.r[4]) || 'AUTO', dest: mesNorm_(mesTxt_(z.r[5])), agregada: norm_(z.r[13]) === 'AGREGADA', quien: comAuto_(z.r[14]) ? '' : txt_(z.r[14]),
      monto: mm === '' || mm == null || isNaN(Number(mm)) ? null : Number(mm) }; });
  const previos = comFilas_(h.r).filter(r => mesNorm_(mesTxt_(r[0])) === cierre);
  // lo manual se conserva: bono / adelanto / dscto por persona, el pago marcado y las líneas del detalle que no son AUTO (escrituras, conceptos)
  const manual = {}, pago = {}; previos.forEach(r => { const k = norm_(r[1]) + '|' + norm_(r[2]); manual[k] = [r[7], r[8], r[9]]; pago[k] = [r[14] || '', r[15] || '', r[16] || '']; });

  const bm = ss_().getSheetByName(HOJA_BM), v = bm.getRange(PRIMERA_FILA, 1, ULTIMA_FILA - PRIMERA_FILA + 1, 24).getValues();
  const filas = v.filter(r => r[22] && mesNorm_(mesTxt_(r[3])) === cierre);
  if (!filas.length && !difs.some(z => norm_(z.r[4]) === 'PASA' && mesNorm_(mesTxt_(z.r[5])) === cierre)) throw new Error('No hay ventas con cierre ' + cierre + ' en BASE MAESTRA.');
  const depas = filas.filter(r => /DEPART/i.test(String(r[5]))), otros = filas.filter(r => !/DEPART/i.test(String(r[5])));

  // por vendedor: tramo según la cantidad de departamentos del mes
  const porV = {}; depas.forEach(r => { const k = norm_(r[22]); (porV[k] = porV[k] || { nombre: String(r[22]).trim(), jefe: String(r[23] || '').trim(), filas: [] }).filas.push(r); });
  const det = [], res = [], ahora = new Date(), sinMod = [];
  Object.keys(porV).sort().forEach(k => {
    const p = porV[k], n = p.filas.length;
    p.filas.forEach(r => {
      const mod = comModalidad_(r[19]), tabla = esc[mod], f = tabla ? comTramo_(tabla, n) : null;
      let monto = 0, pend = 0, concepto = 'FIRMA DE MINUTA';
      if (!f) { concepto = 'REVISAR: financiamiento "' + (txt_(r[19]) || 'vacío') + '"'; sinMod.push(String(r[0])); }
      else if (mod === 'contado') monto = Number(f[1]) || 0;
      else { monto = Number(f[2]) || 0; pend = Number(f[3]) || 0; }
      // estacionamiento / depósito del mismo cliente y proyecto del vendedor: se muestra en la columna Est / Dep
      const est = otros.filter(o => norm_(o[22]) === k && norm_(o[7]) === norm_(r[7]) && norm_(o[4]) === norm_(r[4])).map(o => txt_(o[6])).join(' / ');
      // subsanación fuera de fecha (o sin subsanar): por defecto pasa al siguiente mes; la decisión manual manda
      const id = String(r[0]).trim(), s = SUB(id), dd = decide[id], man = !!(dd && dd.dec !== 'AUTO'), dec = man ? dd.dec : (s.pasa ? 'PASA' : 'PAGA');
      const dest = dd && dd.dest ? dd.dest : sig, mm = dd ? dd.monto : null, pagar = mm != null ? mm : monto;
      info[id] = { subs: s.txt, tipo: s.tipo, estado: s.estado, razon: s.razon, decision: dec, manual: man, quien: dd ? dd.quien : '', defecto: s.pasa ? 'PASA' : 'PAGA', monto: monto, pend: pend, montoManual: mm, pasaA: dest, agregada: !!(dd && dd.agregada) };
      if (dec === 'PASA' && (monto || pend)) concepto = (dd && dd.agregada ? 'SE PAGA EN ' : 'PASA A ') + dest + ' (S/ ' + (monto + pend).toFixed(2) + ')';
      det.push([cierre, p.jefe, p.nombre, txt_(r[6]), est, txt_(r[4]), txt_(r[7]), r[16] instanceof Date ? r[16] : '', '', txt_(r[19]), dec === 'PASA' ? 0 : pagar, dec === 'PASA' ? 0 : pend, concepto, id, 'AUTO', s.txt]);
    });
    const m = manual['VENDEDOR|' + k] || ['', '', ''];   // comisión, pendiente y T. Pagar se ponen como fórmulas al escribir
    res.push([cierre, 'VENDEDOR', p.nombre, p.jefe, n, comTramoTxt_(n), null, m[0], m[1], m[2], null, null, 'NO', ahora]);
  });
  // ventas que vienen del mes anterior (pasaron por subsanación fuera de fecha): se pagan aquí con su monto original
  const llegan = difs.filter(z => norm_(z.r[4]) === 'PASA' && mesNorm_(mesTxt_(z.r[5])) === cierre && (Number(z.r[6]) || Number(z.r[7])));
  llegan.forEach(z => {
    const r = z.r, ven = String(r[2]).trim(), jefe = String(r[3]).trim(), k = norm_(ven), id = String(r[1]).trim();
    const ag = norm_(r[13]) === 'AGREGADA', ori = mesNorm_(mesTxt_(r[0]));
    det.push([cierre, jefe, ven, txt_(r[8]), '', txt_(r[9]), txt_(r[10]), r[11] instanceof Date ? r[11] : txt_(r[11]), '', txt_(r[12]), Number(r[6]) || 0, Number(r[7]) || 0, ag ? 'AGREGADA DE ' + ori + ' (no se pagó)' : 'DIFERIDO DE ' + ori, id, 'AUTO', ag ? 'VENTA DE ' + ori : txt_(r[13])]);
    info[id] = info[id] || { subs: ag ? 'VENTA DE ' + ori : txt_(r[13]), tipo: ag ? 'AGREGADA' : 'DIFERIDO', estado: '', razon: ag ? 'venta anterior agregada' : 'viene de ' + ori, decision: 'PAGA', defecto: 'PAGA', monto: Number(r[6]) || 0, pend: Number(r[7]) || 0, diferido: ori, agregada: ag, quien: ag ? txt_(r[14]) : '' };
    if (!porV[k] && !res.some(x => x[1] === 'VENDEDOR' && norm_(x[2]) === k)) { const m = manual['VENDEDOR|' + k] || ['', '', '']; res.push([cierre, 'VENDEDOR', ven, jefe, 0, '— (solo diferidos)', null, m[0], m[1], m[2], null, null, 'NO', ahora]); }
  });
  // jefes de ventas: tramo según TODAS las unidades (departamentos) del equipo en el mes; minuta = unidades × 70% del tramo
  if (esc.jefe) {
    const porJ = {};
    depas.forEach(r => { const j = String(r[23] || '').trim(); if (!j) return; const k = norm_(j); porJ[k] = porJ[k] || { nombre: j, n: 0 }; porJ[k].n++; });
    Object.keys(porJ).sort().forEach(k => {
      const j = porJ[k], f = comTramo_(esc.jefe, j.n), m = manual['JEFE|' + k] || ['', '', ''];
      res.push([cierre, 'JEFE', j.nombre, j.nombre, j.n, f ? f[0] : '', f ? j.n * (Number(f[2]) || 0) : 0, m[0], m[1], m[2], null, f ? j.n * (Number(f[3]) || 0) : 0, 'NO', ahora]);
    });
  }
  // gerente comercial: monto por unidad vendida en la empresa (si hay un usuario con rol GERENTE se usa su nombre)
  if (esc.gerente && depas.length) {
    const ac = ss_().getSheetByName(HOJA_ACCESO), g = ac.getRange(2, 1, Math.max(ac.getLastRow() - 1, 1), 3).getValues().find(r => norm_(r[2]) === 'GERENTE');
    const nombre = g ? String(g[0]).trim() : 'GERENTE COMERCIAL', m = manual['GERENTE|' + norm_(nombre)] || ['', '', ''], porU = Number((esc.gerente.f[0] || [])[1]) || 0;
    res.push([cierre, 'GERENTE', nombre, '', depas.length, 'Por unidad', depas.length * porU, m[0], m[1], m[2], null, 0, 'NO', ahora]);
  }
  res.forEach(x => { const pg = pago[norm_(x[1]) + '|' + norm_(x[2])] || ['', '', '']; x.push(pg[0], pg[1], pg[2]); });   // el pago marcado no se pierde al regenerar
  return { det: det, res: res, depas: depas, porV: porV, sinMod: sinMod, previos: previos, info: info, sig: sig };
}
/** Guarda en COMISIONES_DIFERIDOS las ventas que pasan al siguiente mes, con su monto (para pagarlas allí). */
function comGuardarDiferidos_(h, cierre, c) {
  const difs = comDifFilas_(h), ya = {}, yaR = {}; difs.forEach(z => { if (mesNorm_(mesTxt_(z.r[0])) === cierre && norm_(z.r[13]) !== 'AGREGADA') { ya[String(z.r[1]).trim()] = z.fila; yaR[String(z.r[1]).trim()] = z.r; } });
  c.det.forEach(d => {
    const id = String(d[13]).trim(), inf = c.info[id]; if (!inf || inf.tipo === 'DIFERIDO' || inf.tipo === 'AGREGADA' || inf.agregada) return;
    if (inf.decision !== 'PASA') {   // ya no pasa: si había una fila automática, queda como "PAGA" para que el mes siguiente no la pague otra vez
      if (ya[id] && norm_(yaR[id][4]) === 'PASA' && comAuto_(yaR[id][14])) { h.f.getRange(ya[id], 5, 1, 2).setValues([['PAGA', '']]); }
      return;
    }
    const datos = [cierre, id, d[2], d[1], 'PASA', inf.pasaA || c.sig, inf.monto, inf.pend, d[3], d[5], d[6], d[7], d[9], inf.subs];
    if (ya[id]) { h.f.getRange(ya[id], 6).setNumberFormat('@'); h.f.getRange(ya[id], 1, 1, datos.length).setValues([fila_(datos)]); if (!inf.manual) h.f.getRange(ya[id], 15).setValue('(automático: ' + inf.razon + ')'); }
    else { const i = h.f.getLastRow() + 1; h.f.getRange(i, 1).setNumberFormat('@'); h.f.getRange(i, 6).setNumberFormat('@'); h.f.getRange(i, 1, 1, COM_DIF_HDR.length).setValues([fila_(datos.concat(['(automático: ' + inf.razon + ')', new Date(), '']))]); }
  });
}
const comCierreOk_ = c => { c = mesNorm_(c); if (!/^[A-Z]+ \d{4}$/.test(c)) throw new Error('Elige un cierre válido (por ejemplo OCTUBRE 2026).'); return c; };

/** VISTA PREVIA: lo que saldría al generar, por vendedor (incluye las líneas que agregaste a mano). No escribe en el Sheet. */
function comPrevisualizar(token, cierre) {
  const u = sesion_(token); puede_(u, 'accesos');
  cierre = comCierreOk_(cierre);
  const h = comHojas_(), c = comCalcular_(cierre, h, token), n = x => Number(x) || 0;
  const lineas = c.det.concat(comFilas_(h.d).filter(r => mesNorm_(mesTxt_(r[0])) === cierre && norm_(r[14]) !== 'AUTO'));
  const linea = d => { const inf = norm_(d[14]) === 'AUTO' ? c.info[txt_(d[13])] || {} : {};
    return { dpto: txt_(d[3]), est: txt_(d[4]), proyecto: txt_(d[5]), cliente: txt_(d[6]), minuta: d[7] instanceof Date ? Utilities.formatDate(d[7], 'America/Lima', 'dd/MM/yy') : txt_(d[7]), escritura: txt_(d[8]), fin: txt_(d[9]),
      monto: n(d[10]), pendiente: n(d[11]), concepto: txt_(d[12]), id: txt_(d[13]), manual: norm_(d[14]) !== 'AUTO',
      subs: txt_(d[15]) || inf.subs || '', tipo: inf.tipo || (norm_(d[14]) === 'MANUAL' ? 'CONCEPTO' : ''), estado: inf.estado || '', razon: inf.razon || '', decision: inf.decision || '', manualDec: !!inf.manual, quien: inf.quien || '',
      defecto: inf.defecto || '', montoOrig: inf.monto || 0, pendOrig: inf.pend || 0, montoManual: inf.montoManual == null ? null : inf.montoManual, pasaA: inf.pasaA || '', diferido: inf.diferido || '', agregada: !!inf.agregada }; };
  const personas = c.res.map(x => {
    const v = norm_(x[1]) === 'VENDEDOR', l = v ? lineas.filter(d => norm_(d[2]) === norm_(x[2])).map(linea) : [];
    const com = v ? l.reduce((a, d) => a + d.monto, 0) : n(x[6]), pen = v ? l.reduce((a, d) => a + d.pendiente, 0) : n(x[11]);
    return { rol: x[1], nombre: x[2], jefe: x[3], unidades: x[4], tramo: txt_(x[5]), comision: com, bono: n(x[7]), adelanto: n(x[8]), dscto: n(x[9]),
      pendiente: pen, total: com + n(x[7]) - n(x[8]) - n(x[9]), lineas: l, pagado: x[14] ? { fecha: x[14] instanceof Date ? Utilities.formatDate(x[14], 'America/Lima', 'dd/MM/yyyy') : txt_(x[14]), medio: txt_(x[15]), quien: txt_(x[16]) } : null };
  });
  return { cierre: cierre, departamentos: c.depas.length, vendedores: Object.keys(c.porV).length, revisar: c.sinMod, yaGenerado: c.previos.length > 0,
    publicado: c.previos.some(r => norm_(r[12]) === 'SI'), total: personas.reduce((a, p) => a + p.total, 0), personas: personas, avisadas: comAvisadas_(cierre, c.porV),
    siguiente: c.sig, plazo: (typeof plazo_ === 'function' && plazo_(cierre, 'X')) ? fmtPlazo_(plazo_(cierre, 'X')) : '',
    // filtro: cuántas ventas del cierre se subsanaron dentro y fuera de fecha (una por expediente)
    plazos: Object.keys(c.info).filter(id => c.info[id].tipo !== 'DIFERIDO').map(id => { const d = c.det.find(z => String(z[13]).trim() === id) || [];
      return { id: id, tipo: c.info[id].tipo, subs: c.info[id].subs, vendedor: txt_(d[2]), jefe: txt_(d[1]), cliente: txt_(d[6]), proyecto: txt_(d[5]), dpto: txt_(d[3]), decision: c.info[id].decision }; }) };
}
/** PROYECCIÓN: departamentos avisados en el portal (📣 Avisar venta) en el mes del cierre que aún NO están en BASE MAESTRA.
 *  No se pagan: solo muestran a qué tramo podría llegar cada vendedor si se registran. */
function comAvisadas_(cierre, porV) {
  try {
    if (typeof rutaFilas_ !== 'function' || typeof RC === 'undefined') return [];
    const ini = comPrimerDia_(cierre), fin = new Date(ini.getFullYear(), ini.getMonth() + 1, 1), por = {};
    rutaFilas_().forEach(x => {
      const r = x.r, est = String(r[RC.estado - 1]), f = r[RC.fecha - 1];
      if (['VALIDADO', 'ANULADO'].indexOf(est) >= 0 || !(f instanceof Date) || f < ini || f >= fin || !/DEPART/i.test(String(r[RC.tipo - 1]))) return;
      const k = norm_(r[RC.vendedor - 1]), p = por[k] = por[k] || { vendedor: String(r[RC.vendedor - 1]).trim(), avisadas: 0, detalle: [] };
      p.avisadas++; p.detalle.push(txt_(r[RC.proyecto - 1]) + ' ' + txt_(r[RC.unidad - 1]) + (r[RC.situacion - 1] ? ' · ' + txt_(r[RC.situacion - 1]) : ''));
    });
    return Object.keys(por).sort().map(k => { const reg = porV[k] ? porV[k].filas.length : 0, p = por[k];
      return { vendedor: p.vendedor, registradas: reg, avisadas: p.avisadas, posible: reg + p.avisadas, tramoActual: reg ? comTramoTxt_(reg) : '—', tramoPosible: comTramoTxt_(reg + p.avisadas), detalle: p.detalle }; });
  } catch (e) { return []; }
}

function comGenerar(token, cierre, forzar) {
  const u = sesion_(token); puede_(u, 'accesos');
  cierre = comCierreOk_(cierre);
  const h = comHojas_(), c = comCalcular_(cierre, h, token), det = c.det, res = c.res, depas = c.depas, porV = c.porV, sinMod = c.sinMod;
  if (c.previos.some(r => norm_(r[12]) === 'SI') && !forzar) throw new Error('El cierre ' + cierre + ' ya está PUBLICADO. Despublícalo primero si quieres regenerarlo.');

  // escribe: borra lo AUTO del cierre (lo manual del detalle se queda) y el resumen del cierre
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    comBorrar_(h.d, r => mesNorm_(mesTxt_(r[0])) === cierre && norm_(r[14]) === 'AUTO');
    comBorrar_(h.r, r => mesNorm_(mesTxt_(r[0])) === cierre);
    if (det.length) { const i = h.d.getLastRow() + 1; h.d.getRange(i, 1, det.length, 1).setNumberFormat('@'); h.d.getRange(i, 1, det.length, COM_DET_HDR.length).setValues(det.map(fila_)); }
    comGuardarDiferidos_(h, cierre, c);   // lo que pasa al siguiente mes queda guardado con su monto
    // comisión y pendiente del vendedor = suma de sus líneas del cierre (AUTO + las que agregaste a mano). Se calcula aquí
    // (las fórmulas SUMIFS daban #ERROR! con la configuración en español del Sheet).
    const suma = {}; comFilas_(h.d).forEach(r => { if (mesNorm_(mesTxt_(r[0])) !== cierre) return; const k = norm_(r[2]), s = suma[k] = suma[k] || [0, 0]; s[0] += Number(r[10]) || 0; s[1] += Number(r[11]) || 0; });
    res.forEach(x => { if (x[1] === 'VENDEDOR') { const s = suma[norm_(x[2])] || [0, 0]; x[6] = s[0]; x[11] = s[1]; } });
    if (res.length) {
      const i = h.r.getLastRow() + 1;
      h.r.getRange(i, 1, res.length, 1).setNumberFormat('@');
      h.r.getRange(i, 1, res.length, COM_RES_HDR.length).setValues(res.map(x => x.map(c => c === null ? '' : seguro_(c))));
      // T. Pagar = comisión + bono − adelanto − dscto (fórmula simple: se recalcula si llenas bono / adelanto / dscto a mano)
      res.forEach((x, k) => { const f = i + k; h.r.getRange(f, 11).setFormula('=N(G' + f + ')+N(H' + f + ')-N(I' + f + ')-N(J' + f + ')'); });
    }
  } finally { lock.releaseLock(); }
  log_(u.nombre, 'COMISIONES_GENERADAS', cierre + ' · ' + depas.length + ' departamentos · ' + Object.keys(porV).length + ' vendedores' + (sinMod.length ? ' · revisar financiamiento: ' + sinMod.join(', ') : ''));
  return { cierre: cierre, departamentos: depas.length, vendedores: Object.keys(porV).length, revisar: sinMod, hoja: ss_().getUrl() + '#gid=' + h.r.getSheetId() };
}
const comTramoTxt_ = n => n >= 9 ? '9 depas a más' : n >= 6 ? '6 a 8 depas' : n >= 4 ? '4 a 5 depas' : n >= 2 ? '2 a 3 depas' : '1 depa';

function comPublicar(token, cierre, publicar) {
  const u = sesion_(token); puede_(u, 'accesos');
  cierre = mesNorm_(cierre); const h = comHojas_(), v = comFilas_(h.r); let n = 0;
  v.forEach((r, i) => { if (mesNorm_(mesTxt_(r[0])) === cierre) { h.r.getRange(i + 2, 13).setValue(publicar ? 'SI' : 'NO'); n++; } });
  if (!n) throw new Error('Primero genera las comisiones de ' + cierre + '.');
  log_(u.nombre, publicar ? 'COMISIONES_PUBLICADAS' : 'COMISIONES_OCULTAS', cierre + ' · ' + n + ' personas');
  return { cierre: cierre, personas: n, publicado: !!publicar };
}

/** FORMATO "CIERRE DE MES" (igual al Excel de Contratos y Cobranzas): una pestaña GRUPO <JEFE>, un bloque por vendedor.
 *  Se arma en un Google Sheet de Drive (carpeta "COMISIONES - MG", uno por cierre, se reemplaza al volver a exportar)
 *  y se descarga como Excel (editable) o PDF (para imprimir). Las sumas quedan como fórmulas simples (sin comas). */
const COM_COL = { naranja: '#F8CBAD', morado: '#78206E', amarillo: '#FFC000', verde: '#92D050', aviso: '#FFFF00', rojo: '#FF0000', okBg: '#C6EFCE', okTx: '#006100', malBg: '#FFC7CE', malTx: '#9C0006' };
const COM_SOL = '"S/." #,##0.00;[Red]"S/." -#,##0.00';

function comBloqueVendedor_(sh, f, p, lineas, cierre) {
  const C = COM_COL, primero = comPrimerDia_(cierre);
  // 1) nombre · mes  (columnas B..L: se agregó SUBSANACIÓN después de TIPO FINANC.)
  sh.getRange(f, 2, 1, 11).setBackground(C.naranja);
  sh.getRange(f, 3, 1, 3).merge().setValue(String(p[2]).toUpperCase()).setBackground(C.morado).setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(f, 6).setValue(primero).setNumberFormat('mmm-yy').setBackground(C.morado).setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  f++;
  // 2) cabecera en dos filas
  sh.getRange(f, 2, 2, 11).setValues([
    ['DPTO', 'EST/DEP', 'PROYECTO', 'CLIENTE', 'FECHA DE', 'FECHA DE', 'TIPO', 'SUBSANACIÓN', 'MONTO', 'PENDIENTE', 'CONCEPTO'],
    ['', '', '', '', 'MINUTA', 'ESCRITURA', 'FINANC.', '(PLAZO)', 'COMISION S/.', '', '']])
    .setBackground(C.naranja).setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, false);
  [2, 3, 4, 5, 11, 12].forEach(c => sh.getRange(f, c, 2, 1).merge());
  f += 2;
  // 3) líneas
  const ini = f;
  if (lineas.length) {
    sh.getRange(f, 2, lineas.length, 11).setValues(lineas.map(d => [txt_(d[3]), txt_(d[4]), txt_(d[5]), txt_(d[6]), d[7] instanceof Date ? d[7] : txt_(d[7]),
      d[8] instanceof Date ? d[8] : txt_(d[8]), txt_(d[9]), txt_(d[15]), Number(d[10]) || 0, Number(d[11]) || 0, txt_(d[12])]))
      .setFontSize(9).setBorder(true, true, true, true, true, true).setVerticalAlignment('middle');
    sh.getRange(f, 2, lineas.length, 3).setHorizontalAlignment('center');
    sh.getRange(f, 6, lineas.length, 2).setNumberFormat('dd/mm/yyyy').setHorizontalAlignment('center');
    sh.getRange(f, 8, lineas.length, 2).setHorizontalAlignment('center');
    sh.getRange(f, 10, lineas.length, 2).setNumberFormat(COM_SOL);
    lineas.forEach((d, i) => {   // verde: sin observaciones / en fecha · rojo: fuera de fecha / sin subsanar (y concepto en rojo si pasa al siguiente mes)
      const t = norm_(d[15]); if (!t) return; const mal = /FUERA|SIN SUBSANAR/.test(t);
      sh.getRange(f + i, 9).setBackground(mal ? C.malBg : C.okBg).setFontColor(mal ? C.malTx : C.okTx).setFontWeight(mal ? 'bold' : 'normal');
      if (/^PASA A/.test(norm_(d[12]))) sh.getRange(f + i, 12).setFontColor(C.rojo).setFontWeight('bold');
    });
    f += lineas.length;
  }
  // 4) total de las líneas
  const tot = f;
  sh.getRange(f, 2).setValue('TOTAL').setFontWeight('bold');
  sh.getRange(f, 10).setFormula(lineas.length ? '=SUM(J' + ini + ':J' + (f - 1) + ')' : '=0').setBackground(C.amarillo);
  sh.getRange(f, 11).setFormula(lineas.length ? '=SUM(K' + ini + ':K' + (f - 1) + ')' : '=0');
  sh.getRange(f, 10, 1, 2).setNumberFormat(COM_SOL).setFontWeight('bold').setBorder(true, true, true, true, true, false);
  f += 2;
  // 5) total por pagar
  sh.getRange(f, 2).setValue('TOTAL POR PAGAR').setFontWeight('bold'); f++;
  sh.getRange(f, 5, 1, 5).setValues([['COMISION', 'BONO(*)', 'ADELANTO', 'DSCTO', 'T.PAGAR']]).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(f, 6).setFontColor(C.rojo); sh.getRange(f, 9).setBackground(C.verde); f++;
  sh.getRange(f, 4).setValue(String(p[2]).toUpperCase()).setFontWeight('bold');
  sh.getRange(f, 5).setFormula('=J' + tot);
  sh.getRange(f, 6, 1, 3).setValues([[Number(p[7]) || 0, Number(p[8]) || 0, Number(p[9]) || 0]]);
  sh.getRange(f, 9).setFormula('=E' + f + '+F' + f + '-G' + f + '-H' + f).setBackground(C.verde).setFontWeight('bold');
  sh.getRange(f, 5, 1, 5).setNumberFormat(COM_SOL);
  sh.getRange(f, 10).setValue(p[14] ? 'PAGADO ' + (p[14] instanceof Date ? Utilities.formatDate(p[14], 'America/Lima', 'dd/MM/yy') : txt_(p[14])) : 'PROGRAMADO').setFontColor(p[14] ? '#006100' : C.rojo).setBackground(p[14] ? C.verde : C.aviso).setFontWeight('bold').setHorizontalAlignment('center');
  return f + 3;
}
function comBloqueJefe_(sh, f, jr, cierre) {
  const C = COM_COL;
  sh.getRange(f, 2, 1, 10).setBackground(C.naranja);
  sh.getRange(f, 3, 1, 3).merge().setValue('JEFE DE VENTAS · ' + String(jr[2]).toUpperCase()).setBackground(C.morado).setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(f, 6).setValue(comPrimerDia_(cierre)).setNumberFormat('mmm-yy').setBackground(C.morado).setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  f++;
  sh.getRange(f, 3, 1, 8).setValues([['UNIDADES', 'TRAMO', 'COMISION (70%)', 'PENDIENTE (30%)', 'BONO(*)', 'ADELANTO', 'DSCTO', 'T.PAGAR']])
    .setBackground(C.naranja).setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center').setBorder(true, true, true, true, true, true); f++;
  sh.getRange(f, 3, 1, 7).setValues([[jr[4], txt_(jr[5]), Number(jr[6]) || 0, Number(jr[11]) || 0, Number(jr[7]) || 0, Number(jr[8]) || 0, Number(jr[9]) || 0]]).setBorder(true, true, true, true, true, true).setHorizontalAlignment('center');
  sh.getRange(f, 10).setFormula('=E' + f + '+G' + f + '-H' + f + '-I' + f).setBackground(C.verde).setFontWeight('bold').setBorder(true, true, true, true, false, false);
  sh.getRange(f, 5, 1, 6).setNumberFormat(COM_SOL);
  sh.getRange(f, 11).setValue(jr[14] ? 'PAGADO ' + (jr[14] instanceof Date ? Utilities.formatDate(jr[14], 'America/Lima', 'dd/MM/yy') : txt_(jr[14])) : 'PROGRAMADO').setFontColor(jr[14] ? '#006100' : C.rojo).setBackground(jr[14] ? C.verde : C.aviso).setFontWeight('bold').setHorizontalAlignment('center');
  return f + 3;
}
const comPrimerDia_ = c => { const p = c.split(' '); return new Date(Number(p[1]), MESES_.indexOf(p[0]), 1); };
function comAnchos_(sh) {
  [[1, 20], [2, 70], [3, 60], [4, 110], [5, 290], [6, 85], [7, 85], [8, 100], [9, 110], [10, 110], [11, 100], [12, 230]].forEach(x => sh.setColumnWidth(x[0], x[1]));
  sh.setHiddenGridlines(true); sh.getRange('A:L').setFontFamily('Arial');
}

/** Arma (o reemplaza) el archivo del cierre en Drive. Devuelve el Spreadsheet. */
function comArmar_(cierre) {
  SpreadsheetApp.flush();
  const h = comHojas_(), det = comFilas_(h.d).filter(r => mesNorm_(mesTxt_(r[0])) === cierre), res = comFilas_(h.r).filter(r => mesNorm_(mesTxt_(r[0])) === cierre);
  if (!res.length) throw new Error('Primero genera las comisiones de ' + cierre + '.');
  const nombre = 'COMISIONES - ' + cierre, pr = PropertiesService.getScriptProperties(), clave = 'COM_EXP_' + cierre.replace(/\s/g, '_');
  let libro = null;
  // se reutiliza el archivo del cierre (solo con SpreadsheetApp: no depende de que el servicio Drive esté habilitado en el proyecto)
  try { const id = pr.getProperty(clave); if (id) libro = SpreadsheetApp.openById(id); } catch (e) { libro = null; }
  if (!libro) {
    libro = SpreadsheetApp.create(nombre); pr.setProperty(clave, libro.getId());
    try { const it = DriveApp.getFoldersByName(CARPETA_COM), carpeta = it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA_COM); DriveApp.getFileById(libro.getId()).moveTo(carpeta); }
    catch (e) { /* sin Drive: el archivo queda en "Mi unidad" */ }
  }
  // se vacía: queda una hoja temporal y se borran las demás
  const tmp = libro.insertSheet('_tmp_' + Date.now()); libro.getSheets().forEach(s => { if (s.getSheetId() !== tmp.getSheetId()) libro.deleteSheet(s); });
  const vend = res.filter(r => norm_(r[1]) === 'VENDEDOR'), jefes = [...new Set(vend.map(r => String(r[3] || '').trim() || 'SIN JEFE'))].sort();
  jefes.forEach(j => {
    const sh = libro.insertSheet(('GRUPO ' + j).toUpperCase().slice(0, 90), libro.getSheets().length);
    comAnchos_(sh); let f = 2;
    vend.filter(r => (String(r[3] || '').trim() || 'SIN JEFE') === j).forEach(p => { f = comBloqueVendedor_(sh, f, p, det.filter(d => norm_(d[2]) === norm_(p[2])), cierre); });
    const jr = res.find(r => norm_(r[1]) === 'JEFE' && norm_(r[2]) === norm_(j)); if (jr) f = comBloqueJefe_(sh, f, jr, cierre);
  });
  const g = res.find(r => norm_(r[1]) === 'GERENTE');
  if (g) {
    const sh = libro.insertSheet('GERENTE COMERCIAL', libro.getSheets().length); comAnchos_(sh);
    sh.getRange(2, 2, 1, 11).setBackground(COM_COL.naranja);
    sh.getRange(2, 3, 1, 3).merge().setValue('GERENTE COMERCIAL · ' + String(g[2]).toUpperCase()).setBackground(COM_COL.morado).setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
    sh.getRange(2, 6).setValue(comPrimerDia_(cierre)).setNumberFormat('mmm-yy').setBackground(COM_COL.morado).setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
    sh.getRange(3, 3, 1, 6).setValues([['UNIDADES', 'COMISION', 'BONO(*)', 'ADELANTO', 'DSCTO', 'T.PAGAR']]).setBackground(COM_COL.naranja).setFontWeight('bold').setHorizontalAlignment('center').setBorder(true, true, true, true, true, true);
    sh.getRange(4, 3, 1, 5).setValues([[g[4], Number(g[6]) || 0, Number(g[7]) || 0, Number(g[8]) || 0, Number(g[9]) || 0]]).setHorizontalAlignment('center').setBorder(true, true, true, true, true, true);
    sh.getRange(4, 8).setFormula('=D4+E4-F4-G4').setBackground(COM_COL.verde).setFontWeight('bold').setBorder(true, true, true, true, false, false);
    sh.getRange(4, 4, 1, 5).setNumberFormat(COM_SOL);
  }
  libro.deleteSheet(tmp);
  if (libro.getName() !== nombre) libro.rename(nombre);
  SpreadsheetApp.flush();
  return libro;
}
/** Descarga del archivo de Drive como Excel o PDF (A4 horizontal, ajustado al ancho, sin cuadrícula). */
function comDescarga_(libro, formato, nombre) {
  const base = 'https://docs.google.com/spreadsheets/d/' + libro.getId() + '/export?format=';
  const url = formato === 'pdf'
    ? base + 'pdf&size=A4&portrait=false&scale=2&gridlines=false&printtitle=false&sheetnames=false&pagenum=UNDEFINED&fzr=false&horizontal_alignment=CENTER&top_margin=0.4&bottom_margin=0.4&left_margin=0.3&right_margin=0.3'
    : base + 'xlsx';
  const r = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('Google no pudo convertir el archivo a ' + (formato === 'pdf' ? 'PDF' : 'Excel') + ' (código ' + r.getResponseCode() + '). Intenta de nuevo en un minuto.');
  return { nombre: nombre + (formato === 'pdf' ? '.pdf' : '.xlsx'), tipo: formato === 'pdf' ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', b64: Utilities.base64Encode(r.getBlob().getBytes()) };
}
/** Exporta el cierre en el formato CIERRE DE MES. formato: 'xlsx' (Excel editable) · 'pdf' (imprimir) · vacío (solo el Google Sheet). */
function comExportar(token, cierre, formato) {
  const u = sesion_(token); puede_(u, 'accesos');
  cierre = mesNorm_(cierre); if (!/^[A-Z]+ \d{4}$/.test(cierre)) throw new Error('Elige un cierre válido.');
  formato = formato === 'pdf' || formato === 'xlsx' ? formato : '';
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  let libro; try { libro = comArmar_(cierre); } finally { lock.releaseLock(); }
  log_(u.nombre, 'COMISIONES_EXPORTADAS', cierre + (formato ? ' · ' + formato.toUpperCase() : ' · Google Sheet'));
  const out = { url: libro.getUrl(), nombre: libro.getName() };
  if (formato) out.archivo = comDescarga_(libro, formato, libro.getName());
  return out;
}

/* ---------- interruptor: ¿Comercial puede ver sus comisiones? (apagado por defecto; lo decide la administración) ---------- */
const comVerComercial_ = () => PropertiesService.getScriptProperties().getProperty('COM_VER_COMERCIAL') === 'SI';
/** Sin "modo": consulta (para mostrar u ocultar el botón). Con "modo" SI / NO: lo cambia (solo administración). */
function comVisible(token, modo) {
  const u = sesion_(token);
  if (modo === 'SI' || modo === 'NO') { puede_(u, 'accesos'); PropertiesService.getScriptProperties().setProperty('COM_VER_COMERCIAL', modo); log_(u.nombre, 'COMISIONES_VISIBLES', modo); }
  return { ver: comVerComercial_() || u.permisos.indexOf('accesos') >= 0, activo: comVerComercial_() };
}
try { if (API_PUBLICA.indexOf('comVisible') < 0) API_PUBLICA.push('comVisible'); if (SOLO_LECTURA.indexOf('comVisible') < 0) SOLO_LECTURA.push('comVisible'); } catch (e) {}

/* ---------- lo que ve cada persona (solo lo publicado y solo lo suyo) ---------- */
function misComisiones(token) {
  const u = sesion_(token), yo = [norm_(u.ref), norm_(u.nombre)].filter(Boolean);
  if (!comVerComercial_() && u.permisos.indexOf('accesos') < 0) return [];   // interruptor apagado: Comercial no ve montos
  const ss = ss_(), r = ss.getSheetByName(HOJA_COM_RES), d = ss.getSheetByName(HOJA_COM_DET); if (!r) return [];
  const det = d ? comFilas_(d) : [], out = [];
  r.getDataRange().getValues().slice(1).forEach(x => {
    if (norm_(x[12]) !== 'SI') return;
    const rol = norm_(x[1]), quien = norm_(x[2]);
    const mio = u.rol === 'ADMIN' ? false : (rol === 'VENDEDOR' && yo.indexOf(quien) >= 0) || (rol === 'JEFE' && u.rol === 'JEFE' && yo.indexOf(quien) >= 0) || (rol === 'GERENTE' && u.rol === 'GERENTE');
    if (!mio) return;
    const cierre = mesNorm_(mesTxt_(x[0]));
    out.push({ cierre: cierre, rol: rol, unidades: x[4], tramo: String(x[5]), comision: Number(x[6]) || 0, bono: Number(x[7]) || 0, adelanto: Number(x[8]) || 0, dscto: Number(x[9]) || 0, total: Number(x[10]) || 0, pendiente: Number(x[11]) || 0, pagado: x[14] ? (x[14] instanceof Date ? Utilities.formatDate(x[14], 'America/Lima', 'dd/MM/yyyy') : txt_(x[14])) : '',
      lineas: rol !== 'VENDEDOR' ? [] : det.filter(l => mesNorm_(mesTxt_(l[0])) === cierre && norm_(l[2]) === quien).map(l => ({ dpto: txt_(l[3]), est: txt_(l[4]), proyecto: txt_(l[5]), cliente: txt_(l[6]), minuta: txt_(l[7]), escritura: txt_(l[8]), fin: txt_(l[9]), monto: Number(l[10]) || 0, pendiente: Number(l[11]) || 0, concepto: txt_(l[12]) })) });
  });
  const k = c => { const p = c.split(' '); return Number(p[1]) * 12 + MESES_.indexOf(p[0]); };
  return out.sort((a, b) => k(b.cierre) - k(a.cierre));
}
