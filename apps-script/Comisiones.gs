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
const COM_DET_HDR = ['Cierre', 'Jefe de ventas', 'Vendedor', 'Dpto', 'Est / Dep', 'Proyecto', 'Cliente', 'Fecha minuta', 'Fecha escritura', 'Tipo financ.', 'Monto comisión S/', 'Pendiente S/', 'Concepto', 'ID expediente', 'Origen'];
const COM_RES_HDR = ['Cierre', 'Rol', 'Persona', 'Jefe de ventas', 'Unidades', 'Tramo', 'Comisión S/', 'Bono S/', 'Adelanto S/', 'Dscto S/', 'T. Pagar S/', 'Pendiente S/', 'Publicado', 'Generado'];
const COM_FN = ['comEstado', 'comPrevisualizar', 'comGenerar', 'comPublicar', 'comExportar', 'misComisiones'];
try {
  COM_FN.forEach(f => { if (API_PUBLICA.indexOf(f) < 0) API_PUBLICA.push(f); });
  ['comEstado', 'comPrevisualizar', 'misComisiones'].forEach(f => { if (SOLO_LECTURA.indexOf(f) < 0) SOLO_LECTURA.push(f); });
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
  return { d: d, r: r };
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
function comCalcular_(cierre, h) {
  const esc = comEscala_();
  const previos = comFilas_(h.r).filter(r => mesNorm_(mesTxt_(r[0])) === cierre);
  // lo manual se conserva: bono / adelanto / dscto por persona y las líneas del detalle que no son AUTO (escrituras, ajustes)
  const manual = {}; previos.forEach(r => { manual[norm_(r[1]) + '|' + norm_(r[2])] = [r[7], r[8], r[9]]; });

  const bm = ss_().getSheetByName(HOJA_BM), v = bm.getRange(PRIMERA_FILA, 1, ULTIMA_FILA - PRIMERA_FILA + 1, 24).getValues();
  const filas = v.filter(r => r[22] && mesNorm_(mesTxt_(r[3])) === cierre);
  if (!filas.length) throw new Error('No hay ventas con cierre ' + cierre + ' en BASE MAESTRA.');
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
      det.push([cierre, p.jefe, p.nombre, txt_(r[6]), est, txt_(r[4]), txt_(r[7]), r[16] instanceof Date ? r[16] : '', '', txt_(r[19]), monto, pend, concepto, String(r[0]), 'AUTO']);
    });
    const m = manual['VENDEDOR|' + k] || ['', '', ''];   // comisión, pendiente y T. Pagar se ponen como fórmulas al escribir
    res.push([cierre, 'VENDEDOR', p.nombre, p.jefe, n, comTramoTxt_(n), null, m[0], m[1], m[2], null, null, 'NO', ahora]);
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
  return { det: det, res: res, depas: depas, porV: porV, sinMod: sinMod, previos: previos };
}
const comCierreOk_ = c => { c = mesNorm_(c); if (!/^[A-Z]+ \d{4}$/.test(c)) throw new Error('Elige un cierre válido (por ejemplo OCTUBRE 2026).'); return c; };

/** VISTA PREVIA: lo que saldría al generar, por vendedor (incluye las líneas que agregaste a mano). No escribe en el Sheet. */
function comPrevisualizar(token, cierre) {
  const u = sesion_(token); puede_(u, 'accesos');
  cierre = comCierreOk_(cierre);
  const h = comHojas_(), c = comCalcular_(cierre, h), n = x => Number(x) || 0;
  const lineas = c.det.concat(comFilas_(h.d).filter(r => mesNorm_(mesTxt_(r[0])) === cierre && norm_(r[14]) !== 'AUTO'));
  const linea = d => ({ dpto: txt_(d[3]), est: txt_(d[4]), proyecto: txt_(d[5]), cliente: txt_(d[6]), minuta: txt_(d[7]), escritura: txt_(d[8]), fin: txt_(d[9]),
    monto: n(d[10]), pendiente: n(d[11]), concepto: txt_(d[12]), id: txt_(d[13]), manual: norm_(d[14]) !== 'AUTO' });
  const personas = c.res.map(x => {
    const v = norm_(x[1]) === 'VENDEDOR', l = v ? lineas.filter(d => norm_(d[2]) === norm_(x[2])).map(linea) : [];
    const com = v ? l.reduce((a, d) => a + d.monto, 0) : n(x[6]), pen = v ? l.reduce((a, d) => a + d.pendiente, 0) : n(x[11]);
    return { rol: x[1], nombre: x[2], jefe: x[3], unidades: x[4], tramo: txt_(x[5]), comision: com, bono: n(x[7]), adelanto: n(x[8]), dscto: n(x[9]),
      pendiente: pen, total: com + n(x[7]) - n(x[8]) - n(x[9]), lineas: l };
  });
  return { cierre: cierre, departamentos: c.depas.length, vendedores: Object.keys(c.porV).length, revisar: c.sinMod, yaGenerado: c.previos.length > 0,
    publicado: c.previos.some(r => norm_(r[12]) === 'SI'), total: personas.reduce((a, p) => a + p.total, 0), personas: personas };
}

function comGenerar(token, cierre, forzar) {
  const u = sesion_(token); puede_(u, 'accesos');
  cierre = comCierreOk_(cierre);
  const h = comHojas_(), c = comCalcular_(cierre, h), det = c.det, res = c.res, depas = c.depas, porV = c.porV, sinMod = c.sinMod;
  if (c.previos.some(r => norm_(r[12]) === 'SI') && !forzar) throw new Error('El cierre ' + cierre + ' ya está PUBLICADO. Despublícalo primero si quieres regenerarlo.');

  // escribe: borra lo AUTO del cierre (lo manual del detalle se queda) y el resumen del cierre
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    comBorrar_(h.d, r => mesNorm_(mesTxt_(r[0])) === cierre && norm_(r[14]) === 'AUTO');
    comBorrar_(h.r, r => mesNorm_(mesTxt_(r[0])) === cierre);
    if (det.length) { const i = h.d.getLastRow() + 1; h.d.getRange(i, 1, det.length, 1).setNumberFormat('@'); h.d.getRange(i, 1, det.length, COM_DET_HDR.length).setValues(det.map(fila_)); }
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
const COM_TELEFONO = '472-7813';
const COM_COL = { naranja: '#F8CBAD', morado: '#78206E', amarillo: '#FFC000', verde: '#92D050', aviso: '#FFFF00', rojo: '#FF0000' };
const COM_SOL = '"S/." #,##0.00;[Red]"S/." -#,##0.00';

function comBloqueVendedor_(sh, f, p, lineas, cierre) {
  const C = COM_COL, primero = comPrimerDia_(cierre);
  // 1) nombre · mes · teléfono
  sh.getRange(f, 2, 1, 10).setBackground(C.naranja);
  sh.getRange(f, 3, 1, 3).merge().setValue(String(p[2]).toUpperCase()).setBackground(C.morado).setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(f, 6).setValue(primero).setNumberFormat('mmm-yy').setBackground(C.morado).setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(f, 9).setValue(COM_TELEFONO).setFontWeight('bold').setHorizontalAlignment('center');
  f++;
  // 2) cabecera en dos filas
  sh.getRange(f, 2, 2, 10).setValues([
    ['DPTO', 'EST/DEP', 'PROYECTO', 'CLIENTE', 'FECHA DE', 'FECHA DE', 'TIPO', 'MONTO', 'PENDIENTE', 'CONCEPTO'],
    ['', '', '', '', 'MINUTA', 'ESCRITURA', 'FINANC.', 'COMISION S/.', '', '']])
    .setBackground(C.naranja).setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, false);
  [2, 3, 4, 5, 10, 11].forEach(c => sh.getRange(f, c, 2, 1).merge());
  f += 2;
  // 3) líneas
  const ini = f;
  if (lineas.length) {
    sh.getRange(f, 2, lineas.length, 10).setValues(lineas.map(d => [txt_(d[3]), txt_(d[4]), txt_(d[5]), txt_(d[6]), d[7] instanceof Date ? d[7] : txt_(d[7]),
      d[8] instanceof Date ? d[8] : txt_(d[8]), txt_(d[9]), Number(d[10]) || 0, Number(d[11]) || 0, txt_(d[12])]))
      .setFontSize(9).setBorder(true, true, true, true, true, true).setVerticalAlignment('middle');
    sh.getRange(f, 2, lineas.length, 3).setHorizontalAlignment('center');
    sh.getRange(f, 6, lineas.length, 2).setNumberFormat('dd/mm/yyyy').setHorizontalAlignment('center');
    sh.getRange(f, 8, lineas.length, 1).setHorizontalAlignment('center');
    sh.getRange(f, 9, lineas.length, 2).setNumberFormat(COM_SOL);
    f += lineas.length;
  }
  // 4) total de las líneas
  const tot = f;
  sh.getRange(f, 2).setValue('TOTAL').setFontWeight('bold');
  sh.getRange(f, 9).setFormula(lineas.length ? '=SUM(I' + ini + ':I' + (f - 1) + ')' : '=0').setBackground(C.amarillo);
  sh.getRange(f, 10).setFormula(lineas.length ? '=SUM(J' + ini + ':J' + (f - 1) + ')' : '=0');
  sh.getRange(f, 9, 1, 2).setNumberFormat(COM_SOL).setFontWeight('bold').setBorder(true, true, true, true, true, false);
  f += 2;
  // 5) total por pagar
  sh.getRange(f, 2).setValue('TOTAL POR PAGAR').setFontWeight('bold'); f++;
  sh.getRange(f, 5, 1, 5).setValues([['COMISION', 'BONO(*)', 'ADELANTO', 'DSCTO', 'T.PAGAR']]).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(f, 6).setFontColor(C.rojo); sh.getRange(f, 9).setBackground(C.verde); f++;
  sh.getRange(f, 4).setValue(String(p[2]).toUpperCase()).setFontWeight('bold');
  sh.getRange(f, 5).setFormula('=I' + tot);
  sh.getRange(f, 6, 1, 3).setValues([[Number(p[7]) || 0, Number(p[8]) || 0, Number(p[9]) || 0]]);
  sh.getRange(f, 9).setFormula('=E' + f + '+F' + f + '-G' + f + '-H' + f).setBackground(C.verde).setFontWeight('bold');
  sh.getRange(f, 5, 1, 5).setNumberFormat(COM_SOL);
  sh.getRange(f, 10).setValue(norm_(p[12]) === 'SI' ? 'PUBLICADO' : 'PROGRAMADO').setFontColor(C.rojo).setBackground(C.aviso).setFontWeight('bold').setHorizontalAlignment('center');
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
  sh.getRange(f, 11).setValue(norm_(jr[12]) === 'SI' ? 'PUBLICADO' : 'PROGRAMADO').setFontColor(C.rojo).setBackground(C.aviso).setFontWeight('bold').setHorizontalAlignment('center');
  return f + 3;
}
const comPrimerDia_ = c => { const p = c.split(' '); return new Date(Number(p[1]), MESES_.indexOf(p[0]), 1); };
function comAnchos_(sh) {
  [[1, 20], [2, 70], [3, 60], [4, 110], [5, 290], [6, 85], [7, 85], [8, 100], [9, 110], [10, 100], [11, 210]].forEach(x => sh.setColumnWidth(x[0], x[1]));
  sh.setHiddenGridlines(true); sh.getRange('A:K').setFontFamily('Arial');
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
    sh.getRange(2, 2, 1, 10).setBackground(COM_COL.naranja);
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

/* ---------- lo que ve cada persona (solo lo publicado y solo lo suyo) ---------- */
function misComisiones(token) {
  const u = sesion_(token), yo = [norm_(u.ref), norm_(u.nombre)].filter(Boolean);
  const ss = ss_(), r = ss.getSheetByName(HOJA_COM_RES), d = ss.getSheetByName(HOJA_COM_DET); if (!r) return [];
  const det = d ? comFilas_(d) : [], out = [];
  r.getDataRange().getValues().slice(1).forEach(x => {
    if (norm_(x[12]) !== 'SI') return;
    const rol = norm_(x[1]), quien = norm_(x[2]);
    const mio = u.rol === 'ADMIN' ? false : (rol === 'VENDEDOR' && yo.indexOf(quien) >= 0) || (rol === 'JEFE' && u.rol === 'JEFE' && yo.indexOf(quien) >= 0) || (rol === 'GERENTE' && u.rol === 'GERENTE');
    if (!mio) return;
    const cierre = mesNorm_(mesTxt_(x[0]));
    out.push({ cierre: cierre, rol: rol, unidades: x[4], tramo: String(x[5]), comision: Number(x[6]) || 0, bono: Number(x[7]) || 0, adelanto: Number(x[8]) || 0, dscto: Number(x[9]) || 0, total: Number(x[10]) || 0, pendiente: Number(x[11]) || 0,
      lineas: rol !== 'VENDEDOR' ? [] : det.filter(l => mesNorm_(mesTxt_(l[0])) === cierre && norm_(l[2]) === quien).map(l => ({ dpto: txt_(l[3]), est: txt_(l[4]), proyecto: txt_(l[5]), cliente: txt_(l[6]), minuta: txt_(l[7]), escritura: txt_(l[8]), fin: txt_(l[9]), monto: Number(l[10]) || 0, pendiente: Number(l[11]) || 0, concepto: txt_(l[12]) })) });
  });
  const k = c => { const p = c.split(' '); return Number(p[1]) * 12 + MESES_.indexOf(p[0]); };
  return out.sort((a, b) => k(b.cierre) - k(a.cierre));
}
