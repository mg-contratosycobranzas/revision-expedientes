/**
 * SEGUIMIENTO DEL EXPEDIENTE — se calcula con el historial que el portal ya guarda (hoja SUBSANACIONES), así funciona
 * también para expedientes revisados antes de esta mejora. Archivo independiente (lo usa Code.gs si existe).
 * Por área (CIERRE = Contractual · SBS):
 *   - Revisado por: quien hizo la PRIMERA revisión del área (y quién hizo la última)
 *   - N° de revisión: 1 en la primera revisión; +1 cada vez que Comercial subsana esa área y se vuelve a revisar
 * Del expediente: fecha de registro (BASE MAESTRA), N° de subsanaciones, días transcurridos y días hasta quedar completo.
 */
const SEG_AREAS = ['CIERRE', 'SBS'];
const segFecha_ = d => d instanceof Date ? Utilities.formatDate(d, 'America/Lima', 'dd/MM/yyyy HH:mm') : '';
const segDias_ = (a, b) => (a instanceof Date && b instanceof Date) ? Math.max(0, Math.round((b - a) / 864e5 * 10) / 10) : null;

/** Historial ordenado de un expediente en SUBSANACIONES (filas crudas). */
function segEventos_(id, filas) {
  return filas.filter(r => String(r[1]).trim() === id && r[0] instanceof Date).sort((a, b) => a[0] - b[0]);
}
/** Rondas, revisores y fechas por área a partir del historial. */
function segAreas_(ev) {
  const out = {};
  SEG_AREAS.forEach(A => out[A] = { rondas: 0, revisor: '', primera: null, ultima: null, ultimoRevisor: '', pendiente: false });
  ev.forEach(r => {
    const quien = String(r[4] || ''), cc = /Contratos y Cobranzas/i.test(quien), nombre = quien.replace(/\s*\(.*\)\s*$/, '').trim();
    const ar = String(r[5] || '').toUpperCase(), esRevision = cc && /^REVISI/i.test(String(r[2]).trim());
    SEG_AREAS.forEach(A => {
      if (ar && ar.indexOf(A) < 0) return;   // área vacía = aplica a ambas
      const s = out[A];
      if (esRevision) {
        if (!s.rondas) { s.rondas = 1; s.revisor = nombre; s.primera = r[0]; }
        else if (s.pendiente) { s.rondas++; s.pendiente = false; }
        s.ultima = r[0]; s.ultimoRevisor = nombre;
      } else if (!cc && s.rondas) s.pendiente = true;   // Comercial subsanó después de una revisión
    });
  });
  return out;
}
/** Si un área se revisó directo en el Sheet (sin historial en el portal): SBS → la persona de CONTACTOS con área SBS
 *  (nombre completo según ACCESO_WEB); CIERRE → "Revisado por" de CONTROL DE REVISIONES. */
function segRevisorRespaldo_(area, revisorCierre) {
  if (area === 'CIERRE') return String(revisorCierre || '').trim();
  try {
    const c = (typeof contactos_ === 'function' ? contactos_() : []).find(x => x.area === area); if (!c) return '';
    const ac = ss_().getSheetByName(HOJA_ACCESO), k = norm_(c.nombre).split(' ')[0];
    const u = ac.getRange(2, 1, Math.max(ac.getLastRow() - 1, 1), 1).getValues().map(r => String(r[0]).trim()).find(n => norm_(n).split(' ')[0] === k);
    return u || c.nombre;
  } catch (e) { return ''; }
}
/** Seguimiento completo de un expediente (lo que ve el portal).
 *  estados = {CIERRE: estado contractual, SBS: estado SBS, revisorCierre} (de CONTROL DE REVISIONES) para el respaldo. */
function seguimiento_(id, estadoActual, estados) {
  estados = estados || {};
  id = String(id || '').trim();
  const ss = ss_(), su = ss.getSheetByName(HOJA_SUB);
  const filas = su && su.getLastRow() >= 2 ? su.getRange(2, 1, su.getLastRow() - 1, 8).getValues() : [];
  const ev = segEventos_(id, filas), ar = segAreas_(ev);
  // registro: FECHA INGRESO de BASE MAESTRA (col C); si no hay, el primer evento
  let registro = null;
  try {
    const bm = ss.getSheetByName(HOJA_BM), v = bm.getRange(PRIMERA_FILA, 1, ULTIMA_FILA - PRIMERA_FILA + 1, 3).getValues();
    const f = v.find(r => String(r[0]).trim() === id); if (f && f[2] instanceof Date) registro = f[2];
  } catch (e) {}
  if (!registro && ev.length) registro = ev[0][0];
  const subs = ev.filter(r => !/Contratos y Cobranzas/i.test(String(r[4] || ''))).length;
  const ultRev = SEG_AREAS.map(A => ar[A].ultima).filter(Boolean).sort((a, b) => b - a)[0] || null;
  const completo = norm_(estadoActual) === 'COMPLETO', hoy = new Date();
  const areas = {};
  SEG_AREAS.forEach(A => {   // área revisada en el Sheet sin pasar por el portal: revisor de respaldo y al menos 1 revisión
    const s = ar[A]; if (!s.rondas && String(estados[A] || '').trim()) { s.rondas = 1; s.revisor = s.revisor || segRevisorRespaldo_(A, estados.revisorCierre); s.ultimoRevisor = s.revisor; }
  });
  SEG_AREAS.forEach(A => { const s = ar[A]; areas[A] = { rondas: s.rondas, revisor: s.revisor, ultimoRevisor: s.ultimoRevisor, primera: segFecha_(s.primera), ultima: segFecha_(s.ultima), pendiente: s.pendiente,
    diasHastaPrimera: segDias_(registro, s.primera) }; });
  return { registro: segFecha_(registro), subsanaciones: subs, areas: areas, completo: completo,
    dias: completo ? segDias_(registro, ultRev || hoy) : segDias_(registro, hoy), finalizado: completo && ultRev ? segFecha_(ultRev) : '' };
}
/** N° de revisión que corresponde ahora a un área (para guardarlo en el Sheet al revisar). */
function segRondaNueva_(id, area, actual) {
  const su = ss_().getSheetByName(HOJA_SUB), filas = su && su.getLastRow() >= 2 ? su.getRange(2, 1, su.getLastRow() - 1, 8).getValues() : [];
  const s = segAreas_(segEventos_(String(id).trim(), filas))[area] || { rondas: 0, pendiente: false };
  const base = Math.max(Number(actual) || 0, s.rondas);
  return base === 0 ? 1 : (s.pendiente ? base + 1 : base);
}
