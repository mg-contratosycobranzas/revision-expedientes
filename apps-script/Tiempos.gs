/**
 * TIEMPOS DE SUBSANACIÓN — archivo independiente (se conecta solo con Code.gs). Solo lee: no escribe en ninguna hoja.
 * Mide con el historial de SUBSANACIONES, expediente por expediente:
 *   - Ronda: Contratos devuelve el expediente y Comercial responde. Varios guardados seguidos de Contratos sin respuesta cuentan como una sola ronda.
 *   - Respuesta: desde la primera devolución de la ronda hasta el primer envío de Comercial.
 *   - Revisión: desde ese envío hasta la siguiente respuesta de Contratos.
 *   - Reproceso: cada ronda después de la primera (la meta es máximo 1 ronda).
 * Las horas son corridas (no descuentan noches ni fines de semana).
 */
try {
  if (API_PUBLICA.indexOf('tiemposSubsanacion') < 0) API_PUBLICA.push('tiemposSubsanacion');
  if (SOLO_LECTURA.indexOf('tiemposSubsanacion') < 0) SOLO_LECTURA.push('tiemposSubsanacion');
} catch (e) {}

const tieH_ = (a, b) => Math.round((b - a) / 36e5 * 10) / 10;
const tieF_ = d => Utilities.formatDate(d, 'America/Lima', 'dd/MM/yyyy HH:mm');

/** Saca los puntos observados de "🔴 Observado: CIERRE · X, SBS · Y." o "🔴 Sigue observado: X (comentario), Y." */
function tiePuntos_(texto) {
  const m = String(texto || '').match(/(?:Sigue observado|Observado):\s*([\s\S]*)/i);
  if (!m) return [];
  const b = m[1].split(/\.(?:\s|$)/)[0];
  const partes = /·/.test(b) ? b.split(/,\s*(?=(?:CIERRE|SBS)\s*·)/i) : b.split(/,\s*/);
  return partes.map(p => p.replace(/^\s*(?:CIERRE|SBS)\s*·\s*/i, '').replace(/\s*\([^)]*\)\s*$/, '').trim().toUpperCase().slice(0, 120)).filter(Boolean);
}

/** ev: filas de SUBSANACIONES (A:I) · ct: filas de CONTROL DE REVISIONES (A:Q, desde la fila 5). */
function tiempoCalc_(ev, ct, ahora) {
  const info = {}, sinEv = {};
  ct.forEach(r => { const id = String(r[0]).trim(); if (/^EXP/i.test(id)) { info[id] = { cliente: String(r[1] || ''), proyecto: String(r[2] || ''), unidad: String(r[3] || ''), vendedor: String(r[4] || ''), jefe: String(r[5] || ''), estado: String(r[11] || '').trim(), mes: String(r[16] || '').trim() }; sinEv[id] = true; } });
  const porExp = {};
  ev.forEach(r => {
    const id = String(r[1]).trim();
    if (!(r[0] instanceof Date) || !/^EXP/i.test(id)) return;
    (porExp[id] = porExp[id] || []).push({ d: r[0], texto: String(r[3] || ''), cc: /Contratos y Cobranzas/i.test(String(r[4] || '')), est: String(r[7] || ''), plazo: String(r[8] || '') });
  });
  let primera = null, huerfanas = 0;
  const filas = Object.keys(porExp).map(id => {
    delete sinEv[id];
    const lista = porExp[id].sort((a, b) => a.d - b.d), rondas = [];
    let cur = null;
    lista.forEach(e => {
      if (!primera || e.d < primera) primera = e.d;
      if (e.cc) {
        if (cur && cur.fr && !cur.frev) cur.frev = e.d;
        if (/Devuelto/i.test(e.est) || /Sigue observado/i.test(e.texto)) {
          if (!cur || cur.fr) { cur = { ini: e.d, fr: null, frev: null, fuera: false, puntos: [] }; rondas.push(cur); }
          const p = tiePuntos_(e.texto); if (p.length) cur.puntos = p;
        }
      } else if (!cur) huerfanas++;
      else if (!cur.frev) {
        if (!cur.fr) cur.fr = e.d;
        if (/Fuera de plazo/i.test(e.plazo)) cur.fuera = true;
      }
    });
    if (!rondas.length) return null;
    const ult = rondas[rondas.length - 1], base = info[id] || { cliente: '', proyecto: '', unidad: '', vendedor: '', jefe: '', estado: '', mes: '' };
    const espera = !ult.fr ? 'Comercial' : (!ult.frev ? 'Contratos' : '');
    return Object.assign({ id: id }, base, {
      espera: espera, esperaH: espera === 'Comercial' ? tieH_(ult.ini, ahora) : (espera === 'Contratos' ? tieH_(ult.fr, ahora) : null),
      rondas: rondas.map((r, k) => ({ ini: tieF_(r.ini), resp: r.fr ? tieH_(r.ini, r.fr) : null, rev: r.fr && r.frev ? tieH_(r.fr, r.frev) : null, fuera: r.fuera, puntos: r.puntos,
        repite: k ? r.puntos.filter(p => rondas[k - 1].puntos.indexOf(p) >= 0) : [] }))
    });
  }).filter(Boolean);
  return { generado: tieF_(ahora), desde: primera ? tieF_(primera) : '', filas: filas, huerfanas: huerfanas,
    sinHistorial: Object.keys(sinEv).filter(id => /OBSERVADO/i.test(info[id].estado)).length };
}

/** Administración: datos del reporte (el portal calcula los totales según el mes que se elija). */
function tiemposSubsanacion(token) {
  const u = sesion_(token); puede_(u, 'accesos');
  const sub = ss_().getSheetByName(HOJA_SUB), ctr = ss_().getSheetByName('CONTROL DE REVISIONES');
  const nS = sub ? sub.getLastRow() : 0, nC = ctr ? ctr.getLastRow() : 0;
  return tiempoCalc_(nS > 1 ? sub.getRange(2, 1, nS - 1, 9).getValues() : [], nC > 4 ? ctr.getRange(5, 1, nC - 4, 17).getValues() : [], new Date());
}
