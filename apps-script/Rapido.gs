/**
 * ⚡ ABRIR EXPEDIENTES MÁS RÁPIDO (archivo independiente; se conecta solo con Code.gs).
 * El detalle y la ficha CIERRE/SBS de cada expediente se guardan 2 minutos en caché.
 * La caché se renueva sola en cuanto alguien guarda algo desde el portal (verDatos_ cambia), así que nunca muestra
 * un dato viejo después de una revisión o una subsanación hecha en el portal. Si se edita el Sheet a mano, el cambio
 * aparece como máximo 2 minutos después.
 * Cada usuario tiene su propia caché (lo que ve depende de sus permisos).
 */
try {
  ['detalleRapido', 'fichaRapida'].forEach(f => {
    if (API_PUBLICA.indexOf(f) < 0) API_PUBLICA.push(f);
    if (SOLO_LECTURA.indexOf(f) < 0) SOLO_LECTURA.push(f);
  });
} catch (e) {}

function rapClave_(tipo, token, id) {
  const u = sesion_(token);
  const per = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, (u.permisos || []).join(','))).slice(0, 12);
  return 'rap_' + tipo + '_' + (typeof verDatos_ === 'function' ? verDatos_() : '') + '_' + u.fila + '_' + per + '_' + String(id || '').trim();
}
function rapCache_(clave, calcular) {
  const c = CacheService.getScriptCache(), z = c.get(clave);
  if (z) try { return JSON.parse(z); } catch (e) {}
  const v = calcular();
  try { const s = JSON.stringify(v); if (s.length < 95000) c.put(clave, s, 120); } catch (e) {}
  return v;
}
/** Igual que detalle(), con caché de 2 minutos. */
function detalleRapido(token, id) { return rapCache_(rapClave_('det', token, id), () => detalle(token, id)); }
/** Igual que ficha(), con caché de 2 minutos. */
function fichaRapida(token, id) { return rapCache_(rapClave_('fic', token, id), () => ficha(token, id)); }
