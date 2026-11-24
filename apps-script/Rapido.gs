/**
 * ⚡ ABRIR EXPEDIENTES MÁS RÁPIDO (archivo independiente; se conecta solo con Code.gs).
 * 1) Listas desplegables: la ficha CIERRE/SBS leía del Sheet, columna por columna, las opciones de cada lista
 *    (OK / OBSERVADO / NO APLICA…). Eran decenas de lecturas en cada expediente. Ahora cada lista se lee una vez y
 *    se recuerda 10 minutos (si cambias una lista en el Sheet, el portal la toma como máximo 10 minutos después).
 * 2) El detalle y la ficha de cada expediente se guardan 10 minutos en caché. La caché se renueva sola en cuanto
 *    alguien guarda algo desde el portal (verDatos_ cambia), así que nunca muestra un dato viejo después de una
 *    revisión o subsanación hecha en el portal. Si se edita el Sheet a mano, el cambio aparece en máximo 10 minutos.
 *    Cada usuario tiene su propia caché (lo que ve depende de sus permisos).
 * 3) perfilAbrir (solo ADMIN): mide cuánto tarda cada paso de abrir un expediente, para saber qué sigue lento.
 */
try {
  ['detalleRapido', 'fichaRapida', 'perfilAbrir'].forEach(f => {
    if (API_PUBLICA.indexOf(f) < 0) API_PUBLICA.push(f);
    if (SOLO_LECTURA.indexOf(f) < 0) SOLO_LECTURA.push(f);
  });
} catch (e) {}

const RAP_SEG = 600;

/* ---------- 1) opciones de las listas desplegables con memoria ----------
 * Reemplaza a opcionesDv_ de Code.gs (este archivo carga después, así que esta versión es la que se usa).
 * Hace exactamente lo mismo, pero la lista que viene de un rango del Sheet se lee una sola vez. */
const RAP_DV = {};
function opcionesDv_(dv, respaldo) {
  try {
    if (!dv) return respaldo;
    const t = dv.getCriteriaType(), v = dv.getCriteriaValues(); let l = [];
    if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) l = v[0];
    else if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) l = rapRango_(v[0]);
    l = (l || []).map(x => String(x).trim()).filter(x => x && x !== 'XXX');
    return l.length ? l : respaldo;
  } catch (e) { return respaldo; }
}
function rapRango_(rg) {
  let k = '';
  try { k = 'rapdv_' + rg.getSheet().getSheetId() + '_' + rg.getA1Notation(); } catch (e) {}
  if (k && RAP_DV[k]) return RAP_DV[k];
  const c = CacheService.getScriptCache();
  if (k) { const z = c.get(k); if (z) try { return (RAP_DV[k] = JSON.parse(z)); } catch (e) {} }
  const l = rg.getValues().map(r => r[0]).filter(x => x !== '' && x != null).map(String);
  if (k) { RAP_DV[k] = l; try { const s = JSON.stringify(l); if (s.length < 90000) c.put(k, s, RAP_SEG); } catch (e) {} }
  return l;
}

/* ---------- 2) caché del detalle y la ficha ---------- */
function rapClave_(tipo, token, id) {
  const u = sesion_(token);
  const per = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, (u.permisos || []).join(','))).slice(0, 12);
  return 'rap2_' + tipo + '_' + (typeof verDatos_ === 'function' ? verDatos_() : '') + '_' + u.fila + '_' + per + '_' + String(id || '').trim();
}
function rapCache_(clave, calcular) {
  const c = CacheService.getScriptCache(), z = c.get(clave);
  if (z) try { return JSON.parse(z); } catch (e) {}
  const v = calcular();
  try { const s = JSON.stringify(v); if (s.length < 95000) c.put(clave, s, RAP_SEG); } catch (e) {}
  return v;
}
/** El portal pide "detalle", que el motor atiende con detalleCompleto_ (incluye todas las columnas y el seguimiento). */
const rapDetalle_ = (token, id) => typeof detalleCompleto_ === 'function' ? detalleCompleto_(token, id) : detalle(token, id);
/** Igual que el detalle del portal, con caché. */
function detalleRapido(token, id) { return rapCache_(rapClave_('det', token, id), () => rapDetalle_(token, id)); }
/** Igual que ficha(), con caché. */
function fichaRapida(token, id) { return rapCache_(rapClave_('fic', token, id), () => ficha(token, id)); }

/* ---------- 3) cronómetro de cada paso (solo ADMIN, no usa la caché) ---------- */
function perfilAbrir(token, id) {
  const u = sesion_(token); if (u.rol !== 'ADMIN') throw new Error('Solo la administración.');
  id = String(id || '').trim();
  const P = [], paso = (t, f) => { const a = Date.now(); let err = ''; try { f(); } catch (e) { err = String(e.message || e); } P.push({ paso: t, ms: Date.now() - a, error: err }); };
  paso('Revisar la sesión', () => sesion_(token));
  paso('Lista de expedientes (con caché)', () => misExpedientes(token));
  paso('Lista de expedientes (sin caché)', () => misExpedientesSinCache_(token));
  paso('Detalle básico (CIERRE, SBS, subsanaciones)', () => detalle(token, id));
  paso('Ficha CIERRE (1ª vez)', () => leerFicha_('CIERRE', id));
  paso('Ficha SBS (1ª vez)', () => leerFicha_('SBS', id));
  paso('Ficha CIERRE (2ª vez, listas en memoria)', () => leerFicha_('CIERRE', id));
  paso('Seguimiento', () => { if (typeof seguimiento_ === 'function') seguimiento_(id, '', {}); });
  paso('Detalle completo', () => rapDetalle_(token, id));
  paso('Ficha completa', () => ficha(token, id));
  return { id: id, pasos: P, total: P.reduce((a, p) => a + p.ms, 0) };
}
