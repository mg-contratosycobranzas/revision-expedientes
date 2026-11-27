/**
 * 📋 REGISTRO DEL CLIENTE — al tocar "① Registro" en la ficha del expediente se ve su fila de BASE MAESTRA como resumen
 * (con los datos vacíos marcados) y, solo para quien tenga permiso, se puede corregir dato por dato.
 *
 * Permisos (se dan por persona en 👥 Accesos; ADMIN siempre los tiene):
 *   - registro_ver    → ver el resumen del registro
 *   - registro_editar → corregir datos del registro (manual o con la ficha de cliente en PDF)
 * Columnas automáticas (ID EXPEDIENTE, MES, JEFE COMERCIAL) nunca se escriben.
 * Cada corrección queda en LOG_WEB con el valor anterior y el nuevo.
 * Archivo independiente: se conecta con Code.gs agregando sus funciones a API_PUBLICA y sus permisos a PERMISOS.
 */
const REG_COLS = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W'];
const REG_FECHAS = { C: 1, P: 1, Q: 1 }, REG_MONTO = { R: 1 }, REG_DOC = { I: 1, M: 1 };
const REG_FILA_TIT = 4;   // fila de títulos de BASE MAESTRA

try {
  ['registroVer', 'registroGuardar'].forEach(f => { if (API_PUBLICA.indexOf(f) < 0) API_PUBLICA.push(f); });
} catch (e) {}
try { ['registro_ver', 'registro_editar'].forEach(p => { if (PERMISOS.indexOf(p) < 0) PERMISOS.push(p); }); } catch (e) {}

const regCol_ = l => l.charCodeAt(0) - 64;   // 'C' → 3
function regPuede_(u, p) { return u.rol === 'ADMIN' || (u.permisos || []).indexOf(p) >= 0; }
function regFila_(token, id) {
  id = String(id || '').trim();
  const ex = misExpedientes(token).find(x => x.id === id);
  if (!ex) throw new Error('No tienes acceso al expediente ' + id + '.');
  const bm = ss_().getSheetByName(HOJA_BM), ids = bm.getRange(PRIMERA_FILA, 1, ULTIMA_FILA - PRIMERA_FILA + 1, 1).getDisplayValues();
  const i = ids.findIndex(r => String(r[0]).trim() === id);
  if (i < 0) throw new Error('No encontré ' + id + ' en BASE MAESTRA.');
  return { bm: bm, fila: PRIMERA_FILA + i, ex: ex };
}
/** Opciones de cada columna según la validación de datos de esa fila (se guardan 10 min). */
function regOpciones_(bm, fila) {
  const ck = 'regOpc_' + fila, c = CacheService.getScriptCache(), g = c.get(ck);
  if (g) try { return JSON.parse(g); } catch (e) {}
  const reglas = bm.getRange(fila, 3, 1, REG_COLS.length).getDataValidations()[0], out = {};
  reglas.forEach((r, k) => {
    if (!r) return;
    const t = r.getCriteriaType(), v = r.getCriteriaValues();
    let op = null;
    if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) op = (v[0] || []).map(String);
    else if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE && v[0]) op = v[0].getDisplayValues().map(z => String(z[0]).trim());
    if (op) out[REG_COLS[k]] = op.filter((x, i, a) => x && a.indexOf(x) === i);
  });
  try { c.put(ck, JSON.stringify(out), 600); } catch (e) {}
  return out;
}
const regISO_ = d => d instanceof Date ? Utilities.formatDate(d, 'America/Lima', 'yyyy-MM-dd') : '';
/** ¿Falta este dato? (los que dependen de otro: separación de bienes solo si es casado; banco y funcionario no si es contado;
 *  copropietario: si hay uno de los dos datos, falta el otro). */
function regFalta_(c, val, fila) {
  const v = k => String(fila[k] || '').trim(), vacio = !String(val || '').trim();
  if (!vacio) return false;
  if (c === 'L') return /CASAD/i.test(v('K'));
  if (c === 'U' || c === 'V') return !/CONTADO/i.test(v('T'));
  if (c === 'M') return !!v('N');
  if (c === 'N') return !!v('M');
  if (c === 'Q') return false;   // la firma de la MCV puede ser posterior
  return true;
}

/** Resumen del registro de un expediente. */
function registroVer(token, id) {
  const u = sesion_(token);
  if (!regPuede_(u, 'registro_ver') && !regPuede_(u, 'registro_editar')) throw new Error('No tienes permiso para ver el registro del cliente.');
  const r = regFila_(token, id), rg = r.bm.getRange(r.fila, 1, 1, 24);
  const tit = r.bm.getRange(REG_FILA_TIT, 1, 1, 24).getDisplayValues()[0], val = rg.getValues()[0], dis = rg.getDisplayValues()[0];
  const opc = regOpciones_(r.bm, r.fila), F = {};
  REG_COLS.forEach(c => F[c] = dis[regCol_(c) - 1]);
  const campos = REG_COLS.map(c => {
    const i = regCol_(c) - 1, v = val[i];
    return { c: c, t: String(tit[i] || c).trim(), d: String(dis[i] || '').trim(),
      v: REG_FECHAS[c] ? regISO_(v) : (REG_MONTO[c] ? (v === '' ? '' : String(v)) : String(dis[i] || '').trim()),
      tipo: REG_FECHAS[c] ? 'fecha' : REG_MONTO[c] ? 'monto' : opc[c] ? 'lista' : 'texto', op: opc[c] || [], falta: regFalta_(c, dis[i], F) };
  });
  const auto = [['A', 0], ['B', 1], ['X', 23]].map(([c, i]) => ({ c: c, t: String(tit[i] || c).trim(), d: String(dis[i] || '').trim() }));
  return { id: r.ex.id, fila: r.fila, campos: campos, auto: auto, faltan: campos.filter(x => x.falta).length, puedeEditar: regPuede_(u, 'registro_editar') };
}

/** Corrige datos del registro. cambios = {C:'2026-11-26', T:'CHIP AHORRO', ...}; solo se escriben las columnas que cambian. */
function registroGuardar(token, id, cambios) {
  const u = sesion_(token);
  if (!regPuede_(u, 'registro_editar')) throw new Error('No tienes permiso para corregir el registro del cliente.');
  cambios = cambios || {};
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const r = regFila_(token, id), opc = regOpciones_(r.bm, r.fila), log = [];
    Object.keys(cambios).forEach(c => {
      if (REG_COLS.indexOf(c) < 0) throw new Error('La columna ' + c + ' no se puede corregir desde el portal.');
      const cel = r.bm.getRange(r.fila, regCol_(c)), antes = String(cel.getDisplayValue()).trim();
      let nv = String(cambios[c] == null ? '' : cambios[c]).trim(), escribir = nv;
      if (nv && REG_FECHAS[c]) {
        const m = nv.match(/^(\d{4})-(\d{2})-(\d{2})$/); if (!m) throw new Error('Fecha no válida en ' + c + ': ' + nv);
        escribir = new Date(+m[1], +m[2] - 1, +m[3]);
        if (regISO_(cel.getValue()) === nv) return;
      } else if (nv && REG_MONTO[c]) {
        const n = Number(nv.replace(/[^\d.\-]/g, '')); if (isNaN(n)) throw new Error('Monto no válido: ' + nv);
        escribir = n; if (Number(cel.getValue()) === n) return;
      } else if (nv && REG_DOC[c]) {
        nv = nv.replace(/\D/g, ''); if ([8, 9, 11].indexOf(nv.length) < 0) throw new Error('El DNI / CE / RUC debe tener 8, 9 u 11 dígitos.');
        escribir = /^0/.test(nv) ? "'" + nv : nv; if (antes.replace(/\D/g, '') === nv) return;
      } else {
        if (nv && opc[c] && opc[c].indexOf(nv) < 0) throw new Error('"' + nv + '" no está en la lista de ' + c + '.');
        if (antes === nv) return;
      }
      cel.setValue(escribir);
      log.push(c + ': ' + (antes || '(vacío)') + ' → ' + (nv || '(vacío)'));
    });
    if (log.length) {
      log_(u.nombre, 'REGISTRO_CORREGIDO', r.ex.id + ' · ' + log.join(' · '));
      try { tocarDatos_(); } catch (e) {}
    }
    return { cambios: log.length, detalle: log, registro: registroVer(token, id) };
  } finally { lock.releaseLock(); }
}
