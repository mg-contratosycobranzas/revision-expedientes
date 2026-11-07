/**
 * BLOQUEO DE SUBSANACIÓN — archivo independiente (se conecta solo con Code.gs).
 * Comercial (quien subsana y NO revisa) queda bloqueado:
 *   - ESPERA: ya subsanó y espera la respuesta de Contratos (el expediente está 🟠 SUBSANADO)
 *   - PLAZO : venció "Comercial presenta subsanados" (CALENDARIO) y sigue observado
 * Puede "Pedir habilitación" con un motivo. Contratos / administración decide (hoja SUBSANACION_HABILITA):
 *   HABILITADO                → puede subsanar una vez más (dentro de fecha; no salta el plazo vencido)
 *   HABILITADO FUERA DE FECHA → excepción: puede subsanar una vez aunque el plazo haya vencido
 *   RECHAZADO                 → sigue bloqueado
 * Al usarla, la habilitación queda USADA y el expediente vuelve a bloquearse hasta la siguiente revisión.
 */
const HOJA_HAB = 'SUBSANACION_HABILITA';
const HAB_HDR = ['ID expediente', 'Estado', 'Pidió', 'Fecha pedido', 'Motivo', 'Decidió', 'Fecha decisión', 'Usada'];
const HAB_SI = ['HABILITADO', 'HABILITADO FUERA DE FECHA'];
try {
  ['habEstado', 'habPedir', 'habDecidir', 'habPendientes', 'subsanarGeneral'].forEach(f => { if (API_PUBLICA.indexOf(f) < 0) API_PUBLICA.push(f); });
  ['habEstado', 'habPendientes'].forEach(f => { if (SOLO_LECTURA.indexOf(f) < 0) SOLO_LECTURA.push(f); });
} catch (e) {}

const habEsCC_ = u => u.permisos.indexOf('revisar') >= 0 || u.permisos.indexOf('accesos') >= 0;
const habF_ = d => d instanceof Date ? Utilities.formatDate(d, 'America/Lima', 'dd/MM/yyyy HH:mm') : txt_(d);
function habHoja_() {
  let h = ss_().getSheetByName(HOJA_HAB);
  if (!h) {
    h = ss_().insertSheet(HOJA_HAB);
    h.getRange(1, 1, 1, HAB_HDR.length).setValues([HAB_HDR]).setFontWeight('bold').setBackground('#0C1A3D').setFontColor('#FFFFFF');
    h.setFrozenRows(1); h.setColumnWidth(5, 320);
  }
  return h;
}
/** Última fila de habilitación del expediente: {fila, estado, pidio, fPedido, motivo, decidio, fDecision} */
function habUlt_(id) {
  const h = habHoja_(), n = h.getLastRow(); if (n < 2) return null;
  const v = h.getRange(2, 1, n - 1, HAB_HDR.length).getValues();
  for (let i = v.length - 1; i >= 0; i--) if (String(v[i][0]).trim() === id) {
    const r = v[i]; return { fila: i + 2, estado: txt_(r[1]), pidio: txt_(r[2]), fPedido: habF_(r[3]), motivo: txt_(r[4]), decidio: txt_(r[5]), fDecision: habF_(r[6]) };
  }
  return null;
}
function habExp_(token, id) {
  id = String(id || '').trim();
  const ex = misExpedientes(token).find(x => x.id === id); if (!ex) throw new Error('No tiene acceso a este expediente.');
  return ex;
}
/** Estado de bloqueo para el usuario que mira el expediente. */
function habCalc_(u, ex) {
  const h = habUlt_(ex.id), lim = typeof plazo_ === 'function' ? plazo_(ex.mes, 'X') : null, vencido = !!(lim && new Date() > lim);
  const out = { cc: habEsCC_(u), bloqueado: false, motivo: '', limite: lim ? fmtPlazo_(lim) : '', vencido: vencido,
    pedido: h && h.estado === 'PEDIDO' ? { pidio: h.pidio, fecha: h.fPedido, motivo: h.motivo } : null,
    habilitacion: h && HAB_SI.indexOf(h.estado) >= 0 ? { tipo: h.estado, decidio: h.decidio, fecha: h.fDecision } : null,
    rechazado: h && h.estado === 'RECHAZADO' ? { decidio: h.decidio, fecha: h.fDecision } : null,
    enviado: ex.sub ? ex.sub.fecha : '' };
  if (out.cc) return out;
  const fuera = out.habilitacion && out.habilitacion.tipo === 'HABILITADO FUERA DE FECHA', normal = !!out.habilitacion;
  if (ex.estado === 'SUBSANADO' && !normal) { out.bloqueado = true; out.motivo = 'ESPERA'; }
  else if (vencido && ex.estado === 'OBSERVADO' && !fuera) { out.bloqueado = true; out.motivo = 'PLAZO'; }
  return out;
}
function habEstado(token, id) { const u = sesion_(token); return habCalc_(u, habExp_(token, id)); }

/** Lo usan los envíos de Comercial: si está bloqueado no deja subsanar; si había habilitación, la marca USADA. */
function habAntesDeSubsanar_(token, id) {
  const u = sesion_(token), ex = habExp_(token, id), e = habCalc_(u, ex);
  if (e.bloqueado) throw new Error(e.motivo === 'ESPERA'
    ? '🔒 Ya subsanaste este expediente: espera la respuesta de Contratos y Cobranzas. Si te equivocaste, usa "Pedir habilitación".'
    : '🔒 Venció el plazo para subsanar (' + e.limite + '). Usa "Pedir habilitación" para solicitar una excepción.');
  return { u: u, ex: ex, hab: e.habilitacion };
}
function habUsar_(id, hab) {
  if (!hab) return;
  const x = habUlt_(id); if (x && HAB_SI.indexOf(x.estado) >= 0) habHoja_().getRange(x.fila, 8).setValue(new Date()).offset(0, -6).setValue('USADO · ' + x.estado);
}
const habNota_ = (hab) => hab && hab.tipo === 'HABILITADO FUERA DE FECHA' ? '\n⏰ Fuera de plazo autorizado por ' + hab.decidio + ' (' + hab.fecha + ')' : (hab ? '\n🔓 Habilitado por ' + hab.decidio : '');

/** "Tu respuesta" (mensaje general) de Comercial, con el control de bloqueo. */
function subsanarGeneral(token, id, item, texto, adjunto) {
  const c = habAntesDeSubsanar_(token, id);
  const r = subsanar(token, id, item, String(texto || '') + (c.ex && !habEsCC_(c.u) ? habNota_(c.hab) : ''), adjunto);
  if (!habEsCC_(c.u)) habUsar_(String(id).trim(), c.hab);
  return r;
}

/** Comercial pide que le habiliten el expediente (se equivocó al subsanar o se le venció el plazo). */
function habPedir(token, id, motivo) {
  const u = sesion_(token); puede_(u, 'subsanar');
  const ex = habExp_(token, id), e = habCalc_(u, ex);
  if (e.cc) throw new Error('Tu usuario ya puede responder: no necesita habilitación.');
  if (!e.bloqueado) throw new Error('Este expediente no está bloqueado: ya puedes subsanar.');
  motivo = String(motivo || '').trim(); if (motivo.length < 5) throw new Error('Escribe por qué necesitas que te habiliten.');
  if (e.pedido) throw new Error('Ya pediste habilitación el ' + e.pedido.fecha + ': espera la respuesta.');
  habHoja_().appendRow(fila_([ex.id, 'PEDIDO', u.nombre, new Date(), motivo.slice(0, 500), '', '', '']));
  log_(u.nombre, 'HAB_PEDIDO', ex.id + ' · ' + (e.motivo === 'PLAZO' ? 'plazo vencido' : 'espera revisión') + ' · ' + motivo);
  try { MailApp.sendEmail(admins_(), '🙋 Piden habilitar ' + ex.id + ' · ' + u.nombre, u.nombre + ' pide habilitar el expediente ' + ex.id + ' (' + ex.cliente + ' · ' + ex.proyecto + ' ' + ex.unidad + ').\nMotivo: ' + motivo + '\n' + (e.motivo === 'PLAZO' ? 'El plazo venció el ' + e.limite + '.' : 'Ya subsanó y espera la revisión.') + '\n\nResponde en el portal (🔔 o en el expediente).'); } catch (z) {}
  return { ok: true };
}

/** Contratos / administración: HABILITADO · HABILITADO FUERA DE FECHA · RECHAZADO (con o sin pedido previo). */
function habDecidir(token, id, decision) {
  const u = sesion_(token); if (!habEsCC_(u)) puede_(u, 'revisar');
  const ex = habExp_(token, id); decision = String(decision || '').toUpperCase();
  if (HAB_SI.concat(['RECHAZADO']).indexOf(decision) < 0) throw new Error('Decisión no válida.');
  const x = habUlt_(ex.id), ahora = new Date(), h = habHoja_();
  if (x && x.estado === 'PEDIDO') h.getRange(x.fila, 2).setValue(decision).offset(0, 4, 1, 2).setValues([[u.nombre, ahora]]);
  else h.appendRow(fila_([ex.id, decision, '', '', '(sin pedido: lo decidió ' + u.nombre + ')', u.nombre, ahora, '']));
  log_(u.nombre, 'HAB_' + decision.replace(/ /g, '_'), ex.id);
  return { decision: decision };
}

/** Pedidos de habilitación pendientes (para la 🔔 de Contratos). */
function habPendientes(token) {
  const u = sesion_(token); if (!habEsCC_(u)) return [];
  const h = habHoja_(), n = h.getLastRow(); if (n < 2) return [];
  const ult = {}; h.getRange(2, 1, n - 1, HAB_HDR.length).getValues().forEach(r => { ult[String(r[0]).trim()] = r; });
  const exps = {}; misExpedientes(token).forEach(x => exps[x.id] = x);
  return Object.keys(ult).filter(id => txt_(ult[id][1]) === 'PEDIDO' && exps[id]).map(id => ({ id: id, cliente: exps[id].cliente, proyecto: exps[id].proyecto, unidad: exps[id].unidad,
    pidio: txt_(ult[id][2]), fecha: habF_(ult[id][3]), motivo: txt_(ult[id][4]) }));
}
