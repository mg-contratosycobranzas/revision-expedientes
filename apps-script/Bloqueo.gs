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
  out.excepcion = typeof habExcDe_ === 'function' ? habExcDe_(u) : null;   // excepción de plazo dada desde Accesos (usuario o TODOS)
  const fuera = (out.habilitacion && out.habilitacion.tipo === 'HABILITADO FUERA DE FECHA') || !!out.excepcion, normal = !!out.habilitacion;
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

/* ---------- EXCEPCIÓN DE PLAZO por usuario o para TODOS (administración, desde Accesos) ----------
 * Permite subsanar aunque haya vencido "Comercial presenta subsanados" del CALENDARIO, hasta la fecha y hora indicada.
 * No quita la espera de revisión (ESPERA); solo el bloqueo por PLAZO. La subsanación igual queda marcada "fuera de plazo" en el historial.
 * Hoja SUBSANACION_EXCEPCION: Persona (nombre de Accesos o TODOS) | Hasta | Motivo | Otorgó | Fecha | Estado (ACTIVA / QUITADA) */
const HOJA_EXC = 'SUBSANACION_EXCEPCION', EXC_HDR = ['Persona', 'Hasta', 'Motivo', 'Otorgó', 'Fecha', 'Estado'];
try { ['habExcLista', 'habExcGuardar', 'habExcQuitar'].forEach(f => { if (API_PUBLICA.indexOf(f) < 0) API_PUBLICA.push(f); }); if (SOLO_LECTURA.indexOf('habExcLista') < 0) SOLO_LECTURA.push('habExcLista'); } catch (e) {}
function habExcHoja_() {
  let h = ss_().getSheetByName(HOJA_EXC);
  if (!h) { h = ss_().insertSheet(HOJA_EXC); h.getRange(1, 1, 1, EXC_HDR.length).setValues([EXC_HDR]).setFontWeight('bold').setBackground('#0C1A3D').setFontColor('#FFFFFF'); h.setFrozenRows(1); h.setColumnWidth(3, 300); }
  return h;
}
/** Excepciones ACTIVAS y vigentes: [{fila, persona, hasta(Date), motivo, otorgo, fecha}] */
function habExcVigentes_() {
  const h = ss_().getSheetByName(HOJA_EXC); if (!h || h.getLastRow() < 2) return [];
  const ahora = new Date();
  return h.getRange(2, 1, h.getLastRow() - 1, EXC_HDR.length).getValues().map((r, i) => ({ fila: i + 2, persona: String(r[0]).trim(), hasta: r[1], motivo: txt_(r[2]), otorgo: txt_(r[3]), fecha: r[4], estado: norm_(r[5]) }))
    .filter(x => x.estado === 'ACTIVA' && x.hasta instanceof Date && x.hasta >= ahora);
}
/** La excepción que aplica a este usuario (la suya o la de TODOS; la que dure más). */
function habExcDe_(u) {
  const yo = [norm_(u.nombre), norm_(u.ref)].filter(Boolean);
  const l = habExcVigentes_().filter(x => norm_(x.persona) === 'TODOS' || yo.indexOf(norm_(x.persona)) >= 0).sort((a, b) => b.hasta - a.hasta);
  return l.length ? { persona: l[0].persona, hasta: habF_(l[0].hasta), motivo: l[0].motivo, otorgo: l[0].otorgo, todos: norm_(l[0].persona) === 'TODOS' } : null;
}
function habExcLista(token) {
  const u = sesion_(token); puede_(u, 'accesos');
  return habExcVigentes_().map(x => ({ persona: x.persona, hasta: habF_(x.hasta), hastaISO: Utilities.formatDate(x.hasta, 'America/Lima', "yyyy-MM-dd'T'HH:mm"), motivo: x.motivo, otorgo: x.otorgo, fecha: habF_(x.fecha), todos: norm_(x.persona) === 'TODOS' }));
}
/** persona: nombre del usuario (como en Accesos) o 'TODOS' · hastaISO: 'yyyy-MM-ddTHH:mm' (hora de Lima) */
function habExcGuardar(token, persona, hastaISO, motivo) {
  const u = sesion_(token); puede_(u, 'accesos');
  persona = String(persona || '').trim(); motivo = String(motivo || '').trim().slice(0, 300);
  if (!persona) throw new Error('Elige a quién.');
  const m = String(hastaISO || '').match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})$/); if (!m) throw new Error('Indica la fecha y hora hasta cuándo.');
  const hasta = new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5]); if (hasta <= new Date()) throw new Error('La fecha "hasta" debe ser futura.');
  if (hasta - new Date() > 31 * 864e5) throw new Error('Máximo 31 días: si necesitas más, vuelve a darla después.');
  const h = habExcHoja_(), lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    habExcVigentes_().filter(x => norm_(x.persona) === norm_(persona)).forEach(x => h.getRange(x.fila, 6).setValue('QUITADA'));   // reemplaza la anterior de esa persona
    h.appendRow(fila_([norm_(persona) === 'TODOS' ? 'TODOS' : persona, hasta, motivo, u.nombre, new Date(), 'ACTIVA']));
  } finally { lock.releaseLock(); }
  log_(u.nombre, 'EXCEPCION_PLAZO', persona + ' hasta ' + habF_(hasta) + (motivo ? ' · ' + motivo : ''));
  return { persona: persona, hasta: habF_(hasta) };
}
function habExcQuitar(token, persona) {
  const u = sesion_(token); puede_(u, 'accesos'); persona = String(persona || '').trim();
  const h = habExcHoja_(), l = habExcVigentes_().filter(x => norm_(x.persona) === norm_(persona)); if (!l.length) throw new Error('No hay una excepción vigente para ' + persona + '.');
  l.forEach(x => h.getRange(x.fila, 6).setValue('QUITADA'));
  log_(u.nombre, 'EXCEPCION_PLAZO_QUITADA', persona);
  return { ok: true };
}
