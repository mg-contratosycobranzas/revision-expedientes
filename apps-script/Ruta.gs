/**
 * RUTA DEL FILE — control de entrega del expediente físico y de sus tiempos (archivo independiente).
 *   1) Vendedor: "Presentar file" = pre-registro en la hoja RUTA_FILE (NO toca BASE MAESTRA). Código PR-2026-0001.
 *   2) Asistente comercial (permiso "bajar", p. ej. Lucía Rosales): "Recibir y revisar" con el checklist
 *      → completo = RECIBIDO · incompleto = DEVUELTO al vendedor con la lista de faltantes.
 *   3) Asistente comercial: "Entregar a Contratos" → una entrega (lote ER-2026-0001) con varios files completos.
 *   4) Contratos y Cobranzas: confirma la recepción del lote y "Valida y registra" cada PR → recién ahí
 *      se escribe en BASE MAESTRA (registrarVenta) y nace el EXP, enlazado al PR. También puede devolverlo.
 * Cada paso queda en RUTA_LOG (fecha, hora, usuario) para medir cuánto tiempo pasa el file en cada etapa.
 * Se conecta solo con Code.gs: agrega sus funciones a API_PUBLICA y los permisos "presentar" y "bajar".
 */
const HOJA_RUTA = 'RUTA_FILE', HOJA_RUTA_LOG = 'RUTA_LOG', RUTA_ROJO_H = 24;
const RUTA_FN = ['rutaOpciones', 'rutaPresentar', 'rutaMias', 'rutaBandeja', 'rutaRecibir', 'rutaEntregar', 'rutaRecibirLote', 'rutaDevolver', 'rutaValidar', 'rutaPanel', 'rutaDeExp', 'anularCargo', 'agregarOpcion'];
try {
  RUTA_FN.forEach(f => { if (API_PUBLICA.indexOf(f) < 0) API_PUBLICA.push(f); });
  ['presentar', 'bajar'].forEach(p => { if (PERMISOS.indexOf(p) < 0) PERMISOS.push(p); });
  ['VENDEDOR', 'JEFE', 'APOYO'].forEach(r => { if (PERMISOS_BASE[r] && PERMISOS_BASE[r].indexOf('presentar') < 0) PERMISOS_BASE[r].push('presentar'); });
} catch (e) {}

// Columnas de RUTA_FILE
const RC = { pr: 1, estado: 2, fecha: 3, vendedor: 4, presento: 5, proyecto: 6, tipo: 7, unidad: 8, cliente: 9, dni: 10, tipoCliente: 11, fin: 12, nota: 13,
  recibio: 14, fRecibio: 15, docs: 16, faltantes: 17, devol: 18, lote: 19, fEntrega: 20, recibioCC: 21, fRecibioCC: 22, exp: 23, valido: 24, fValido: 25, fPrimera: 26, folios: 27 };
const RUTA_NC = 27;
const RUTA_HDR = ['PR', 'Estado', 'Fecha presentación', 'Vendedor', 'Presentó (usuario)', 'Proyecto', 'Tipo', 'Unidad', 'Cliente', 'DNI / CE / RUC', 'Tipo de cliente', 'Financiamiento', 'Nota del vendedor',
  'Recibió (asistente)', 'Fecha recibido', 'Documentos completos', 'Faltantes / motivo', 'N° devoluciones', 'Entrega (lote)', 'Fecha entrega a Contratos', 'Recibió Contratos', 'Fecha recepción Contratos',
  'EXP', 'Validó', 'Fecha validación', 'Primera presentación', 'Folios'];
// Quién tiene el file después de cada evento (para medir el cuello de botella)
const RUTA_ETAPA = { PRESENTADO: 'Asistente: recibir', RECIBIDO: 'Asistente: entregar', DEVUELTO: 'Vendedor: corregir', ENTREGADO: 'Contratos: recibir', 'EN CONTRATOS': 'Contratos: validar' };

function hojaRuta_() {
  const ss = ss_(); let h = ss.getSheetByName(HOJA_RUTA);
  if (!h) {
    h = ss.insertSheet(HOJA_RUTA); h.appendRow(RUTA_HDR); h.setFrozenRows(1);
    h.getRange(1, 1, 1, RUTA_NC).setFontWeight('bold').setBackground('#0C1A3D').setFontColor('#FFFFFF');
    h.getRange(2, RC.dni, 998, 1).setNumberFormat('@');
  }
  if (!ss.getSheetByName(HOJA_RUTA_LOG)) {
    const l = ss.insertSheet(HOJA_RUTA_LOG); l.appendRow(['Fecha', 'PR', 'Evento', 'Usuario', 'Detalle']); l.setFrozenRows(1);
    l.getRange('A1:E1').setFontWeight('bold').setBackground('#0C1A3D').setFontColor('#FFFFFF');
  }
  return h;
}
function rutaLog_(pr, evento, usuario, detalle) { RUTA_LOG_CACHE_ = null; ss_().getSheetByName(HOJA_RUTA_LOG).appendRow(fila_([new Date(), pr, evento, usuario, String(detalle || '').slice(0, 300)])); }
const rFecha_ = d => d instanceof Date ? Utilities.formatDate(d, 'America/Lima', 'dd/MM/yyyy HH:mm') : '';
function rutaFilas_() {
  const h = hojaRuta_(), n = h.getLastRow(); if (n < 2) return [];
  return h.getRange(2, 1, n - 1, RUTA_NC).getValues().map((r, i) => ({ fila: i + 2, r: r }));
}
function rutaObj_(x) {
  const r = x.r, g = k => r[RC[k] - 1], est = String(g('estado'));
  const desde = { PRESENTADO: g('fecha'), RECIBIDO: g('fRecibio'), DEVUELTO: g('fecha'), ENTREGADO: g('fEntrega'), 'EN CONTRATOS': g('fRecibioCC'), VALIDADO: g('fValido') }[est];
  const ult = rutaUltimo_(String(g('pr')));
  const t0 = (est === 'DEVUELTO' && ult) ? ult : desde;
  return { pr: String(g('pr')), estado: est, fecha: rFecha_(g('fecha')), vendedor: String(g('vendedor')), presento: String(g('presento')), proyecto: String(g('proyecto')), tipo: String(g('tipo')),
    unidad: String(g('unidad')), cliente: String(g('cliente')), dni: String(g('dni')), tipoCliente: String(g('tipoCliente')), fin: String(g('fin')), nota: String(g('nota')),
    recibio: String(g('recibio')), fRecibio: rFecha_(g('fRecibio')), faltantes: String(g('faltantes')), devol: Number(g('devol')) || 0, lote: String(g('lote')), fEntrega: rFecha_(g('fEntrega')),
    recibioCC: String(g('recibioCC')), fRecibioCC: rFecha_(g('fRecibioCC')), exp: String(g('exp')), valido: String(g('valido')), fValido: rFecha_(g('fValido')), folios: String(g('folios')),
    fechaISO: g('fecha') instanceof Date ? Utilities.formatDate(g('fecha'), 'America/Lima', 'yyyy-MM-dd') : '',
    horas: t0 instanceof Date && est !== 'VALIDADO' ? Math.round((new Date() - t0) / 36e5 * 10) / 10 : 0 };
}
let RUTA_LOG_CACHE_ = null;
function rutaLogTodo_() {
  if (RUTA_LOG_CACHE_) return RUTA_LOG_CACHE_;
  const l = ss_().getSheetByName(HOJA_RUTA_LOG); if (!l || l.getLastRow() < 2) return (RUTA_LOG_CACHE_ = []);
  return (RUTA_LOG_CACHE_ = l.getRange(2, 1, l.getLastRow() - 1, 5).getValues().filter(r => r[0] instanceof Date));
}
function rutaUltimo_(pr) { const e = rutaLogTodo_().filter(r => String(r[1]) === pr); return e.length ? e[e.length - 1][0] : null; }

function rutaVendedores_(u) {
  const v = vendedoresActivos_();
  if (u.rol === 'VENDEDOR' || u.rol === 'APOYO') return v.filter(x => norm_(x.nombre) === norm_(u.ref)).map(x => x.nombre);
  if (u.rol === 'JEFE') return v.filter(x => norm_(x.jefe) === norm_(u.ref) || norm_(x.nombre) === norm_(u.ref)).map(x => x.nombre);
  return v.map(x => x.nombre);
}
function rutaPuede_(u, ...ps) { if (u.permisos.indexOf('accesos') >= 0 || ps.some(p => u.permisos.indexOf(p) >= 0)) return; puede_(u, ps[0]); }
function rutaVe_(u, o) {
  if (u.permisos.indexOf('accesos') >= 0 || u.permisos.indexOf('bajar') >= 0 || u.permisos.indexOf('revisar') >= 0 || ['ADMIN', 'ASISTENTE', 'GERENTE'].indexOf(u.rol) >= 0) return true;
  return rutaVendedores_(u).some(n => norm_(n) === norm_(o.vendedor)) || norm_(o.presento) === norm_(u.nombre);
}
function rutaBuscar_(pr) { const x = rutaFilas_().find(f => String(f.r[0]) === String(pr)); if (!x) throw new Error('No encontré ' + pr + '.'); return x; }
function rutaSet_(fila, cambios) { const h = ss_().getSheetByName(HOJA_RUTA); Object.keys(cambios).forEach(k => h.getRange(fila, RC[k]).setValue(seguro_(cambios[k]))); }
function rutaDocsReq_(tipoCliente, fin) {
  const out = [];
  CHECK_GRUPOS.filter(g => !g.si.length || g.si.indexOf(tipoCliente) >= 0 || g.si.indexOf(fin) >= 0).forEach(g => g.items.forEach(it => out.push(g.t + '|' + it)));
  return out;
}

function rutaOpciones(token) {
  const u = sesion_(token); rutaPuede_(u, 'presentar', 'bajar', 'revisar');
  const bm = ss_().getSheetByName(HOJA_BM);
  return { proyecto: lista_(bm, 'E5', proyectos_(bm)), tipo: lista_(bm, 'F5', ['DEPARTAMENTO', 'ESTACIONAMIENTO', 'DEPOSITO']), vendedor: rutaVendedores_(u),
    tipos: CHECK_TIPOS, fin: CHECK_FIN, grupos: CHECK_GRUPOS, rojo: RUTA_ROJO_H };
}

/** Vendedor presenta (o vuelve a presentar, si le devolvieron) un file. */
function rutaPresentar(token, d, prPrevio) {
  const u = sesion_(token); rutaPuede_(u, 'presentar'); d = d || {};
  const req = { proyecto: 'Proyecto', tipo: 'Tipo', unidad: 'N° de unidad', cliente: 'Cliente', dni: 'DNI / CE / RUC', vendedor: 'Vendedor', tipoCliente: 'Tipo de cliente', fin: 'Financiamiento' };
  for (const k in req) if (!String(d[k] == null ? '' : d[k]).trim()) throw new Error('Falta: ' + req[k]);
  const dni = String(d.dni).replace(/\D/g, ''); if ([8, 9, 11].indexOf(dni.length) < 0) throw new Error('DNI / CE / RUC debe tener 8, 9 u 11 dígitos.');
  if (!rutaVendedores_(u).some(n => norm_(n) === norm_(d.vendedor))) throw new Error('Vendedor no permitido para su usuario.');
  if (!CHECK_TIPOS[d.tipoCliente] || !CHECK_FIN[d.fin]) throw new Error('Elige el tipo de cliente y el financiamiento.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const filas = rutaFilas_(), ahora = new Date(), uni = norm_(d.unidad).replace(/\s/g, ''), pro = norm_(d.proyecto);
    const otro = filas.find(f => String(f.r[0]) !== String(prPrevio || '') && ['VALIDADO', 'ANULADO'].indexOf(String(f.r[1])) < 0 && norm_(f.r[RC.proyecto - 1]) === pro && norm_(f.r[RC.unidad - 1]).replace(/\s/g, '') === uni);
    if (otro) throw new Error('Esa unidad ya fue presentada en ' + otro.r[0] + ' (' + otro.r[RC.cliente - 1] + ').');
    const datos = { fecha: ahora, vendedor: String(d.vendedor).trim(), presento: u.nombre, proyecto: String(d.proyecto).trim(), tipo: String(d.tipo).trim(), unidad: String(d.unidad).trim(),
      cliente: String(d.cliente).trim().toUpperCase(), dni: dni, tipoCliente: d.tipoCliente, fin: d.fin, nota: String(d.nota || '').trim().slice(0, 500) };
    if (prPrevio) {
      const x = rutaBuscar_(prPrevio), o = rutaObj_(x);
      if (o.estado !== 'DEVUELTO') throw new Error('Este file no está devuelto: no hace falta volver a presentarlo.');
      if (!rutaVe_(u, o)) throw new Error('No tiene acceso a este file.');
      datos.estado = 'PRESENTADO'; rutaSet_(x.fila, datos);
      rutaLog_(o.pr, 'PRESENTADO', u.nombre, 'Vuelve a presentar (devolución N° ' + o.devol + ')');
      log_(u.nombre, 'RUTA_PRESENTA', o.pr + ' · ' + datos.cliente + ' (de nuevo)');
      return { pr: o.pr };
    }
    const anio = Utilities.formatDate(ahora, 'America/Lima', 'yyyy'), n = filas.filter(f => String(f.r[0]).indexOf('PR-' + anio + '-') === 0).length + 1;
    const pr = 'PR-' + anio + '-' + ('000' + n).slice(-4), fila = new Array(RUTA_NC).fill('');
    fila[RC.pr - 1] = pr; fila[RC.estado - 1] = 'PRESENTADO'; fila[RC.devol - 1] = 0; fila[RC.fPrimera - 1] = ahora;
    Object.keys(datos).forEach(k => fila[RC[k] - 1] = datos[k]);
    hojaRuta_().appendRow(fila_(fila));
    rutaLog_(pr, 'PRESENTADO', u.nombre, datos.cliente + ' · ' + datos.proyecto + ' ' + datos.unidad);
    log_(u.nombre, 'RUTA_PRESENTA', pr + ' · ' + datos.cliente + ' · ' + datos.proyecto + ' ' + datos.unidad);
    return { pr: pr };
  } finally { lock.releaseLock(); }
}

/** Files del vendedor (o de su equipo, si es jefe). */
function rutaMias(token) {
  const u = sesion_(token); rutaPuede_(u, 'presentar', 'bajar', 'revisar');
  return rutaFilas_().map(rutaObj_).filter(o => o.estado !== 'ANULADO' && (rutaVendedores_(u).some(n => norm_(n) === norm_(o.vendedor)) || norm_(o.presento) === norm_(u.nombre))).reverse().slice(0, 200);
}

/** Bandeja de la asistente comercial y de Contratos: todo lo que no está validado (+ los últimos validados). */
function rutaBandeja(token) {
  const u = sesion_(token); rutaPuede_(u, 'bajar', 'revisar');
  const l = rutaFilas_().map(rutaObj_), act = l.filter(o => ['VALIDADO', 'ANULADO'].indexOf(o.estado) < 0), val = l.filter(o => o.estado === 'VALIDADO').slice(-30).reverse();
  return { lista: act.concat(val), rojo: RUTA_ROJO_H, puedeBajar: u.permisos.indexOf('bajar') >= 0 || u.permisos.indexOf('accesos') >= 0, puedeValidar: u.permisos.indexOf('revisar') >= 0 || u.permisos.indexOf('accesos') >= 0 };
}

/** Asistente: recibe el file y marca los documentos. Completo → RECIBIDO; incompleto → DEVUELTO al vendedor. */
function rutaRecibir(token, pr, docsOk, folios) {
  const u = sesion_(token); rutaPuede_(u, 'bajar');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const x = rutaBuscar_(pr), o = rutaObj_(x); if (o.estado !== 'PRESENTADO') throw new Error(pr + ' no está pendiente de recibir (' + o.estado + ').');
    const req = rutaDocsReq_(o.tipoCliente, o.fin), ok = (docsOk || []).filter(k => req.indexOf(k) >= 0), falta = req.filter(k => ok.indexOf(k) < 0).map(k => k.split('|')[1]);
    const ahora = new Date();
    if (!falta.length) {
      rutaSet_(x.fila, { estado: 'RECIBIDO', recibio: u.nombre, fRecibio: ahora, docs: req.length + ' de ' + req.length, faltantes: '', folios: String(folios || '').trim() });
      rutaLog_(pr, 'RECIBIDO', u.nombre, 'Completo ' + req.length + '/' + req.length + (folios ? ' · ' + folios + ' folios' : ''));
      log_(u.nombre, 'RUTA_RECIBE', pr + ' · completo');
      return { completo: true, total: req.length };
    }
    rutaSet_(x.fila, { estado: 'DEVUELTO', recibio: u.nombre, fRecibio: ahora, docs: ok.length + ' de ' + req.length, faltantes: 'Faltan: ' + falta.join(', '), devol: o.devol + 1 });
    rutaLog_(pr, 'DEVUELTO', u.nombre, 'Asistente: faltan ' + falta.length + ' (' + falta.join(', ') + ')');
    log_(u.nombre, 'RUTA_DEVUELVE', pr + ' · faltan ' + falta.length);
    return { completo: false, faltantes: falta, total: req.length, vendedor: o.vendedor, cliente: o.cliente };
  } finally { lock.releaseLock(); }
}

/** Asistente: entrega a Contratos varios files completos en un solo lote. */
function rutaEntregar(token, prs) {
  const u = sesion_(token); rutaPuede_(u, 'bajar');
  prs = (prs || []).map(String); if (!prs.length) throw new Error('Elige al menos un file.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const filas = rutaFilas_(), sel = prs.map(p => { const x = filas.find(f => String(f.r[0]) === p); if (!x) throw new Error('No encontré ' + p + '.'); const o = rutaObj_(x); if (o.estado !== 'RECIBIDO') throw new Error(p + ' no está listo para entregar (' + o.estado + ').'); return { x: x, o: o }; });
    const ahora = new Date(), anio = Utilities.formatDate(ahora, 'America/Lima', 'yyyy');
    const lotes = [...new Set(filas.map(f => String(f.r[RC.lote - 1])).filter(l => l.indexOf('ER-' + anio + '-') === 0))];
    const lote = 'ER-' + anio + '-' + ('000' + (lotes.length + 1)).slice(-4);
    sel.forEach(s => { rutaSet_(s.x.fila, { estado: 'ENTREGADO', lote: lote, fEntrega: ahora }); rutaLog_(s.o.pr, 'ENTREGADO', u.nombre, lote); });
    log_(u.nombre, 'RUTA_ENTREGA', lote + ' · ' + sel.length + ' files');
    const fh = rFecha_(ahora), texto = '📥 Entrega ' + lote + ' a Contratos y Cobranzas\nEntrega: ' + u.nombre + ' · ' + fh + '\n' + sel.map((s, i) => (i + 1) + '. ' + s.o.pr + ' · ' + s.o.cliente + ' · ' + s.o.proyecto + ' ' + s.o.unidad + (s.o.folios ? ' · ' + s.o.folios + ' folios' : '')).join('\n') + '\nTotal: ' + sel.length + ' file(s) completos.';
    return { lote: lote, fecha: fh, files: sel.map(s => s.o), texto: texto };
  } finally { lock.releaseLock(); }
}

/** Contratos confirma que recibió físicamente el lote. */
function rutaRecibirLote(token, lote) {
  const u = sesion_(token); rutaPuede_(u, 'revisar');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sel = rutaFilas_().filter(f => String(f.r[RC.lote - 1]) === String(lote) && String(f.r[1]) === 'ENTREGADO'); if (!sel.length) throw new Error('No hay files pendientes de recepción en ' + lote + '.');
    const ahora = new Date();
    sel.forEach(f => { rutaSet_(f.fila, { estado: 'EN CONTRATOS', recibioCC: u.nombre, fRecibioCC: ahora }); rutaLog_(String(f.r[0]), 'EN CONTRATOS', u.nombre, lote); });
    log_(u.nombre, 'RUTA_RECEPCION', lote + ' · ' + sel.length + ' files');
    return { n: sel.length };
  } finally { lock.releaseLock(); }
}

/** Devolver al vendedor: la asistente ("no lo tengo") o Contratos (con motivo). */
function rutaDevolver(token, pr, motivo) {
  const u = sesion_(token); rutaPuede_(u, 'bajar', 'revisar');
  motivo = String(motivo || '').trim(); if (motivo.length < 4) throw new Error('Escribe el motivo.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const x = rutaBuscar_(pr), o = rutaObj_(x);
    if (['PRESENTADO', 'RECIBIDO', 'ENTREGADO', 'EN CONTRATOS'].indexOf(o.estado) < 0) throw new Error(pr + ' no se puede devolver (' + o.estado + ').');
    rutaSet_(x.fila, { estado: 'DEVUELTO', faltantes: motivo, devol: o.devol + 1, lote: '' });
    rutaLog_(pr, 'DEVUELTO', u.nombre, motivo);
    log_(u.nombre, 'RUTA_DEVUELVE', pr + ' · ' + motivo);
    return { vendedor: o.vendedor, cliente: o.cliente };
  } finally { lock.releaseLock(); }
}

/** Contratos valida los datos y registra en BASE MAESTRA (crea el EXP), enlazado al PR. */
function rutaValidar(token, pr, d) {
  const u = sesion_(token); rutaPuede_(u, 'revisar');
  const x0 = rutaBuscar_(pr), o0 = rutaObj_(x0);
  if (['ENTREGADO', 'EN CONTRATOS'].indexOf(o0.estado) < 0) throw new Error(pr + ' todavía no llegó a Contratos (' + o0.estado + ').');
  const r = registrarVenta(token, d);   // valida y escribe en BASE MAESTRA (C:W), igual que "Nueva venta"
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const x = rutaBuscar_(pr), ahora = new Date(), c = { estado: 'VALIDADO', exp: r.id, valido: u.nombre, fValido: ahora };
    if (o0.estado === 'ENTREGADO') { c.recibioCC = u.nombre; c.fRecibioCC = ahora; rutaLog_(pr, 'EN CONTRATOS', u.nombre, 'Recepción al validar'); }
    rutaSet_(x.fila, c); rutaLog_(pr, 'VALIDADO', u.nombre, r.id);
    log_(u.nombre, 'RUTA_VALIDA', pr + ' → ' + r.id);
  } finally { lock.releaseLock(); }
  r.pr = pr; return r;
}

/** Ruta de un expediente ya validado (para el detalle del EXP). */
function rutaDeExp(token, id) {
  const u = sesion_(token); id = String(id || '').trim();
  const x = rutaFilas_().find(f => String(f.r[RC.exp - 1]) === id); if (!x) return null;
  const o = rutaObj_(x); if (!rutaVe_(u, o)) return null;
  o.eventos = rutaLogTodo_().filter(r => String(r[1]) === o.pr).map(r => ({ fecha: rFecha_(r[0]), evento: String(r[2]), usuario: String(r[3]), detalle: String(r[4]) }));
  return o;
}

/** Panel de cuellos de botella: tiempo promedio en cada etapa, retenidos ahora y devoluciones. */
function rutaPanel(token, desdeISO) {
  const u = sesion_(token); rutaPuede_(u, 'dashboard', 'revisar', 'bajar');
  const files = rutaFilas_().map(rutaObj_).filter(o => o.estado !== 'ANULADO' && (!desdeISO || o.fechaISO >= desdeISO));
  const prs = {}; files.forEach(o => prs[o.pr] = o);
  const ev = {}; rutaLogTodo_().forEach(r => { const p = String(r[1]); if (prs[p]) (ev[p] = ev[p] || []).push(r); });
  const ahora = new Date(), etapas = {}, add = (k, h, abierto) => { const e = etapas[k] = etapas[k] || { suma: 0, n: 0, abiertos: 0 }; e.suma += h; e.n++; if (abierto) e.abiertos++; };
  Object.keys(ev).forEach(p => {
    const l = ev[p];
    for (let i = 0; i < l.length; i++) {
      const et = RUTA_ETAPA[String(l[i][2])]; if (!et) continue;
      const fin = i + 1 < l.length ? l[i + 1][0] : (String(l[i][2]) === 'VALIDADO' ? null : ahora); if (!fin) continue;
      add(et, (fin - l[i][0]) / 36e5, i + 1 >= l.length);
    }
  });
  const ahoraEn = {}; files.filter(o => RUTA_ETAPA[o.estado]).forEach(o => { const k = RUTA_ETAPA[o.estado], a = ahoraEn[k] = ahoraEn[k] || { n: 0, max: 0, rojo: 0, quien: {} }; a.n++; a.max = Math.max(a.max, o.horas); if (o.horas >= RUTA_ROJO_H) a.rojo++; const q = o.estado === 'DEVUELTO' ? o.vendedor : o.estado === 'PRESENTADO' ? 'Asistente comercial' : o.estado === 'RECIBIDO' ? (o.recibio || 'Asistente comercial') : (o.recibioCC || 'Contratos y Cobranzas'); a.quien[q] = (a.quien[q] || 0) + 1; });
  const dev = {}; files.forEach(o => { if (o.devol) dev[o.vendedor] = (dev[o.vendedor] || 0) + o.devol; });
  const orden = ['Asistente: recibir', 'Vendedor: corregir', 'Asistente: entregar', 'Contratos: recibir', 'Contratos: validar'];
  return {
    total: files.length, validados: files.filter(o => o.estado === 'VALIDADO').length, rojo: RUTA_ROJO_H,
    etapas: orden.map(k => ({ etapa: k, promedio: etapas[k] ? Math.round(etapas[k].suma / etapas[k].n * 10) / 10 : 0, veces: etapas[k] ? etapas[k].n : 0, ahora: ahoraEn[k] ? ahoraEn[k].n : 0, max: ahoraEn[k] ? ahoraEn[k].max : 0, enRojo: ahoraEn[k] ? ahoraEn[k].rojo : 0, quien: ahoraEn[k] ? ahoraEn[k].quien : {} })),
    devoluciones: Object.keys(dev).map(k => ({ vendedor: k, n: dev[k] })).sort((a, b) => b.n - a.n).slice(0, 15)
  };
}

/* ---------- ANULAR CARGO DE ENTREGA (pruebas o errores) ----------
 * No borra nada: busca el N° de cargo (CE-2026-0001) en la hoja donde se guardan los cargos,
 * cambia su Estado a ANULADO y anota el motivo. Solo cargos pendientes de recepción.
 * Solo Contratos y Cobranzas (permiso "revisar") o administración. */
function anularCargo(token, numero, motivo) {
  const u = sesion_(token); rutaPuede_(u, 'revisar');
  numero = String(numero || '').trim(); motivo = String(motivo || '').trim();
  if (!/^CE-\d{4}-\d+$/i.test(numero)) throw new Error('N° de cargo inválido.');
  if (motivo.length < 3) throw new Error('Escribe el motivo de la anulación.');
  const L = s => norm_(s).normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const hojas = ss_().getSheets();
    for (let h = 0; h < hojas.length; h++) {
      const sh = hojas[h], nr = sh.getLastRow(), nc = sh.getLastColumn(); if (nr < 2 || nc < 2) continue;
      const v = sh.getRange(1, 1, nr, Math.min(nc, 40)).getDisplayValues();
      // fila de encabezados: la primera (de las 6 primeras) que tenga "ESTADO" y algo de "CARGO"
      const hi = v.slice(0, 6).findIndex(r => r.some(x => /^ESTADO/.test(L(x))) && r.some(x => /CARGO|N.? ?CE/.test(L(x))));
      if (hi < 0) continue;
      const hd = v[hi].map(L), cE = hd.findIndex(x => /^ESTADO/.test(x)), cO = hd.findIndex(x => /^OBSERV/.test(x));
      const fi = v.findIndex((r, i) => i > hi && r.slice(0, 4).some(x => String(x).trim().toUpperCase() === numero.toUpperCase()));
      if (fi < 0) continue;
      const est = norm_(v[fi][cE]);
      if (est === 'ANULADO') throw new Error(numero + ' ya está anulado.');
      if (est && est !== 'PENDIENTE') throw new Error(numero + ' ya fue recibido (' + est + '): no se puede anular.');
      sh.getRange(fi + 1, cE + 1).setValue('ANULADO');
      const nota = 'ANULADO por ' + u.nombre + ' · ' + Utilities.formatDate(new Date(), 'America/Lima', 'dd/MM/yyyy HH:mm') + ' · ' + motivo;
      if (cO >= 0) { const prev = String(v[fi][cO] || '').trim(); sh.getRange(fi + 1, cO + 1).setValue(seguro_(prev ? prev + ' | ' + nota : nota)); }
      else sh.getRange(fi + 1, cE + 1).setNote(nota);
      log_(u.nombre, 'CARGO_ANULADO', numero + ' · ' + motivo + ' · hoja ' + sh.getName());
      return { numero: numero, hoja: sh.getName() };
    }
  } finally { lock.releaseLock(); }
  throw new Error('No encontré el cargo ' + numero + ' en el Sheet.');
}

/* ---------- AGREGAR UNA OPCIÓN A LAS LISTAS (solo administración, desde el portal) ----------
 * proyecto → primera casilla libre de CATÁLOGOS bajo "Proyecto" (A34:A59, la lista que usa BASE MAESTRA).
 * empresa / financiamiento / banco → se suma a la lista desplegable de esa columna de BASE MAESTRA (O / T / U).
 * Así una incidencia como "falta MG PROYECTO 26" se corrige al momento, sin abrir el Sheet. */
function agregarOpcion(token, campo, valor) {
  const u = sesion_(token); puede_(u, 'accesos');
  valor = String(valor || '').replace(/\s+/g, ' ').trim().toUpperCase();
  if (valor.length < 2 || valor.length > 60 || /^[=+\-@]/.test(valor)) throw new Error('Escribe un nombre válido.');
  const ss = ss_(), bm = ss.getSheetByName(HOJA_BM), igual = x => norm_(x) === valor;
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (campo === 'proyecto') {
      const cat = ss.getSheetByName('CATÁLOGOS'), a = cat.getRange(1, 1, 120, 1).getDisplayValues().map(r => String(r[0]).trim());
      const hi = a.findIndex(x => norm_(x) === 'PROYECTO'); if (hi < 0) throw new Error('No encontré la lista "Proyecto" en CATÁLOGOS.');
      const lista = a.slice(hi + 1, hi + 27);   // 26 casillas (A34:A59), las que lee BASE MAESTRA
      if (lista.some(igual)) throw new Error(valor + ' ya está en la lista de proyectos.');
      const libre = lista.findIndex(x => !x); if (libre < 0) throw new Error('La lista de proyectos de CATÁLOGOS está llena (26). Amplíala en el Sheet.');
      cat.getRange(hi + 2 + libre, 1).setValue(valor);
    } else {
      const col = { empresa: 15, financiamiento: 20, banco: 21 }[campo]; if (!col) throw new Error('Lista no válida.');
      const dv = bm.getRange(PRIMERA_FILA, col).getDataValidation(); if (!dv) throw new Error('Esa columna no tiene lista desplegable en BASE MAESTRA.');
      const t = dv.getCriteriaType(), cv = dv.getCriteriaValues();
      if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
        const l = cv[0].map(x => String(x).trim()).filter(String); if (l.some(igual)) throw new Error(valor + ' ya está en la lista.');
        const regla = dv.copy().requireValueInList(l.concat([valor]), cv[1] !== false).build();
        bm.getRange(PRIMERA_FILA, col, ULTIMA_FILA - PRIMERA_FILA + 1, 1).setDataValidation(regla);
      } else if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
        const rg = cv[0], v = rg.getDisplayValues().map(r => String(r[0]).trim());
        if (v.some(igual)) throw new Error(valor + ' ya está en la lista.');
        const libre = v.findIndex(x => !x); if (libre < 0) throw new Error('La lista está llena: amplíala en el Sheet.');
        rg.getSheet().getRange(rg.getRow() + libre, rg.getColumn()).setValue(valor);
      } else throw new Error('Esa lista no se puede editar desde el portal.');
    }
  } finally { lock.releaseLock(); }
  log_(u.nombre, 'LISTA_AGREGADA', campo + ': ' + valor);
  return { campo: campo, valor: valor };
}

/* ---------- ACTUALIZAR SISTEMA DESDE EL PORTAL (solo administración) ----------
 * Trae de GitHub (carpeta apps-script/ del repositorio del portal) los archivos listados en manifest.json,
 * los instala en este proyecto, crea una versión nueva y la PRUEBA en una implementación aparte antes de
 * aplicarla a la del portal. Si la prueba falla, no cambia nada. Code.gs y appsscript.json nunca se tocan.
 * Requisitos (una sola vez): API de Apps Script activada en script.google.com/home/usersettings,
 * los permisos script.projects y script.deployments en appsscript.json, y ejecutar autorizarActualizador(). */
const ACT_REPO = 'https://raw.githubusercontent.com/mg-contratosycobranzas/revision-expedientes/main/apps-script/';
const ACT_PROTEGIDOS = ['Code', 'appsscript'];
try { ['sistemaEstado', 'actualizarSistema', 'revertirSistema'].forEach(f => { if (API_PUBLICA.indexOf(f) < 0) API_PUBLICA.push(f); }); } catch (e) {}

function actApi_(metodo, ruta, cuerpo) {
  const op = { method: metodo, muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } };
  if (cuerpo) { op.contentType = 'application/json'; op.payload = JSON.stringify(cuerpo); }
  const r = UrlFetchApp.fetch('https://script.googleapis.com/v1/projects/' + ScriptApp.getScriptId() + ruta, op), c = r.getResponseCode(), t = r.getContentText();
  if (c >= 300) {
    if (/has not been used|is disabled|SERVICE_DISABLED|User has not enabled/i.test(t)) throw new Error('Falta activar la "API de Google Apps Script" en script.google.com/home/usersettings.');
    if (c === 403) throw new Error('Falta autorizar el actualizador: en el editor de Apps Script ejecuta autorizarActualizador una vez.');
    throw new Error('API de Apps Script (' + c + '): ' + t.slice(0, 180));
  }
  return t ? JSON.parse(t) : {};
}
function actGet_(archivo) {
  const r = UrlFetchApp.fetch(ACT_REPO + archivo + '?t=' + Date.now(), { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('No pude descargar ' + archivo + ' de GitHub (' + r.getResponseCode() + ').');
  return r.getContentText();
}
function actManifest_() {
  const m = JSON.parse(actGet_('manifest.json'));
  (m.archivos || []).forEach(n => { if (!/^[A-Za-z0-9_]+$/.test(n) || ACT_PROTEGIDOS.indexOf(n) >= 0) throw new Error('Archivo no permitido en manifest: ' + n); });
  return m;
}
const actIdDe_ = url => (String(url || '').match(/\/macros\/s\/([A-Za-z0-9_-]+)\//) || [])[1] || '';

/** Ejecutar UNA vez desde el editor (▶ Ejecutar) para aceptar los permisos nuevos. */
function autorizarActualizador() {
  actApi_('get', '/content');
  Logger.log('Listo: el actualizador tiene permiso. Ya puedes usar "🔄 Actualizar sistema" en Accesos del portal.');
}

function sistemaEstado(token) {
  const u = sesion_(token); puede_(u, 'accesos');
  const p = PropertiesService.getScriptProperties(), m = actManifest_();
  return { instalada: p.getProperty('ACT_VER') || '(instalación manual)', disponible: m.version, notas: m.notas || '', archivos: m.archivos || [],
    anterior: p.getProperty('ACT_PREV') || '', fecha: p.getProperty('ACT_FECHA') || '' };
}

function actualizarSistema(token, urlPortal) {
  const u = sesion_(token); puede_(u, 'accesos');
  const mainId = actIdDe_(urlPortal); if (!mainId) throw new Error('No reconocí la dirección del motor del portal.');
  const p = PropertiesService.getScriptProperties(), m = actManifest_();
  const lock = LockService.getScriptLock(); if (!lock.tryLock(30000)) throw new Error('Ya hay una actualización en curso.');
  try {
    const deps = (actApi_('get', '/deployments').deployments || []), main = deps.find(d => d.deploymentId === mainId);
    if (!main) throw new Error('No encontré la implementación del portal en este proyecto.');
    const antes = actApi_('get', '/content'), limpio = f => ({ name: f.name, type: f.type, source: f.source });
    const files = antes.files.map(limpio);
    m.archivos.forEach(n => {
      const src = actGet_(n + '.gs'); if (src.length < 50) throw new Error(n + '.gs llegó vacío.');
      const i = files.findIndex(f => f.name === n), nf = { name: n, type: 'SERVER_JS', source: src };
      if (i >= 0) files[i] = nf; else files.push(nf);
    });
    actApi_('put', '/content', { scriptId: ScriptApp.getScriptId(), files: files });
    const ver = actApi_('post', '/versions', { description: 'Actualización ' + m.version }).versionNumber;
    // 1) prueba en una implementación aparte (el portal sigue intacto)
    const cfg = { scriptId: ScriptApp.getScriptId(), versionNumber: ver, manifestFileName: 'appsscript', description: 'Prueba ' + m.version };
    let testId = p.getProperty('ACT_TEST_DEP');
    if (testId && deps.some(d => d.deploymentId === testId)) actApi_('put', '/deployments/' + testId, { deploymentConfig: cfg });
    else { testId = actApi_('post', '/deployments', cfg).deploymentId; p.setProperty('ACT_TEST_DEP', testId); }
    let ok = false, det = '';
    for (let k = 0; k < 3 && !ok; k++) {
      if (k) Utilities.sleep(3000);
      try { const r = UrlFetchApp.fetch('https://script.google.com/macros/s/' + testId + '/exec', { muteHttpExceptions: true, followRedirects: true }); det = r.getResponseCode() + ' ' + r.getContentText().slice(0, 80); ok = r.getResponseCode() === 200 && JSON.parse(r.getContentText()).ok === true; } catch (e) { det = String(e.message || e).slice(0, 80); }
    }
    if (!ok) { actApi_('put', '/content', { scriptId: ScriptApp.getScriptId(), files: antes.files.map(limpio) }); throw new Error('La actualización ' + m.version + ' no pasó la prueba y NO se aplicó (' + det + '). El portal sigue igual.'); }
    // 2) aplica al portal
    const prev = main.deploymentConfig && main.deploymentConfig.versionNumber;
    actApi_('put', '/deployments/' + mainId, { deploymentConfig: { scriptId: ScriptApp.getScriptId(), versionNumber: ver, manifestFileName: 'appsscript', description: 'v' + ver + ' · ' + m.version } });
    p.setProperties({ ACT_VER: m.version, ACT_PREV: String(prev || ''), ACT_FECHA: Utilities.formatDate(new Date(), 'America/Lima', 'dd/MM/yyyy HH:mm') });
    log_(u.nombre, 'SISTEMA_ACTUALIZADO', m.version + ' · versión ' + ver + ' (antes ' + prev + ')');
    return { version: m.version, numero: ver, anterior: prev, notas: m.notas || '' };
  } finally { lock.releaseLock(); }
}

/** Vuelve el portal a la versión anterior a la última actualización. */
function revertirSistema(token, urlPortal) {
  const u = sesion_(token); puede_(u, 'accesos');
  const mainId = actIdDe_(urlPortal), p = PropertiesService.getScriptProperties(), prev = Number(p.getProperty('ACT_PREV'));
  if (!mainId || !prev) throw new Error('No hay una versión anterior guardada.');
  actApi_('put', '/deployments/' + mainId, { deploymentConfig: { scriptId: ScriptApp.getScriptId(), versionNumber: prev, manifestFileName: 'appsscript', description: 'v' + prev + ' · revertido' } });
  p.setProperties({ ACT_VER: 'revertido a v' + prev, ACT_PREV: '' });
  log_(u.nombre, 'SISTEMA_REVERTIDO', 'versión ' + prev);
  return { numero: prev };
}
