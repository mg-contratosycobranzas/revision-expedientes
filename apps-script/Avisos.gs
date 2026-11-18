/**
 * 📧 AVISOS POR CORREO — recordatorios del plazo de subsanación (archivo independiente; se conecta solo con Code.gs).
 * Una vez al día, a la hora elegida en el portal, revisa los expedientes que tienen puntos pendientes y, según el
 * CALENDARIO de cierre ("Comercial presenta subsanados"), envía a cada vendedor UN correo con SUS expedientes:
 * qué falta subsanar, qué ya está en revisión y cuándo vence.
 *   Etapas: 2 días antes · 1 día antes · el mismo día · al día siguiente de vencer (cada una se puede apagar).
 *   Copia: su jefe de ventas (opcional) y los correos fijos que indique la administración (p. ej. mvalles@grupomg.pe).
 *   Resumen diario: un solo correo a la administración con todo lo que vence o venció.
 * Cada aviso sale UNA sola vez por expediente y etapa: todo queda anotado en la hoja AVISOS_CORREO.
 * Los correos salen desde la cuenta dueña del Apps Script (la de Contratos y Cobranzas).
 */
const HOJA_AVI = 'AVISOS_CORREO';
const AVI_HDR = ['Fecha', 'Tipo', 'Vendedor', 'Para', 'Copia', 'Expedientes', 'Asunto', 'Resultado', 'Claves'];
const AVI_DEF = { activo: false, hora: 8, d2: true, d1: true, d0: true, vencido: true, resumen: true, copiaJefe: true,
  copia: 'mvalles@grupomg.pe', portal: 'https://mg-contratosycobranzas.github.io/revision-expedientes/' };
try {
  ['avisosConfig', 'avisosGuardar', 'avisosVista', 'avisosPrueba'].forEach(f => { if (API_PUBLICA.indexOf(f) < 0) API_PUBLICA.push(f); });
  ['avisosConfig', 'avisosVista'].forEach(f => { if (SOLO_LECTURA.indexOf(f) < 0) SOLO_LECTURA.push(f); });
} catch (e) {}

function aviAdmin_(u) { if (u.rol !== 'ADMIN') throw new Error('Solo la administración configura los avisos por correo.'); }
function aviCfg_() {
  let c = {}; try { c = JSON.parse(PropertiesService.getScriptProperties().getProperty('AVISOS_CFG') || '{}'); } catch (e) {}
  return Object.assign({}, AVI_DEF, c);
}
const aviMails_ = t => String(t || '').split(/[,;\s]+/).map(x => x.trim()).filter(x => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x));
const aviDia_ = d => Utilities.formatDate(d, 'America/Lima', 'yyyy-MM-dd');
const aviEsc_ = t => String(t == null ? '' : t).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
function aviHoja_() {
  let h = ss_().getSheetByName(HOJA_AVI);
  if (!h) {
    h = ss_().insertSheet(HOJA_AVI);
    h.getRange(1, 1, 1, AVI_HDR.length).setValues([AVI_HDR]).setFontWeight('bold').setBackground('#0C1A3D').setFontColor('#FFFFFF');
    h.setFrozenRows(1);
  }
  return h;
}
/** nombre o referencia (como en Accesos) → correo, solo usuarios habilitados */
function aviCorreos_() {
  const h = ss_().getSheetByName(HOJA_ACCESO), n = h.getLastRow(), m = {}; if (n < 2) return m;
  h.getRange(2, 1, n - 1, 6).getValues().forEach(r => {
    const c = aviMails_(r[1])[0]; if (!c || norm_(r[5]) !== 'SI') return;
    [r[0], r[3]].forEach(k => { if (String(k || '').trim()) m[norm_(k)] = c; });
  });
  return m;
}
/** avisos ya enviados (para no repetir): clave = EXP|etapa|fecha límite */
function aviHechos_() {
  const h = ss_().getSheetByName(HOJA_AVI), o = {}; if (!h || h.getLastRow() < 2) return o;
  const n = h.getLastRow(), d = Math.min(n - 1, 3000);
  h.getRange(n - d + 1, 1, d, AVI_HDR.length).getValues().forEach(r => { if (/^ENVIADO/.test(String(r[7]))) String(r[8] || '').split(',').forEach(k => { if (k) o[k] = 1; }); });
  return o;
}
/** Expedientes con puntos pendientes y su plazo (días de calendario hasta el límite). */
function aviDatos_(ahora) {
  const ss = ss_(), ct = ss.getSheetByName(HOJA_CONTROL), n = ct.getLastRow(); if (n < 5) return [];
  const est = {}, hp = ss.getSheetByName(HOJA_PUN);
  if (hp && hp.getLastRow() > 1) hp.getRange(2, 1, hp.getLastRow() - 1, 4).getValues().forEach(r => { est[String(r[0]).trim() + '#' + punClave_(r[1], r[2])] = norm_(r[3]); });
  const anu = typeof anulados_ === 'function' ? anulados_() : {}, hoy = Date.parse(aviDia_(ahora)), out = [];
  ct.getRange(5, 1, n - 4, 17).getValues().forEach(r => {
    const id = String(r[0]).trim(); if (!/^EXP/i.test(id) || anu[id]) return;
    const ex = { id: id, cliente: txt_(r[1]), proyecto: txt_(r[2]), unidad: txt_(r[3]), vendedor: txt_(r[4]), jefe: txt_(r[5]), sbs: txt_(r[6]), obsSbs: txt_(r[7]), ctr: txt_(r[8]), obsCtr: txt_(r[9]), estado: txt_(r[11]), mes: txt_(r[16]) };
    if (!ex.vendedor) return;
    const P = punPartes_(ex); if (!P.length) return;
    const pend = [], rev = [];
    P.forEach(p => { const e = est[id + '#' + punClave_(p.area, p.punto)]; if (e === 'CONFORME') return; (e === 'POR VERIFICAR' ? rev : pend).push(p); });
    if (!pend.length) return;
    const lim = plazo_(ex.mes, 'X'); if (!lim) return;
    out.push(Object.assign(ex, { pend: pend, rev: rev, lim: lim, dias: Math.round((Date.parse(aviDia_(lim)) - hoy) / 864e5) }));
  });
  return out;
}
const AVI_ET = { D2: ['d2', 'Vence en 2 días'], D1: ['d1', 'Vence mañana'], D0: ['d0', 'Vence HOY'], VENCIDO: ['vencido', 'Venció ayer'] };
const aviEtapa_ = (x, c) => { const e = x.dias === 2 ? 'D2' : x.dias === 1 ? 'D1' : x.dias === 0 ? 'D0' : x.dias === -1 ? 'VENCIDO' : null; return e && c[AVI_ET[e][0]] ? e : null; };
/** correo de un vendedor con sus expedientes */
function aviCorreo_(ven, l, c, sinCorreo) {
  const orden = { D0: 0, D1: 1, VENCIDO: 2, D2: 3 }; l.sort((a, b) => orden[a.etapa] - orden[b.etapa] || a.lim - b.lim);
  const urg = l[0], n = l.length, lim = fmtPlazo_(urg.lim);
  const asunto = (urg.etapa === 'VENCIDO' ? '⚠️ Venció el plazo para subsanar' : '⏰ ' + AVI_ET[urg.etapa][1] + ' (' + lim + ')') + ' · ' + n + ' expediente' + (n > 1 ? 's' : '') + ' por subsanar · ' + ven;
  const li = a => a.map(p => '<li>' + aviEsc_(p.area) + ' · ' + aviEsc_(p.punto) + '</li>').join('');
  const fila = x => '<tr><td style="padding:8px;border:1px solid #D9DEE8;vertical-align:top"><b>' + aviEsc_(x.cliente) + '</b><br>' + aviEsc_(x.proyecto + ' ' + x.unidad) + '<br><small style="color:#667">' + aviEsc_(x.id) + '</small></td>'
    + '<td style="padding:8px;border:1px solid #D9DEE8;vertical-align:top;white-space:nowrap;color:' + (x.dias <= 0 ? '#B42318' : '#B54708') + '"><b>' + aviEsc_(AVI_ET[x.etapa][1]) + '</b><br>' + aviEsc_(fmtPlazo_(x.lim)) + '</td>'
    + '<td style="padding:8px;border:1px solid #D9DEE8;vertical-align:top"><b style="color:#B42318">🔴 Falta subsanar (' + x.pend.length + ')</b><ul style="margin:4px 0 0 16px;padding:0">' + li(x.pend) + '</ul>'
    + (x.rev.length ? '<b style="color:#B54708">🕓 En revisión de Contratos (' + x.rev.length + ')</b><ul style="margin:4px 0 0 16px;padding:0">' + li(x.rev) + '</ul>' : '') + '</td></tr>';
  const html = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#1F2937;max-width:760px">'
    + (sinCorreo ? '<p style="background:#FFF6D6;padding:8px;border-radius:6px">⚠️ ' + aviEsc_(ven) + ' no tiene correo registrado en Accesos: este aviso solo llega en copia.</p>' : '')
    + '<p>Hola ' + aviEsc_(String(ven).split(' ')[0]) + ',</p><p>Tienes <b>' + n + ' expediente' + (n > 1 ? 's' : '') + '</b> con observaciones por subsanar. El plazo es el del calendario de cierre (<i>Comercial presenta subsanados</i>):</p>'
    + '<table style="border-collapse:collapse;width:100%"><tr style="background:#0C1A3D;color:#fff"><th style="padding:8px;text-align:left">Expediente</th><th style="padding:8px;text-align:left">Plazo</th><th style="padding:8px;text-align:left">Qué falta</th></tr>' + l.map(fila).join('') + '</table>'
    + '<p>Entra al <a href="' + aviEsc_(c.portal) + '">Cuadro de Revisiones</a>, marca en cada expediente los puntos que ya subsanaste y adjunta el sustento.' + (l.some(x => x.etapa === 'VENCIDO') ? ' Si el plazo ya venció, pide habilitación desde el expediente.' : '') + '</p>'
    + '<p style="color:#667;font-size:12px">Aviso automático de Contratos y Cobranzas · Grupo MG.</p></div>';
  const texto = 'Hola ' + String(ven).split(' ')[0] + ',\n\nTienes ' + n + ' expediente(s) por subsanar:\n\n' + l.map(x => '• ' + x.cliente + ' · ' + x.proyecto + ' ' + x.unidad + ' (' + x.id + ') — ' + AVI_ET[x.etapa][1] + ' (' + fmtPlazo_(x.lim) + ')\n  Falta: ' + x.pend.map(p => p.punto).join(' / ') + (x.rev.length ? '\n  En revisión: ' + x.rev.map(p => p.punto).join(' / ') : '')).join('\n\n') + '\n\nCuadro de Revisiones: ' + c.portal;
  return { asunto: asunto, html: html, texto: texto };
}
/** resumen para la administración */
function aviResumen_(D, c, ahora) {
  const l = D.filter(x => x.dias >= -3 && x.dias <= 2).sort((a, b) => a.dias - b.dias || String(a.vendedor).localeCompare(b.vendedor)); if (!l.length) return null;
  const g = { v: l.filter(x => x.dias < 0), h: l.filter(x => x.dias === 0), p: l.filter(x => x.dias > 0) };
  const t = (tit, a, col) => a.length ? '<h3 style="color:' + col + ';margin:16px 0 6px">' + tit + ' · ' + a.length + '</h3><table style="border-collapse:collapse;width:100%">' + a.map(x => '<tr><td style="padding:6px;border:1px solid #D9DEE8"><b>' + aviEsc_(x.vendedor) + '</b><br><small>' + aviEsc_(x.jefe) + '</small></td><td style="padding:6px;border:1px solid #D9DEE8">' + aviEsc_(x.cliente) + '<br><small>' + aviEsc_(x.proyecto + ' ' + x.unidad + ' · ' + x.id) + '</small></td><td style="padding:6px;border:1px solid #D9DEE8;white-space:nowrap">' + aviEsc_(fmtPlazo_(x.lim)) + '</td><td style="padding:6px;border:1px solid #D9DEE8">🔴 ' + x.pend.length + (x.rev.length ? ' · 🕓 ' + x.rev.length : '') + '</td></tr>').join('') + '</table>' : '';
  return { asunto: '📋 Resumen de plazos de subsanación · ' + Utilities.formatDate(ahora, 'America/Lima', 'dd/MM/yyyy') + ' · ' + g.h.length + ' vencen hoy · ' + g.v.length + ' vencidos',
    html: '<div style="font-family:Arial,sans-serif;font-size:14px;color:#1F2937;max-width:820px"><p>Expedientes con puntos por subsanar (🔴 pendientes · 🕓 en revisión):</p>' + t('⚠️ Vencidos (últimos 3 días)', g.v, '#B42318') + t('⏰ Vencen hoy', g.h, '#B54708') + t('📅 Vencen en 1 o 2 días', g.p, '#1F4E79') + '<p><a href="' + aviEsc_(c.portal) + '">Abrir el Cuadro de Revisiones</a></p></div>',
    texto: 'Vencidos: ' + g.v.length + ' · Vencen hoy: ' + g.h.length + ' · Próximos: ' + g.p.length + '\n' + c.portal, exps: l.map(x => x.id) };
}
/** arma (y si enviar=true, envía) los avisos de hoy */
function aviArmar_(c, enviar) {
  const ahora = new Date(), D = aviDatos_(ahora), hechos = aviHechos_(), mails = aviCorreos_(), porVen = {}, envios = [];
  D.forEach(x => { const et = aviEtapa_(x, c); if (!et) return; x.etapa = et; x.clave = x.id + '|' + et + '|' + aviDia_(x.lim); if (hechos[x.clave]) return; (porVen[x.vendedor] = porVen[x.vendedor] || []).push(x); });
  const fijos = aviMails_(c.copia);
  Object.keys(porVen).forEach(v => {
    const l = porVen[v], para = mails[norm_(v)] || '', jefe = c.copiaJefe ? (mails[norm_(l[0].jefe)] || '') : '';
    const cc = [...new Set([jefe].concat(fijos).filter(Boolean))].filter(m => m !== para), m = aviCorreo_(v, l, c, !para);
    envios.push({ tipo: 'RECORDATORIO', vendedor: v, para: para || cc.join(','), cc: para ? cc : [], sinCorreo: !para, asunto: m.asunto, html: m.html, texto: m.texto, exps: l.map(x => x.id), claves: l.map(x => x.clave) });
  });
  if (c.resumen && fijos.length && !hechos['RESUMEN|' + aviDia_(ahora)]) {
    const r = aviResumen_(D, c, ahora);
    if (r) envios.push({ tipo: 'RESUMEN', vendedor: '', para: fijos.join(','), cc: [], asunto: r.asunto, html: r.html, texto: r.texto, exps: r.exps, claves: ['RESUMEN|' + aviDia_(ahora)] });
  }
  if (!enviar) return envios;
  const h = aviHoja_(), filas = [];
  envios.forEach(e => {
    let res = 'ENVIADO';
    try {
      if (!e.para) res = 'SIN CORREO';
      else if (MailApp.getRemainingDailyQuota() < 1 + e.cc.length) res = 'SIN CUOTA (límite diario de Gmail)';
      else MailApp.sendEmail({ to: e.para, cc: e.cc.join(','), subject: e.asunto, htmlBody: e.html, body: e.texto, name: 'Contratos y Cobranzas MG' });
    } catch (err) { res = 'ERROR: ' + err.message; }
    filas.push([new Date(), e.tipo, e.vendedor, e.para, e.cc.join(', '), e.exps.join(', '), e.asunto, res, e.claves.join(',')]);
  });
  if (filas.length) h.getRange(h.getLastRow() + 1, 1, filas.length, AVI_HDR.length).setValues(filas);
  return envios;
}
/** lo que ejecuta la tarea programada (una vez al día) */
function avisosTick() { const c = aviCfg_(); if (!c.activo) return; aviArmar_(c, true); }

/** Administración: configuración, estado de la tarea diaria, vendedores sin correo y últimos envíos. */
function avisosConfig(token) {
  const u = sesion_(token); aviAdmin_(u);
  const c = aviCfg_(), tarea = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'avisosTick'), mails = aviCorreos_();
  const D = aviDatos_(new Date()), sin = [...new Set(D.map(x => x.vendedor).filter(v => !mails[norm_(v)]))].sort();
  const h = ss_().getSheetByName(HOJA_AVI), n = h ? h.getLastRow() : 0, k = Math.min(Math.max(n - 1, 0), 20);
  const ult = k ? h.getRange(n - k + 1, 1, k, 8).getValues().reverse().map(r => ({ fecha: r[0] instanceof Date ? Utilities.formatDate(r[0], 'America/Lima', 'dd/MM/yyyy HH:mm') : txt_(r[0]), tipo: txt_(r[1]), vendedor: txt_(r[2]), para: txt_(r[3]), asunto: txt_(r[6]), resultado: txt_(r[7]) })) : [];
  let cuota = null; try { cuota = MailApp.getRemainingDailyQuota(); } catch (e) {}
  return { cfg: c, tarea: tarea, sinCorreo: sin, pendientes: D.length, ultimos: ult, cuota: cuota };
}
/** Administración: guarda la configuración y crea o quita la tarea diaria. */
function avisosGuardar(token, d) {
  const u = sesion_(token); aviAdmin_(u); d = d || {};
  const hora = Math.max(6, Math.min(20, Math.round(Number(d.hora) || 8))), copia = aviMails_(d.copia).join(', ');
  const c = { activo: !!d.activo, hora: hora, d2: !!d.d2, d1: !!d.d1, d0: !!d.d0, vencido: !!d.vencido, resumen: !!d.resumen, copiaJefe: !!d.copiaJefe, copia: copia, portal: String(d.portal || AVI_DEF.portal).trim() };
  if (!/^https:\/\//.test(c.portal)) c.portal = AVI_DEF.portal;
  PropertiesService.getScriptProperties().setProperty('AVISOS_CFG', JSON.stringify(c));
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'avisosTick').forEach(t => ScriptApp.deleteTrigger(t));
  if (c.activo) ScriptApp.newTrigger('avisosTick').timeBased().everyDays(1).atHour(hora).inTimezone('America/Lima').create();
  log_(u.nombre, 'AVISOS_CORREO', (c.activo ? 'activados ' + hora + ':00' : 'desactivados') + ' · copia ' + (copia || '—'));
  return avisosConfig(token);
}
/** Administración: qué correos saldrían hoy (sin enviar nada). */
function avisosVista(token) {
  const u = sesion_(token); aviAdmin_(u);
  return aviArmar_(aviCfg_(), false).map(e => ({ tipo: e.tipo, vendedor: e.vendedor, para: e.para, cc: e.cc, sinCorreo: !!e.sinCorreo, asunto: e.asunto, exps: e.exps, html: e.html }));
}
/** Administración: envía UNA muestra solo a quien la pide (no cuenta como aviso enviado). */
function avisosPrueba(token) {
  const u = sesion_(token); aviAdmin_(u);
  const yo = aviMails_(u.correo)[0] || aviMails_(aviCfg_().copia)[0]; if (!yo) throw new Error('Tu usuario no tiene correo en Accesos.');
  const e = aviArmar_(aviCfg_(), false)[0];
  if (!e) { MailApp.sendEmail({ to: yo, subject: '[PRUEBA] Avisos por correo · Cuadro de Revisiones', body: 'Los avisos funcionan. Hoy no hay expedientes que venzan en los próximos 2 días.', name: 'Contratos y Cobranzas MG' }); return { para: yo, asunto: '(sin expedientes por vencer hoy)' }; }
  MailApp.sendEmail({ to: yo, subject: '[PRUEBA] ' + e.asunto, htmlBody: '<p style="background:#EEF1F7;padding:8px;font-family:Arial;font-size:13px">Prueba: este correo saldría a <b>' + aviEsc_(e.para || '(sin correo)') + '</b>' + (e.cc.length ? ' con copia a <b>' + aviEsc_(e.cc.join(', ')) + '</b>' : '') + '.</p>' + e.html, body: e.texto, name: 'Contratos y Cobranzas MG' });
  return { para: yo, asunto: e.asunto };
}
