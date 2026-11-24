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
const AVI_DEF = { activo: false, hora: 8, d2: true, d1: true, d0: true, vencido: true, resumen: true, copiaJefe: true, devuelto: false, habilitado: false, espera: false, jefes: 'NO', jefesCopia: false,
  copia: 'mvalles@grupomg.pe', portal: 'https://mg-contratosycobranzas.github.io/revision-expedientes/' };
try {
  ['avisosConfig', 'avisosGuardar', 'avisosVista', 'avisosPrueba'].forEach(f => { if (API_PUBLICA.indexOf(f) < 0) API_PUBLICA.push(f); });
  ['avisosConfig', 'avisosVista'].forEach(f => { if (SOLO_LECTURA.indexOf(f) < 0) SOLO_LECTURA.push(f); });
} catch (e) {}

/* ---------- identificación completa de un expediente en los correos ----------
 * "EXP-2026-171 · CLIENTE · DNI 12345678 · PROYECTO 803". El DNI del titular sale de BASE MAESTRA (columna I);
 * se lee una vez por ejecución. También la puede usar Code.gs (subsanar) si existe: typeof expRef_ === 'function'. */
let AVI_BM_ = null;
function aviBM_() {
  if (AVI_BM_) return AVI_BM_; AVI_BM_ = {};
  try { ss_().getSheetByName(HOJA_BM).getRange(PRIMERA_FILA, 1, ULTIMA_FILA - PRIMERA_FILA + 1, 9).getDisplayValues().forEach(r => {
    const id = String(r[0]).trim(); if (id) AVI_BM_[id] = { cliente: r[7].trim(), dni: r[8].trim(), proyecto: r[4].trim(), unidad: r[6].trim() }; }); } catch (e) {}
  return AVI_BM_;
}
function aviDni_(id) { const d = (aviBM_()[String(id || '').trim()] || {}).dni || ''; return d ? 'DNI ' + d : ''; }
function expRef_(x) {
  const o = (x && typeof x === 'object') ? x : {}, id = String(o.id || x || '').trim(), b = aviBM_()[id] || {};
  const pu = ((o.proyecto || b.proyecto || '') + ' ' + (o.unidad || b.unidad || '')).trim();
  return [id, o.cliente || b.cliente || '', aviDni_(id), pu].filter(Boolean).join(' · ');
}
/** Segunda línea de la ficha en los correos: "PROYECTO 803 · DNI 12345678 · EXP-2026-171". */
function aviSub_(x) { return [((x.proyecto || '') + ' ' + (x.unidad || '')).trim(), aviDni_(x.id), x.id].filter(Boolean).join(' · '); }

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
/** Todas las ventas de CONTROL DE REVISIONES con sus puntos 🔴 pendientes / 🕓 en revisión y su plazo. */
function aviBase_(ahora) {
  const ss = ss_(), ct = ss.getSheetByName(HOJA_CONTROL), n = ct.getLastRow(); if (n < 5) return [];
  const est = {}, hp = ss.getSheetByName(HOJA_PUN);
  if (hp && hp.getLastRow() > 1) hp.getRange(2, 1, hp.getLastRow() - 1, 4).getValues().forEach(r => { est[String(r[0]).trim() + '#' + punClave_(r[1], r[2])] = norm_(r[3]); });
  const anu = typeof anulados_ === 'function' ? anulados_() : {}, hoy = Date.parse(aviDia_(ahora)), out = [];
  ct.getRange(5, 1, n - 4, 17).getValues().forEach(r => {
    const id = String(r[0]).trim(); if (!/^EXP/i.test(id) || anu[id]) return;
    const ex = { id: id, cliente: txt_(r[1]), proyecto: txt_(r[2]), unidad: txt_(r[3]), vendedor: txt_(r[4]), jefe: txt_(r[5]), sbs: txt_(r[6]), obsSbs: txt_(r[7]), ctr: txt_(r[8]), obsCtr: txt_(r[9]), revisor: txt_(r[10]), estado: txt_(r[11]),
      mes: typeof mesTxt_ === 'function' ? mesTxt_(r[16]) : txt_(r[16]) };
    const pend = [], rev = [];
    punPartes_(ex).forEach(p => { const e = est[id + '#' + punClave_(p.area, p.punto)]; if (e === 'CONFORME') return; (e === 'POR VERIFICAR' ? rev : pend).push(p); });
    const lim = pend.length ? plazo_(ex.mes, 'X') : null;
    out.push(Object.assign(ex, { pend: pend, rev: rev, lim: lim, dias: lim ? Math.round((Date.parse(aviDia_(lim)) - hoy) / 864e5) : null }));
  });
  return out;
}
/** Expedientes con puntos pendientes y su plazo (días de calendario hasta el límite). */
function aviDatos_(ahora) { return aviBase_(ahora).filter(x => x.vendedor && x.pend.length && x.lim); }
const AVI_ET = { D2: ['d2', 'Vence en 2 días'], D1: ['d1', 'Vence mañana'], D0: ['d0', 'Vence HOY'], VENCIDO: ['vencido', 'Venció ayer'] };
const aviEtapa_ = (x, c) => { const e = x.dias === 2 ? 'D2' : x.dias === 1 ? 'D1' : x.dias === 0 ? 'D0' : x.dias === -1 ? 'VENCIDO' : null; return e && c[AVI_ET[e][0]] ? e : null; };
/** correo de un vendedor con sus expedientes */
function aviCorreo_(ven, l, c, sinCorreo) {
  const orden = { D0: 0, D1: 1, VENCIDO: 2, D2: 3 }; l.sort((a, b) => orden[a.etapa] - orden[b.etapa] || a.lim - b.lim);
  const urg = l[0], n = l.length, lim = fmtPlazo_(urg.lim);
  const asunto = (urg.etapa === 'VENCIDO' ? '⚠️ Venció el plazo para subsanar' : '⏰ ' + AVI_ET[urg.etapa][1] + ' (' + lim + ')') + ' · ' + n + ' expediente' + (n > 1 ? 's' : '') + ' por subsanar · ' + ven;
  const li = a => a.map(p => '<li>' + aviEsc_(p.area) + ' · ' + aviEsc_(p.punto) + '</li>').join('');
  const fila = x => '<tr><td style="padding:8px;border:1px solid #D9DEE8;vertical-align:top"><b>' + aviEsc_(x.cliente) + '</b><br>' + aviEsc_(x.proyecto + ' ' + x.unidad) + '<br><small style="color:#667">' + aviEsc_([aviDni_(x.id), x.id].filter(Boolean).join(' · ')) + '</small></td>'
    + '<td style="padding:8px;border:1px solid #D9DEE8;vertical-align:top;white-space:nowrap;color:' + (x.dias <= 0 ? '#B42318' : '#B54708') + '"><b>' + aviEsc_(AVI_ET[x.etapa][1]) + '</b><br>' + aviEsc_(fmtPlazo_(x.lim)) + '</td>'
    + '<td style="padding:8px;border:1px solid #D9DEE8;vertical-align:top"><b style="color:#B42318">🔴 Falta subsanar (' + x.pend.length + ')</b><ul style="margin:4px 0 0 16px;padding:0">' + li(x.pend) + '</ul>'
    + (x.rev.length ? '<b style="color:#B54708">🕓 En revisión de Contratos (' + x.rev.length + ')</b><ul style="margin:4px 0 0 16px;padding:0">' + li(x.rev) + '</ul>' : '') + '</td></tr>';
  const html = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#1F2937;max-width:760px">'
    + (sinCorreo ? '<p style="background:#FFF6D6;padding:8px;border-radius:6px">⚠️ ' + aviEsc_(ven) + ' no tiene correo registrado en Accesos: este aviso solo llega en copia.</p>' : '')
    + '<p>Hola ' + aviEsc_(String(ven).split(' ')[0]) + ',</p><p>Tienes <b>' + n + ' expediente' + (n > 1 ? 's' : '') + '</b> con observaciones por subsanar. El plazo es el del calendario de cierre (<i>Comercial presenta subsanados</i>):</p>'
    + '<table style="border-collapse:collapse;width:100%"><tr style="background:#0C1A3D;color:#fff"><th style="padding:8px;text-align:left">Expediente</th><th style="padding:8px;text-align:left">Plazo</th><th style="padding:8px;text-align:left">Qué falta</th></tr>' + l.map(fila).join('') + '</table>'
    + '<p>Entra al <a href="' + aviEsc_(c.portal) + '">Cuadro de Revisiones</a>, marca en cada expediente los puntos que ya subsanaste y adjunta el sustento.' + (l.some(x => x.etapa === 'VENCIDO') ? ' Si el plazo ya venció, pide habilitación desde el expediente.' : '') + '</p>'
    + '<p style="color:#667;font-size:12px">Aviso automático de Contratos y Cobranzas · Grupo MG.</p></div>';
  const texto = 'Hola ' + String(ven).split(' ')[0] + ',\n\nTienes ' + n + ' expediente(s) por subsanar:\n\n' + l.map(x => '• ' + expRef_(x) + ' — ' + AVI_ET[x.etapa][1] + ' (' + fmtPlazo_(x.lim) + ')\n  Falta: ' + x.pend.map(p => p.punto).join(' / ') + (x.rev.length ? '\n  En revisión: ' + x.rev.map(p => p.punto).join(' / ') : '')).join('\n\n') + '\n\nCuadro de Revisiones: ' + c.portal;
  return { asunto: asunto, html: html, texto: texto };
}
/** resumen para la administración */
function aviResumen_(D, c, ahora) {
  const l = D.filter(x => x.dias >= -3 && x.dias <= 2).sort((a, b) => a.dias - b.dias || String(a.vendedor).localeCompare(b.vendedor)); if (!l.length) return null;
  const g = { v: l.filter(x => x.dias < 0), h: l.filter(x => x.dias === 0), p: l.filter(x => x.dias > 0) };
  const t = (tit, a, col) => a.length ? '<h3 style="color:' + col + ';margin:16px 0 6px">' + tit + ' · ' + a.length + '</h3><table style="border-collapse:collapse;width:100%">' + a.map(x => '<tr><td style="padding:6px;border:1px solid #D9DEE8"><b>' + aviEsc_(x.vendedor) + '</b><br><small>' + aviEsc_(x.jefe) + '</small></td><td style="padding:6px;border:1px solid #D9DEE8">' + aviEsc_(x.cliente) + '<br><small>' + aviEsc_(aviSub_(x)) + '</small></td><td style="padding:6px;border:1px solid #D9DEE8;white-space:nowrap">' + aviEsc_(fmtPlazo_(x.lim)) + '</td><td style="padding:6px;border:1px solid #D9DEE8">🔴 ' + x.pend.length + (x.rev.length ? ' · 🕓 ' + x.rev.length : '') + '</td></tr>').join('') + '</table>' : '';
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
function avisosTick() {
  const c = aviCfg_();
  if (c.activo) aviArmar_(c, true);
  if (c.jefes !== 'NO') try { aviJefes_(c, true); } catch (e) { console.log('aviJefes_: ' + e.message); }
  if (c.espera) try { aviEspera_(c); } catch (e) { console.log('aviEspera_: ' + e.message); }
}

/* ---------- avisos INMEDIATOS (salen en el momento; cada uno con su interruptor en el portal) ---------- */
function aviMandar_(tipo, vendedor, para, cc, asunto, html, texto, exps, clave) {
  para = aviMails_(para).join(','); cc = [...new Set(aviMails_(cc.join(',')))].filter(m => para.split(',').indexOf(m) < 0);
  if (!para && cc.length) { para = cc.join(','); cc = []; }
  let res = 'ENVIADO';
  try { if (!para) res = 'SIN CORREO'; else MailApp.sendEmail({ to: para, cc: cc.join(','), subject: asunto, htmlBody: html, body: texto, name: 'Contratos y Cobranzas MG' }); }
  catch (e) { res = 'ERROR: ' + e.message; }
  try { aviHoja_().appendRow([new Date(), tipo, vendedor || '', para, cc.join(', '), (exps || []).join(', '), asunto, res, clave || '']); } catch (e) {}
  return res;
}
const aviCaja_ = (c, cuerpo) => '<div style="font-family:Arial,sans-serif;font-size:14px;color:#1F2937;max-width:720px">' + cuerpo + '<p>Entra al <a href="' + aviEsc_(c.portal) + '">Cuadro de Revisiones</a> para ver el detalle.</p><p style="color:#667;font-size:12px">Aviso automático de Contratos y Cobranzas · Grupo MG.</p></div>';
/** 🔴 Contratos devolvió puntos ("sigue observado") → al vendedor (y su jefe si está activa esa copia). */
function aviDevuelto_(ex, obs, comentario, quien) {
  const c = aviCfg_(); if (!c.devuelto || !obs || !obs.length) return;
  const m = aviCorreos_(), para = m[norm_(ex.vendedor)] || '', jefe = c.copiaJefe ? (m[norm_(ex.jefe)] || '') : '';
  const lis = obs.map(p => '<li><b>' + aviEsc_(p.punto) + '</b>' + (p.coment ? '<br><span style="color:#B42318">Motivo: ' + aviEsc_(p.coment) + '</span>' : '') + '</li>').join('');
  const html = aviCaja_(c, '<p>Hola ' + aviEsc_(String(ex.vendedor).split(' ')[0]) + ',</p><p>' + aviEsc_(quien) + ' revisó tu subsanación de <b>' + aviEsc_(ex.cliente) + '</b> (' + aviEsc_(aviSub_(ex)) + ') y <b style="color:#B42318">' + (obs.length > 1 ? 'estos puntos siguen observados' : 'este punto sigue observado') + '</b>:</p><ul>' + lis + '</ul>' + (comentario ? '<p>💬 ' + aviEsc_(comentario) + '</p>' : ''));
  aviMandar_('DEVUELTO', ex.vendedor, para, [jefe], '🔴 Sigue observado: ' + expRef_(ex) + ' (' + obs.length + ' punto' + (obs.length > 1 ? 's' : '') + ')', html,
    expRef_(ex) + ' sigue observado:\n' + obs.map(p => '• ' + p.punto + (p.coment ? ' — ' + p.coment : '')).join('\n') + (comentario ? '\n\n' + comentario : '') + '\n\n' + c.portal, [ex.id], 'DEVUELTO|' + ex.id + '|' + new Date().getTime());
}
/** 🔓 Habilitación de un expediente o excepción de plazo → a la persona (y al vendedor del expediente). */
function aviHabilitado_(personas, asunto, detalle, exps) {
  const c = aviCfg_(); if (!c.habilitado) return;
  const m = aviCorreos_(), para = [...new Set(personas.map(p => m[norm_(p)]).filter(Boolean))];
  if (!para.length) return;
  aviMandar_('HABILITACION', personas[0], para.join(','), [], asunto, aviCaja_(c, '<p>' + detalle + '</p>'), detalle.replace(/<[^>]+>/g, '') + '\n\n' + c.portal, exps || [], '');
}
/** 🕓 Subsanaciones de Comercial que esperan la revisión de Contratos hace más de 2 días → a la copia fija (en la tarea diaria). */
function aviEspera_(c) {
  const fijos = aviMails_(c.copia); if (!fijos.length) return;
  const ahora = new Date(), clave = 'ESPERA|' + aviDia_(ahora); if (aviHechos_()[clave]) return;
  const sh = ss_().getSheetByName(HOJA_SUB), n = sh ? sh.getLastRow() : 0; if (n < 2) return;
  const info = {}, ct = ss_().getSheetByName(HOJA_CONTROL), nc = ct.getLastRow();
  if (nc >= 5) ct.getRange(5, 1, nc - 4, 11).getValues().forEach(r => { info[String(r[0]).trim()] = { cliente: txt_(r[1]), pu: txt_(r[2]) + ' ' + txt_(r[3]), vendedor: txt_(r[4]), revisor: txt_(r[10]) }; });
  const l = [], vis = {};
  sh.getRange(2, 1, n - 1, 8).getValues().forEach(r => {
    const id = String(r[1]).trim(); if (!(r[0] instanceof Date) || !/Pendiente/i.test(String(r[7])) || vis[id]) return;
    const h = (ahora - r[0]) / 36e5; if (h < 48) return; vis[id] = 1;
    l.push({ id: id, dias: Math.floor(h / 24), quien: String(r[4]).replace(/\s*\(Ventas\)\s*$/, ''), areas: txt_(r[5]), i: info[id] || {} });
  });
  if (!l.length) return;
  l.sort((a, b) => b.dias - a.dias);
  const html = aviCaja_(c, '<p><b>' + l.length + ' subsanación' + (l.length > 1 ? 'es esperan' : ' espera') + ' la revisión de Contratos hace más de 2 días:</b></p><table style="border-collapse:collapse;width:100%"><tr style="background:#0C1A3D;color:#fff"><th style="padding:6px;text-align:left">Expediente</th><th style="padding:6px;text-align:left">Subsanó</th><th style="padding:6px;text-align:left">Área · revisor</th><th style="padding:6px">Días</th></tr>'
    + l.map(x => '<tr><td style="padding:6px;border:1px solid #D9DEE8"><b>' + aviEsc_(x.i.cliente || x.id) + '</b><br><small>' + aviEsc_([x.i.pu || '', aviDni_(x.id), x.id].filter(Boolean).join(' · ')) + '</small></td><td style="padding:6px;border:1px solid #D9DEE8">' + aviEsc_(x.quien) + '</td><td style="padding:6px;border:1px solid #D9DEE8">' + aviEsc_(x.areas) + (x.i.revisor ? '<br><small>' + aviEsc_(x.i.revisor) + '</small>' : '') + '</td><td style="padding:6px;border:1px solid #D9DEE8;text-align:center;color:' + (x.dias >= 4 ? '#B42318' : '#B54708') + '"><b>' + x.dias + '</b></td></tr>').join('') + '</table>');
  aviMandar_('ESPERA', '', fijos.join(','), [], '🕓 ' + l.length + ' subsanación' + (l.length > 1 ? 'es' : '') + ' esperando revisión hace más de 2 días', html, l.map(x => '• ' + expRef_(x.i && x.i.cliente ? { id: x.id, cliente: x.i.cliente } : x.id) + ' · ' + x.dias + ' días · ' + x.quien).join('\n'), l.map(x => x.id), clave);
}

/** Administración: configuración, estado de la tarea diaria, vendedores sin correo y últimos envíos. */
function avisosConfig(token) {
  const u = sesion_(token); aviAdmin_(u);
  const c = aviCfg_(), tarea = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'avisosTick'), mails = aviCorreos_();
  const D = aviDatos_(new Date()), nombres = [...new Set(D.map(x => x.vendedor).concat(c.jefes !== 'NO' ? D.map(x => x.jefe) : []).filter(v => v && !mails[norm_(v)]))].sort();
  const ha = ss_().getSheetByName(HOJA_ACCESO), na = ha.getLastRow(), acc = na >= 2 ? ha.getRange(2, 1, na - 1, 4).getValues() : [];
  const sin = nombres.map(v => { const i = acc.findIndex(r => norm_(r[0]) === norm_(v) || norm_(r[3]) === norm_(v)); return { nombre: v, fila: i < 0 ? null : i + 2, usuario: i < 0 ? '' : String(acc[i][0]), rol: i < 0 ? '' : norm_(acc[i][2]) }; });
  const h = ss_().getSheetByName(HOJA_AVI), n = h ? h.getLastRow() : 0, k = Math.min(Math.max(n - 1, 0), 20);
  const ult = k ? h.getRange(n - k + 1, 1, k, 8).getValues().reverse().map(r => ({ fecha: r[0] instanceof Date ? Utilities.formatDate(r[0], 'America/Lima', 'dd/MM/yyyy HH:mm') : txt_(r[0]), tipo: txt_(r[1]), vendedor: txt_(r[2]), para: txt_(r[3]), asunto: txt_(r[6]), resultado: txt_(r[7]) })) : [];
  let cuota = null; try { cuota = MailApp.getRemainingDailyQuota(); } catch (e) {}
  return { cfg: c, tarea: tarea, sinCorreo: sin, pendientes: D.length, ultimos: ult, cuota: cuota };
}
/** Administración: guarda la configuración y crea o quita la tarea diaria. */
function avisosGuardar(token, d) {
  const u = sesion_(token); aviAdmin_(u); d = d || {};
  const hora = Math.max(6, Math.min(20, Math.round(Number(d.hora) || 8))), copia = aviMails_(d.copia).join(', ');
  const c = { activo: !!d.activo, hora: hora, d2: !!d.d2, d1: !!d.d1, d0: !!d.d0, vencido: !!d.vencido, resumen: !!d.resumen, copiaJefe: !!d.copiaJefe, copia: copia, portal: String(d.portal || AVI_DEF.portal).trim(),
    devuelto: !!d.devuelto, habilitado: !!d.habilitado, espera: !!d.espera, jefes: ['DIARIO', 'LUNES'].indexOf(d.jefes) >= 0 ? d.jefes : 'NO', jefesCopia: !!d.jefesCopia };
  if (!/^https:\/\//.test(c.portal)) c.portal = AVI_DEF.portal;
  PropertiesService.getScriptProperties().setProperty('AVISOS_CFG', JSON.stringify(c));
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'avisosTick').forEach(t => ScriptApp.deleteTrigger(t));
  if (c.activo || c.espera || c.jefes !== 'NO') ScriptApp.newTrigger('avisosTick').timeBased().everyDays(1).atHour(hora).inTimezone('America/Lima').create();
  log_(u.nombre, 'AVISOS_CORREO', (c.activo ? 'activados ' + hora + ':00' : 'desactivados') + ' · copia ' + (copia || '—'));
  return avisosConfig(token);
}
/** Administración: qué correos saldrían hoy (sin enviar nada). */
function avisosVista(token) {
  const u = sesion_(token); aviAdmin_(u);
  const c = aviCfg_(), l = aviArmar_(c, false).concat(c.jefes !== 'NO' ? aviJefes_(c, false) : []);
  return l.map(e => ({ tipo: e.tipo, vendedor: e.vendedor, para: e.para, cc: e.cc, sinCorreo: !!e.sinCorreo, asunto: e.asunto, exps: e.exps, html: e.html }));
}
/** Administración: envía UNA muestra solo a quien la pide (no cuenta como aviso enviado). */
function avisosPrueba(token, tipo) {
  const u = sesion_(token); aviAdmin_(u);
  const yo = aviMails_(u.correo)[0] || aviMails_(aviCfg_().copia)[0]; if (!yo) throw new Error('Tu usuario no tiene correo en Accesos.');
  if (tipo === 'JEFES') { const r = aviJefes_(aviCfg_(), true, yo)[0]; if (!r) throw new Error('No hay jefes con ventas para resumir.'); if (!/^ENVIADO/.test(r.resultado)) throw new Error(r.resultado); return { para: yo, asunto: r.asunto }; }
  const e = aviArmar_(aviCfg_(), false)[0];
  if (!e) { MailApp.sendEmail({ to: yo, subject: '[PRUEBA] Avisos por correo · Cuadro de Revisiones', body: 'Los avisos funcionan. Hoy no hay expedientes que venzan en los próximos 2 días.', name: 'Contratos y Cobranzas MG' }); return { para: yo, asunto: '(sin expedientes por vencer hoy)' }; }
  MailApp.sendEmail({ to: yo, subject: '[PRUEBA] ' + e.asunto, htmlBody: '<p style="background:#EEF1F7;padding:8px;font-family:Arial;font-size:13px">Prueba: este correo saldría a <b>' + aviEsc_(e.para || '(sin correo)') + '</b>' + (e.cc.length ? ' con copia a <b>' + aviEsc_(e.cc.join(', ')) + '</b>' : '') + '.</p>' + e.html, body: e.texto, name: 'Contratos y Cobranzas MG' });
  return { para: yo, asunto: e.asunto };
}

/* ---------- 📊 RESUMEN PARA JEFES DE VENTAS: sus ventas observadas + Excel adjunto del cierre ---------- */
const AVI_MESES = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];
const aviMesK_ = m => { const p = norm_(m).replace('SETIEMBRE', 'SEPTIEMBRE').split(' '), i = AVI_MESES.indexOf(p[0]); return i < 0 || !p[1] ? -1 : Number(p[1]) * 12 + i; };
const aviCap_ = t => String(t || '').toLowerCase().replace(/(^|\s)\S/g, x => x.toUpperCase());
const AVI_XLS_HDR = ['EXP', 'Mes de cierre', 'Cliente', 'Proyecto', 'Unidad', 'Vendedor', 'Estado SBS', 'Observación SBS', 'Estado contractual', 'Observación contractual', 'Revisor', 'Estado general', 'Puntos por subsanar', 'Puntos en revisión', 'Plazo para subsanar'];
const aviXlsFila_ = x => [x.id, aviCap_(x.mes), x.cliente, x.proyecto, x.unidad, x.vendedor, x.sbs, x.obsSbs, x.ctr, x.obsCtr, x.revisor, x.estado, x.pend.length || '', x.rev.length || '', x.lim ? fmtPlazo_(x.lim) : ''];
/** Crea un .xlsx (hoja temporal de Google que se convierte y se envía a la papelera). hojas = [[nombre, filas]] */
function aviXlsx_(nombre, hojas) {
  const ss = SpreadsheetApp.create(nombre), anchos = [110, 100, 220, 100, 70, 150, 90, 260, 90, 320, 120, 100, 80, 80, 140];
  try {
    hojas.forEach((h, i) => {
      const sh = i ? ss.insertSheet(h[0]) : ss.getSheets()[0].setName(h[0]), f = [AVI_XLS_HDR].concat(h[1].map(r => r.map(v => seguro_(v))));
      sh.getRange(1, 1, f.length, AVI_XLS_HDR.length).setValues(f).setVerticalAlignment('top');
      sh.getRange(1, 1, 1, AVI_XLS_HDR.length).setFontWeight('bold').setBackground('#0C1A3D').setFontColor('#FFFFFF').setWrap(true);
      sh.setFrozenRows(1); if (f.length > 1) sh.getRange(1, 1, f.length, AVI_XLS_HDR.length).createFilter();
      anchos.forEach((w, j) => sh.setColumnWidth(j + 1, w)); sh.getRange(2, 8, Math.max(f.length - 1, 1), 3).setWrap(true);
    });
    SpreadsheetApp.flush();
    const r = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx', { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) throw new Error('No se pudo generar el Excel (' + r.getResponseCode() + ').');
    return r.getBlob().setName(nombre + '.xlsx');
  } finally { try { DriveApp.getFileById(ss.getId()).setTrashed(true); } catch (e) {} }
}
/** Arma (y si enviar=true, envía) el resumen de cada jefe. conExcel=false en la vista previa. */
function aviJefes_(c, enviar, soloUno) {
  const ahora = new Date(), clave = 'JEFES|' + aviDia_(ahora);
  if (enviar && !soloUno) {
    if (aviHechos_()[clave]) return [];
    if (c.jefes === 'LUNES' && Utilities.formatDate(ahora, 'America/Lima', 'u') !== '1') return [];
  }
  const B = aviBase_(ahora), m = aviCorreos_(), kmax = Math.max.apply(null, B.map(x => aviMesK_(x.mes)).concat([-1]));
  const cierre = (B.find(x => aviMesK_(x.mes) === kmax) || {}).mes || '', fijos = c.jefesCopia ? aviMails_(c.copia) : [];
  const jefes = [...new Set(B.map(x => x.jefe).filter(j => j && !/vigencia/i.test(j)))].sort(), out = [];
  jefes.forEach(j => {
    const L = B.filter(x => norm_(x.jefe) === norm_(j)), ci = L.filter(x => aviMesK_(x.mes) === kmax), ob = L.filter(x => x.pend.length || x.rev.length);
    if (!ci.length && !ob.length) return;
    const cnt = k => ci.filter(x => norm_(x.estado) === k).length, porV = {};
    ob.forEach(x => (porV[x.vendedor] = porV[x.vendedor] || []).push(x));
    const fila = x => '<tr><td style="padding:6px;border:1px solid #D9DEE8;vertical-align:top"><b>' + aviEsc_(x.cliente) + '</b><br><small>' + aviEsc_(aviSub_(x) + ' · ' + aviCap_(x.mes)) + '</small></td>'
      + '<td style="padding:6px;border:1px solid #D9DEE8;vertical-align:top">' + (x.pend.length ? '<b style="color:#B42318">🔴 ' + x.pend.length + '</b> ' + aviEsc_(x.pend.map(p => p.punto).join(' / ')) : '') + (x.rev.length ? '<br><span style="color:#B54708">🕓 ' + x.rev.length + ' en revisión</span>' : '') + '</td>'
      + '<td style="padding:6px;border:1px solid #D9DEE8;vertical-align:top;white-space:nowrap;color:' + (x.dias != null && x.dias <= 0 ? '#B42318' : '#1F2937') + '">' + (x.lim ? aviEsc_(fmtPlazo_(x.lim)) + (x.dias < 0 ? '<br><b>vencido</b>' : x.dias === 0 ? '<br><b>vence hoy</b>' : '') : '—') + '</td></tr>';
    const html = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#1F2937;max-width:820px"><p>Hola ' + aviEsc_(aviCap_(String(j).split(' ')[0])) + ',</p><p>Este es el resumen de revisiones de tu equipo.</p>'
      + (ci.length ? '<p><b>Cierre ' + aviEsc_(aviCap_(cierre)) + ':</b> ' + ci.length + ' venta' + (ci.length > 1 ? 's' : '') + ' · 🟢 ' + cnt('COMPLETO') + ' completas · 🔴 ' + ci.filter(x => x.pend.length).length + ' con observaciones por subsanar · 🕓 ' + ci.filter(x => !x.pend.length && x.rev.length).length + ' en revisión de Contratos · 🟡 ' + cnt('PENDIENTE') + ' por revisar</p>' : '')
      + (ob.length ? '<h3 style="margin:14px 0 6px">Ventas con observaciones (' + ob.length + ')</h3>' + Object.keys(porV).sort().map(v => '<p style="margin:10px 0 4px"><b>👤 ' + aviEsc_(v) + '</b> · ' + porV[v].length + '</p><table style="border-collapse:collapse;width:100%"><tr style="background:#0C1A3D;color:#fff"><th style="padding:6px;text-align:left">Venta</th><th style="padding:6px;text-align:left">Qué falta</th><th style="padding:6px;text-align:left">Plazo</th></tr>' + porV[v].map(fila).join('') + '</table>').join('') : '<p>🎉 Tu equipo no tiene observaciones pendientes.</p>')
      + '<p>📎 Adjunto: Excel con todas las ventas del cierre ' + aviEsc_(aviCap_(cierre)) + ' de tu equipo y una hoja con todo lo pendiente.</p><p><a href="' + aviEsc_(c.portal) + '">Abrir el Cuadro de Revisiones</a></p><p style="color:#667;font-size:12px">Resumen automático de Contratos y Cobranzas · Grupo MG.</p></div>';
    out.push({ tipo: 'RESUMEN JEFE', vendedor: j, para: m[norm_(j)] || '', cc: fijos, sinCorreo: !m[norm_(j)], asunto: '📊 Revisiones de tu equipo · ' + aviCap_(j) + ' · ' + ob.filter(x => x.pend.length).length + ' con observaciones por subsanar',
      html: html, texto: 'Resumen de revisiones de tu equipo (' + aviCap_(j) + '). ' + ob.length + ' ventas con observaciones. Excel adjunto.\n' + c.portal, exps: ob.map(x => x.id), ci: ci, ob: ob, cierre: cierre });
  });
  if (!enviar) return out;
  const lista = soloUno ? out.slice(0, 1) : out, filas = [];
  lista.forEach(e => {
    let res = 'ENVIADO';
    try {
      const para = soloUno || e.para || (e.cc.length ? e.cc.join(',') : '');
      if (!para) res = 'SIN CORREO';
      else {
        const xls = aviXlsx_('Revisiones ' + aviCap_(e.vendedor) + ' - ' + aviCap_(e.cierre), [['Cierre ' + aviCap_(e.cierre), e.ci.map(aviXlsFila_)], ['Pendientes', e.ob.map(aviXlsFila_)]]);
        MailApp.sendEmail({ to: para, cc: soloUno ? '' : (e.para ? e.cc : []).join(','), subject: (soloUno ? '[PRUEBA] ' : '') + e.asunto, htmlBody: (soloUno ? '<p style="background:#EEF1F7;padding:8px;font-family:Arial;font-size:13px">Prueba: este resumen saldría a <b>' + aviEsc_(e.para || '(sin correo)') + '</b>.</p>' : '') + e.html, body: e.texto, attachments: [xls], name: 'Contratos y Cobranzas MG' });
      }
    } catch (err) { res = 'ERROR: ' + err.message; }
    if (!soloUno) filas.push([new Date(), e.tipo, e.vendedor, e.para, e.cc.join(', '), e.exps.join(', '), e.asunto, res, clave]);
    e.resultado = res;
  });
  if (filas.length) { const h = aviHoja_(); h.getRange(h.getLastRow() + 1, 1, filas.length, AVI_HDR.length).setValues(filas); }
  return lista;
}
