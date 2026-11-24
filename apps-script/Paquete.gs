/**
 * ⚡ MODO INSTANTÁNEO — PAQUETE (archivo independiente; se conecta solo con Code.gs).
 *
 * Problema: cada consulta a Google tarda 2–3 s como mínimo, y abrir un expediente hace varias (detalle, ficha,
 * puntos, habilitación) que vuelven a leer las hojas: 15–30 s.
 *
 * Solución: el portal pide UNA vez el "paquete" con todo lo que la persona puede ver (lista, detalle, ficha,
 * puntos y habilitación de cada expediente) y lo guarda en el equipo. Abrir un expediente ya no consulta a Google.
 *
 * Cómo se arma sin duplicar reglas: durante el armado, ss_(), sesion_() y misExpedientes() se reemplazan por
 * versiones "con memoria" (cada hoja se lee UNA sola vez y queda en memoria). Luego se llaman las MISMAS funciones
 * de siempre (detalle, ficha, puntosEstado, habEstado) para cada expediente. Así el resultado es idéntico al
 * de abrir el expediente normal, pero en milisegundos. Al terminar se restauran las funciones originales.
 * Solo se lee: el armado no escribe en ninguna hoja.
 *
 * Vigencia: el paquete lleva la versión de datos (verDatos_). Cualquier guardado desde el portal o edición directa
 * en el Sheet cambia esa versión; el portal lo nota y pide el paquete nuevo en segundo plano.
 * El paquete COMÚN se arma una sola vez por versión (para todos) y cada persona recibe solo lo que puede ver.
 * Ambos quedan 30 minutos en caché. ⚡ Se habilita por persona con el permiso "instantaneo" (la administración siempre).
 */
try {
  if (PERMISOS.indexOf('instantaneo') < 0) PERMISOS.push('instantaneo');   // ⚡ se habilita por persona en Accesos
  ['paquete', 'paqVerificar'].forEach(f => {
    if (API_PUBLICA.indexOf(f) < 0) API_PUBLICA.push(f);
    if (SOLO_LECTURA.indexOf(f) < 0) SOLO_LECTURA.push(f);
  });
} catch (e) {}

const PAQ_MAX_MS = 150000, PAQ_SEG = 1800, PAQ_TROZO = 90000, PAQ_SUPER = '__paquete_comun__';

/* ---------------- memoria de hojas (solo lectura) ---------------- */
function paqProxy_(obj, real) {
  return new Proxy(obj, { get: (t, p) => {
    if (p in t) return t[p];
    const r = typeof real === 'function' ? real() : real, v = r[p];
    return typeof v === 'function' ? v.bind(r) : v;
  } });
}
function paqHoja_(real) {
  const H = { _v: null, _d: null, _f: null, _dv: null, _mr: {}, _lr: null, _lc: null };
  const lr = () => H._lr != null ? H._lr : (H._lr = real.getLastRow()), lc = () => H._lc != null ? H._lc : (H._lc = real.getLastColumn());
  const grid = k => {
    if (H[k]) return H[k];
    const n = Math.max(lr(), 1), m = Math.max(lc(), 1), rg = real.getRange(1, 1, n, m);
    H[k] = k === '_v' ? rg.getValues() : k === '_d' ? rg.getDisplayValues() : k === '_f' ? rg.getFormulas() : rg.getDataValidations();
    return H[k];
  };
  const corte = (k, r, c, nr, nc) => {
    const g = grid(k), vacio = k === '_dv' ? null : '', out = [];
    for (let i = 0; i < nr; i++) {
      const fila = g[r - 1 + i], o = new Array(nc);
      for (let j = 0; j < nc; j++) { const v = fila ? fila[c - 1 + j] : undefined; o[j] = v === undefined ? vacio : v; }
      out.push(o);
    }
    return out;
  };
  const hoja = {
    getName: () => real.getName(), getSheetId: () => real.getSheetId(), getLastRow: lr, getLastColumn: lc,
    getRange: function (a, b, c, d) {
      if (typeof a !== 'number') return real.getRange.apply(real, arguments);   // notación A1: directo a la hoja real
      const r = a, col = b || 1, nr = c || 1, nc = d || 1;
      let rr = null; const realRg = () => rr || (rr = real.getRange(r, col, nr, nc));
      const rng = {
        getValues: () => corte('_v', r, col, nr, nc), getDisplayValues: () => corte('_d', r, col, nr, nc),
        getFormulas: () => corte('_f', r, col, nr, nc), getDataValidations: () => corte('_dv', r, col, nr, nc),
        getValue: () => corte('_v', r, col, 1, 1)[0][0], getDisplayValue: () => corte('_d', r, col, 1, 1)[0][0],
        getFormula: () => corte('_f', r, col, 1, 1)[0][0], getDataValidation: () => corte('_dv', r, col, 1, 1)[0][0],
        getRow: () => r, getColumn: () => col, getNumRows: () => nr, getNumColumns: () => nc, getLastRow: () => r + nr - 1, getLastColumn: () => col + nc - 1,
        getSheet: () => hoja,
        getMergedRanges: () => { const k = r + ',' + col + ',' + nr + ',' + nc; return H._mr[k] || (H._mr[k] = realRg().getMergedRanges()); }
      };
      return paqProxy_(rng, realRg);   // cualquier otra cosa (incluso escribir) va a la hoja real
    }
  };
  return paqProxy_(hoja, real);
}
function paqLibro_(real) {
  const hojas = {};
  const libro = { getSheetByName: n => { if (!(n in hojas)) { const h = real.getSheetByName(n); hojas[n] = h ? paqHoja_(h) : null; } return hojas[n]; }, getId: () => real.getId() };
  return paqProxy_(libro, real);
}

/** Ejecuta fn con ss_ / sesion_ / misExpedientes en memoria y luego deja todo como estaba. */
function paqConMemoria_(fn) {
  const G = (function () { return this; })() || globalThis, orig = { ss_: G.ss_, sesion_: G.sesion_, misExpedientes: G.misExpedientes };
  const libro = paqLibro_(orig.ss_()), ses = {}, exps = {};
  G.ss_ = () => libro;
  G.sesion_ = t => t === PAQ_SUPER ? paqSuper_() : (ses[t] || (ses[t] = orig.sesion_(t)));
  G.misExpedientes = t => JSON.parse(exps[t] || (exps[t] = JSON.stringify(orig.misExpedientes(t))));   // copia nueva cada vez, como hoy
  // si este motor no permite el reemplazo (p. ej. ss_ no es una función global), se arma sin memoria: más lento pero igual de correcto
  let ok = false; try { ok = ss_() === libro; } catch (e) {}
  if (!ok) { G.ss_ = orig.ss_; G.sesion_ = orig.sesion_; G.misExpedientes = orig.misExpedientes; return fn(); }
  try { return fn(); } finally { G.ss_ = orig.ss_; G.sesion_ = orig.sesion_; G.misExpedientes = orig.misExpedientes; }
}

/* ---------------- armado COMÚN (una sola vez por versión de datos, para todos) ----------------
 * Se arma con una sesión interna "Sistema" (rol ADMIN, todos los permisos): así contiene TODOS los expedientes,
 * con las fichas de ambas áreas. Luego a cada persona se le entrega solo lo que puede ver (paqParaUsuario_). */
function paqSuper_() {
  let per = [];
  try { per = conAreas_('ADMIN', permisosDe_('ADMIN', ''), 'Sistema', ''); } catch (e) {}
  if (!per || !per.length) try { per = PERMISOS.slice(); } catch (e) { per = []; }
  ['ver', 'revisar', 'cierre_ver', 'cierre_edit', 'sbs_ver', 'sbs_edit'].forEach(p => { if (per.indexOf(p) < 0) per.push(p); });
  return { fila: 0, nombre: 'Sistema', correo: '', rol: 'ADMIN', ref: '', permisos: per, como: '' };
}
const paqClon_ = o => o == null ? o : JSON.parse(JSON.stringify(o));
const paqObs_ = x => x.estado === 'OBSERVADO' || x.estado === 'SUBSANADO';
function paqArmarComun_() {
  const t0 = Date.now(), ver = verDatos_(), tk = PAQ_SUPER;
  return paqConMemoria_(() => {
    const lista = paqClon_(misExpedientes(tk)) || [];
    const P = { v: ver, t: Date.now(), det: {}, fic: {}, pun: {}, falta: [], ms: 0, n: lista.length };
    // primero lo que más se abre: observados y en revisión, luego pendientes, luego aprobados (los más nuevos primero)
    const peso = x => x.anulado ? 4 : paqObs_(x) ? 0 : x.estado === 'COMPLETO' ? 2 : 1;
    const orden = lista.map((x, i) => ({ x: x, i: i })).sort((a, b) => peso(a.x) - peso(b.x) || b.i - a.i).map(o => o.x);
    for (let k = 0; k < orden.length; k++) {
      const x = orden[k], id = x.id;
      if (Date.now() - t0 > PAQ_MAX_MS) { P.falta = orden.slice(k).map(o => o.id); break; }
      try { P.det[id] = paqClon_(typeof detalleCompleto_ === 'function' ? detalleCompleto_(tk, id) : detalle(tk, id)); } catch (e) { P.det[id] = { error: String(e.message || e) }; }
      try { P.fic[id] = paqClon_(ficha(tk, id)); } catch (e) {}
      if (paqObs_(x) && typeof puntosEstado === 'function') try { P.pun[id] = paqClon_(puntosEstado(tk, id)); } catch (e) {}
    }
    P.ms = Date.now() - t0;
    return P;
  });
}

/* ---------------- caché (en trozos: cada valor de caché admite ~100 KB) ---------------- */
function paqLeer_(clave) {
  const c = CacheService.getScriptCache(), n = Number(c.get(clave + '_n') || 0); if (!n) return null;
  const ks = []; for (let i = 0; i < n; i++) ks.push(clave + '_' + i);
  const m = c.getAll(ks), partes = ks.map(k => m[k]); if (partes.some(p => !p)) return null;
  return partes.join('');
}
function paqGuardar_(clave, txt) {
  const c = CacheService.getScriptCache(), o = {}; let n = 0;
  for (let i = 0; i < txt.length; i += PAQ_TROZO) o[clave + '_' + (n++)] = txt.slice(i, i + PAQ_TROZO);
  if (n > 300) return;   // demasiado grande para la caché: se arma cada vez
  try { c.putAll(o, PAQ_SEG); c.put(clave + '_n', String(n), PAQ_SEG); } catch (e) {}
}
const paqGz_ = o => Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(JSON.stringify(o), 'application/json')).getBytes());
const paqDesGz_ = b => JSON.parse(Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(b), 'application/x-gzip')).getDataAsString());

/** El paquete común de la versión v: de la caché, o se arma. Si otra persona ya lo está armando, se espera a que termine. */
function paqComun_(v) {
  const c = CacheService.getScriptCache(), clave = 'paqc_' + v, flag = 'paqc_armando_' + v;
  let gz = paqLeer_(clave);
  if (!gz && c.get(flag)) {
    const fin = Date.now() + PAQ_MAX_MS + 20000;
    while (!gz && Date.now() < fin && c.get(flag)) { Utilities.sleep(2000); gz = paqLeer_(clave); }
  }
  if (gz) return paqDesGz_(gz);
  c.put(flag, '1', Math.ceil(PAQ_MAX_MS / 1000) + 30);
  try {
    const P = paqArmarComun_();
    if (!P.falta.length) paqGuardar_(clave, paqGz_(P));
    return P;
  } finally { c.remove(flag); }
}

/* ---------------- lo que le toca a cada persona ---------------- */
/** Ficha tal como la recibiría esa persona: áreas que no puede ver → oculto; áreas que no puede editar → solo lectura. */
function paqFichaPara_(u, f) {
  if (!f) return f;
  const r = paqClon_(f);
  Object.keys(FICHA_HOJAS).forEach(a => {
    if (!puedeArea_(u, a, 'ver')) { r[a] = { oculto: true }; return; }
    if (r[a] && !r[a].oculto) r[a].soloLectura = !puedeArea_(u, a, 'edit');
  });
  return r;
}
function paqParaUsuario_(token, C) {
  return paqConMemoria_(() => {
    const u = sesion_(token), lista = paqClon_(misExpedientes(token)) || [];
    const revisa = !!(u.permisos && u.permisos.indexOf('revisar') >= 0);
    const P = { v: C.v, t: C.t, lista: lista, det: {}, fic: {}, pun: {}, hab: {}, falta: [], tab: null, cal: null, ms: C.ms, armado: C.t };
    lista.forEach(x => {
      const id = x.id;
      if (!(id in C.det)) { P.falta.push(id); return; }
      P.det[id] = C.det[id];
      if (revisa && C.fic[id]) P.fic[id] = paqFichaPara_(u, C.fic[id]);
      if (paqObs_(x)) {
        if (id in C.pun) P.pun[id] = C.pun[id];
        if (typeof habEstado === 'function') try { P.hab[id] = paqClon_(habEstado(token, id)); } catch (e) {}
      }
    });
    if (typeof puntosTablero === 'function' && lista.some(paqObs_)) try { P.tab = paqClon_(puntosTablero(token)); } catch (e) {}
    try { P.cal = paqClon_(calendario(token)); } catch (e) {}
    return P;
  });
}

/**
 * El portal: paquete(token, versionQueTengo).
 * Si su versión sigue vigente → { igual: true, v } (respuesta inmediata).
 * Si no → { v, gz } con el paquete de esa persona en JSON comprimido (gzip, base64).
 */
function paquete(token, vCliente) {
  const u = sesion_(token);
  const v = verDatos_();
  if (vCliente && String(vCliente) === String(v)) return { igual: true, v: v };
  const per = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, (u.permisos || []).join(',') + '|' + u.rol + '|' + (u.como || ''))).slice(0, 12);
  const clave = 'paq2_' + u.fila + '_' + per + '_' + v;
  let gz = paqLeer_(clave);
  if (!gz) {
    const P = paqParaUsuario_(token, paqComun_(v));
    gz = paqGz_(P);
    if (!P.falta.length) paqGuardar_(clave, gz);
  }
  return { v: v, gz: gz };
}

/**
 * Control de calidad (ADMIN, o ADMIN usando "Ver como"): compara, para n expedientes, lo que entrega el paquete
 * con lo que da abrir el expediente de la forma normal. Debe dar 0 diferencias.
 */
function paqVerificar(token, n) {
  const u = sesion_(token); if (u.rol !== 'ADMIN' && !u.como) throw new Error('Solo la administración.');
  const P = paqParaUsuario_(token, paqComun_(verDatos_()));
  const ids = Object.keys(P.det).sort((a, b) => (P.pun[b] ? 1 : 0) - (P.pun[a] ? 1 : 0)).slice(0, Math.max(1, Math.min(Number(n) || 10, 40)));
  const revisa = !!(u.permisos && u.permisos.indexOf('revisar') >= 0), dif = [];
  paqConMemoria_(() => {
    ids.forEach(id => {
      const x = P.lista.find(z => z.id === id) || {};
      const cmp = (tipo, a, b) => { if (JSON.stringify(a) !== JSON.stringify(b)) dif.push({ id: id, tipo: tipo, normal: JSON.stringify(a).slice(0, 300), paquete: JSON.stringify(b).slice(0, 300) }); };
      try { cmp('detalle', paqClon_(typeof detalleCompleto_ === 'function' ? detalleCompleto_(token, id) : detalle(token, id)), P.det[id]); } catch (e) { dif.push({ id: id, tipo: 'detalle', normal: 'ERROR ' + e.message, paquete: '' }); }
      if (revisa) try { cmp('ficha', paqClon_(ficha(token, id)), P.fic[id]); } catch (e) {}
      if (paqObs_(x) && typeof puntosEstado === 'function') try { cmp('puntos', paqClon_(puntosEstado(token, id)), P.pun[id]); } catch (e) {}
    });
  });
  return { revisados: ids.length, diferencias: dif.length, ejemplos: dif.slice(0, 5), usuario: u.nombre, armado: P.ms };
}
