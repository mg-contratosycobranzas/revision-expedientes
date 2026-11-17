/**
 * 📋 PLANTILLAS DE OBSERVACIÓN (archivo independiente; se conecta solo con Code.gs).
 * Textos estándar que Contratos y Cobranzas inserta con un toque al observar, con lo que Comercial debe adjuntar o corregir.
 * Sus "palabras clave" sirven para agrupar las observaciones en los 📈 Indicadores (aunque se hayan escrito a mano).
 * Hoja OBS_PLANTILLAS (se crea sola con la lista inicial la primera vez que la administración guarda un cambio).
 * Leer: cualquier usuario con sesión. Agregar, editar o desactivar: solo administración.
 */
const HOJA_PLA = 'OBS_PLANTILLAS';
const PLA_HDR = ['Área', 'Observación', 'Qué debe adjuntar o corregir', 'Palabras clave (separadas por |)', 'Activa'];
const PLA_BASE = [
  ["SBS", "TRASLAPE PENDIENTE", "Sustento del traslape: voucher o constancia del pago trasladado", "TRASLAPE"],
  ["SBS", "SIN ROP / ROP INCOMPLETO", "ROP completo: profesión, ocupación, nacionalidad y dirección", "ROP"],
  ["SBS", "SIN VERIFICACIÓN DE BIENES", "Verificación de bienes marcada y firmada", "VERIFICACION DE BIENES|VERIF DE BIENES|VERIFIC"],
  ["SBS", "DJ DEL CLIENTE INCOMPLETA (FIRMA / RÚBRICA / FECHA / CASILLAS)", "DJ del cliente con firma, rúbrica en cada hoja, fecha y casillas marcadas", "DJ CLIENTE|DJ DEL CLIENTE|DJ DE CLIENTE|DJ CLIENTES|DJ NO INDICA|DJ DE CONYUGE|DJ DE TITULAR"],
  ["SBS", "DNI VENCIDO / CADUCADO", "DNI vigente y legible de todos los compradores", "DNI VENCIDO|DNI CADUCADO|DNI POCO LEGIBLE"],
  ["SBS", "COMPROBANTE INCOMPLETO (BANCO ORIGEN / CUENTA DE MG)", "Comprobante completo: banco de origen y cuenta de MG Proyecto", "COMPROBANTE|VOUCHER INCOMPLETO|EVIDENCIA CARGADA"],
  ["SBS", "BÚSQUEDAS INCOMPLETAS (INTERPOL / OFAC / MIGRACIONES / ESSALUD)", "Búsquedas con nombres completos y nacionalidad", "INTERPOL|OFAC|MIGRACION|ESSALUD"],
  ["CIERRE", "CHECK LIST SIN FIRMA", "Check list firmado por la jefatura", "CHECK LIST|CHECK LIUST"],
  ["CIERRE", "DJ DEL VENDEDOR PENDIENTE / SIN HUELLA", "DJ del vendedor con firma y huella", "DJ VENDEDOR|DJ DE VENDEDOR|DECLARACION JURADA DEL VENDEDOR|DECLARACION JURDA DEL VENDEDOR|DECLARACION JJ DEL VENDEDOR"],
  ["CIERRE", "PDT AHORRO CASA", "Voucher(s) de ahorro casa", "AHORRO CASA"],
  ["CIERRE", "PDT PRECA DE BANCO PROMOTOR", "Precalificación del banco promotor", "PRECA DE BANCO PROMOTOR|PRECA BANCO PROMOTOR|PRECA BANCO SPONSOR|SIN PRECA|PDT PRECA"],
  ["CIERRE", "FECHA DE DESEMBOLSO: 15 DÍAS DESPUÉS DE LA CUOTA INICIAL", "Corregir la fecha de desembolso en la MCV (Anexo I)", "DESEMBOLSO"],
  ["CIERRE", "FECHA DE ENTREGA INCORRECTA", "Corregir la fecha de entrega en la MCV", "FECHA DE ENTREGA"],
  ["CIERRE", "FORMA DE PAGO DE LA MCV NO COINCIDE CON LA PRECA", "Corregir la forma de pago o aclarar con la preca", "FORMA DE PAGO|NO COINCIDE CON PRECA|NO CONCUERDA CON MCV|DIFIERE DE PRECA"],
  ["CIERRE", "PRECIO POR DEBAJO DE LA LISTA", "Conformidad de Gerencia por el precio", "PRECIO POR DEBAJO|PRECIO MENOR|DEBAJO DE LISTA|DEBAJO DE PRECIO"],
  ["CIERRE", "PDT ACUERDO DE SEPARACIÓN (FIRMA / HUELLA)", "Acuerdo de separación firmado, con huella del cliente y firma de MG Proyecto", "ACUERDO DE SEPARACION|SEPARACION NO|SEPARACION FALTA"],
  ["CIERRE", "PDT PROFORMA", "Proforma de la unidad correcta", "PROFORMA"],
  ["CIERRE", "NOMBRE JUNTO / INCOMPLETO / SIN TILDE EN DOCUMENTOS", "Corregir el nombre como figura en el DNI en todos los documentos", "NOMBRE|APELLIDO"],
  ["CIERRE", "DESCRIPCIÓN DEL DEPARTAMENTO INCORRECTA / SIN ÁREA LIBRE", "Corregir la descripción y áreas en la MCV", "DESCRIPCION|AREA LIBRE|AREA OCUPADA"],
  ["CIERRE", "PDT FIRMA EN MG PROYECTO", "Documento con la firma de MG Proyecto", "FIRMA EN MG"],
  ["CIERRE", "CONSTANCIA DE BONIFICACIÓN INCORRECTA", "Constancia de bonificación corregida", "CONSTANCIA DE BONIFICACION|PREMIO"],
  ["CIERRE", "SOLICITUD DE MODIFICACIÓN INCOMPLETA", "Juegos completos de solicitud de modificación", "SOLICITUD DE MODIFICACION"],
  ["CIERRE", "MARGEN / PRECIO PARA CARTA FIANZA", "Sustento del margen para carta fianza", "CARTA FIANZA"]
];
try {
  ['plantillasLista', 'plantillaGuardar', 'plantillaQuitar'].forEach(f => { if (API_PUBLICA.indexOf(f) < 0) API_PUBLICA.push(f); });
  if (SOLO_LECTURA.indexOf('plantillasLista') < 0) SOLO_LECTURA.push('plantillasLista');
} catch (e) {}

function plaAdmin_(u) { if (u.rol !== 'ADMIN' && u.permisos.indexOf('accesos') < 0) throw new Error('Solo la administración edita las plantillas.'); }
function plaHoja_() {
  const ss = ss_(); let h = ss.getSheetByName(HOJA_PLA);
  if (!h) {
    h = ss.insertSheet(HOJA_PLA);
    h.getRange(1, 1, 1, PLA_HDR.length).setValues([PLA_HDR]).setFontWeight('bold').setBackground('#0C1A3D').setFontColor('#FFFFFF');
    h.setFrozenRows(1);
    h.getRange(2, 1, PLA_BASE.length, PLA_HDR.length).setValues(PLA_BASE.map(p => [p[0], p[1], p[2], p[3], 'SI']));
  }
  return h;
}
const plaObj_ = (r, fila) => ({ fila: fila, area: norm_(r[0]) === 'SBS' ? 'SBS' : 'CIERRE', texto: String(r[1] || '').trim(), adjuntar: String(r[2] || '').trim(), claves: String(r[3] || '').trim(), activa: norm_(r[4]) !== 'NO' });
/** Lista de plantillas. Si la hoja aún no existe, devuelve la lista inicial (sin crear nada). */
function plantillasLista(token) {
  sesion_(token);
  const h = ss_().getSheetByName(HOJA_PLA);
  if (!h || h.getLastRow() < 2) return PLA_BASE.map((p, i) => plaObj_([p[0], p[1], p[2], p[3], 'SI'], i + 2))   // misma fila que tendrá al crearse la hoja;
  return h.getRange(2, 1, h.getLastRow() - 1, PLA_HDR.length).getValues().map((r, i) => plaObj_(r, i + 2)).filter(p => p.texto);
}
/** d = { fila?, area, texto, adjuntar, claves, activa } · sin fila = nueva. */
function plantillaGuardar(token, d) {
  const u = sesion_(token); plaAdmin_(u); d = d || {};
  const area = norm_(d.area) === 'SBS' ? 'SBS' : 'CIERRE', texto = String(d.texto || '').trim().toUpperCase().slice(0, 200);
  if (texto.length < 4) throw new Error('Escribe la observación.');
  const fila = [area, seguro_(texto), seguro_(String(d.adjuntar || '').trim().slice(0, 300)), seguro_(String(d.claves || '').trim().toUpperCase().slice(0, 300)), d.activa === false ? 'NO' : 'SI'];
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const h = plaHoja_(), n = Number(d.fila);
    if (n >= 2 && n <= h.getLastRow()) h.getRange(n, 1, 1, PLA_HDR.length).setValues([fila]); else h.appendRow(fila);
  } finally { lock.releaseLock(); }
  log_(u.nombre, 'PLANTILLA', area + ' · ' + texto);
  return plantillasLista(token);
}
/** Desactiva una plantilla (no se borra: los indicadores la siguen usando para agrupar lo anterior si se reactiva). */
function plantillaQuitar(token, fila) {
  const u = sesion_(token); plaAdmin_(u);
  const h = plaHoja_(), n = Number(fila); if (!(n >= 2 && n <= h.getLastRow())) throw new Error('Plantilla no encontrada.');
  h.getRange(n, 5).setValue('NO');
  log_(u.nombre, 'PLANTILLA_QUITADA', String(h.getRange(n, 2).getValue()));
  return plantillasLista(token);
}
