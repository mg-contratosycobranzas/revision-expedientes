/**
 * AYUDA EDITABLE — las guías rápidas de 💡 Ayuda viven en la hoja AYUDA del Sheet (archivo independiente, se conecta solo con Code.gs).
 * Columnas: Orden | Ícono | Título | Puntos (uno por línea) | Visible (SI/NO).
 * La administración las edita desde el portal (✏️ Editar guías) o directo en el Sheet: así un correo o un paso se cambia en un solo lugar.
 * Solo las ve quien inició sesión (los correos ya no quedan en la página pública).
 */
const HOJA_AYUDA = 'AYUDA', AYUDA_HDR = ['Orden', 'Ícono', 'Título', 'Puntos (uno por línea)', 'Visible (SI/NO)'];
const AYUDA_INICIAL = [   // se copian a la hoja AYUDA solo la primera vez
  ["📋", "Checklist comercial obligatorio", ["Proforma y cotización: precios de la lista oficial, promociones vigentes, forma de pago detallada y firma y huella del cliente.","Reserva y depósito: comprobante legible del abono exacto de separación e inicial, con fechas coherentes con la separación y la minuta.","Minuta de compraventa: visada por legal, con firmas y huellas completas y todos los anexos del proyecto.","Si falta un solo documento, la validación se detiene y se retrasa el cobro de comisiones."]],
  ["🛡️", "Validaciones SBS (prevención de lavado de activos)", ["Declaración jurada de origen de fondos: formato oficial, llenada a mano o con firma digital clara, sin tachas, con firma y huella en todas las hojas.","Filtros internacionales: consulta OFAC e Interpol.","Validación local: RUC activo en SUNAT y actividad laboral coherente (EsSalud)."]],
  ["👥", "Requisitos por tipo de cliente", ["Persona natural: DNI/CE vigente, DDJJ de fondos, consultas SUNAT, Interpol y OFAC, y minuta firmada.","Persona jurídica: lo mismo para el representante legal, más la empresa y sus socios (accionistas con más del 25%)."]],
  ["🧾", "El voucher perfecto", ["Fecha y hora de la operación visibles.","Monto pagado igual al registrado en Evolta (importe en S/ visible).","N° de operación bancaria completo.","Nombre del banco (BCP, Interbank, BBVA, etc.).","Cuentas de origen y destino de MG.","Imagen nítida y completa, cargada en Evolta."]],
  ["💻", "Evolta: el vendedor es dueño del dato", ["El contrato se emite automáticamente desde Evolta con los datos que ingresas.","Si un dato está mal en Evolta, el contrato sale con fallas y hay que corregirlo, retrasando el cierre."]],
  ["⚖️", "Textos legales", ["El acuerdo de separación, los contratos y el Anexo I son formatos estándar: no se modifican a mano.","Cualquier excepción se consulta con Jefatura antes de comprometerse con el cliente."]],
  ["🗓️", "Ciclo semanal de validación de pagos", ["Miércoles: cierre de recepción de comprobantes.","Jueves a lunes: revisión interna y envío a Contabilidad.","Martes: Contabilidad devuelve los pagos validados.","Después: se solicita la boleta o factura final.","Casos atípicos: 24 a 48 horas aproximadamente."]],
  ["🔄", "Flujo de validación del expediente", ["1. Armado: reúnes el checklist completo firmado por el cliente.","2. Filtro SBS: DDJJ y consultas OFAC, Interpol y SUNAT.","3. Visación: Contratos verifica las firmas e ingresa el contrato para visado.","4. Cierre: se aprueba el expediente y empieza el cobro de comisiones."]],
  ["🔁", "Traslapes o pagos con error en cuenta", ["Aplica cuando el cliente separó en otro proyecto o el pago se hizo a una cuenta equivocada por un error involuntario.","1. Gestión y aprobación: el jefe de ventas gestiona primero la solicitud con Gerencia General y obtiene el conforme para trasladar el pago.","2. Expediente físico: con el conforme, el jefe de ventas anexa el sustento impreso al expediente en físico.","3. Notificación y regularización: reenvía el correo con la aprobación de Gerencia a abonos@grupomg.pe, jvelez@grupomg.pe y mvalles@grupomg.pe.","📌 Asunto del correo (obligatorio, con esta estructura): Nombres y apellidos del cliente + DNI + Proyecto + Unidad (Dpto / Estac / Depósito)."]],
  ["🏦", "Clientes con carta de aprobación (crédito directo / hipotecario)", ["Si tu cliente ya tiene carta de aprobación del banco, envíala por correo para que Contratos y Cobranzas gestione el desembolso.","Para: mvalles@grupomg.pe; msandoval@grupomg.pe; rmurga@grupomg.pe; jhumala@grupomg.pe","📌 Asunto: Proyecto + Unidad + DNI del titular.","Importante: incluye los datos del ejecutivo del banco (nombre, teléfono y correo) para coordinar."]]
];
try {
  ['ayudaGuias', 'ayudaGuardar'].forEach(f => { if (API_PUBLICA.indexOf(f) < 0) API_PUBLICA.push(f); });
  if (SOLO_LECTURA.indexOf('ayudaGuias') < 0) SOLO_LECTURA.push('ayudaGuias');
} catch (e) {}

function hojaAyuda_() {
  const ss = ss_(); let h = ss.getSheetByName(HOJA_AYUDA);
  if (!h) {
    h = ss.insertSheet(HOJA_AYUDA);
    h.getRange(1, 1, 1, AYUDA_HDR.length).setValues([AYUDA_HDR]).setFontWeight('bold').setBackground('#0C1A3D').setFontColor('#FFFFFF');
    h.getRange(2, 1, AYUDA_INICIAL.length, 5).setValues(AYUDA_INICIAL.map((g, i) => fila_([i + 1, g[0], g[1], g[2].join('\n'), 'SI'])));
    h.setFrozenRows(1); h.setColumnWidth(3, 320); h.setColumnWidth(4, 700); h.getRange('D:D').setWrap(true);
  }
  return h;
}
function ayudaLeer_(todas) {
  const h = hojaAyuda_(); if (h.getLastRow() < 2) return [];
  return h.getRange(2, 1, h.getLastRow() - 1, 5).getValues()
    .filter(r => String(r[2]).trim() && (todas || norm_(r[4]) !== 'NO'))
    .map(r => ({ orden: Number(r[0]) || 999, icono: String(r[1] || '📘').trim(), titulo: String(r[2]).trim(), puntos: String(r[3] || '').split(/\r?\n/).map(p => p.trim()).filter(Boolean), visible: norm_(r[4]) !== 'NO' }))
    .sort((a, b) => a.orden - b.orden);
}
/** Guías para cualquier usuario con sesión. Con editar=true (solo administración) incluye también las ocultas. */
function ayudaGuias(token, editar) {
  const u = sesion_(token);
  if (editar) { puede_(u, 'accesos'); return ayudaLeer_(true); }
  return ayudaLeer_(false);
}
/** Reemplaza todas las guías (administración). lista = [{icono, titulo, puntos:[...], visible}] en el orden deseado. */
function ayudaGuardar(token, lista) {
  const u = sesion_(token); puede_(u, 'accesos');
  if (!Array.isArray(lista)) throw new Error('Lista inválida.');
  const filas = lista.map((g, i) => {
    const t = String(g && g.titulo || '').trim().slice(0, 150); if (!t) throw new Error('La guía ' + (i + 1) + ' no tiene título.');
    const p = (Array.isArray(g.puntos) ? g.puntos : String(g.puntos || '').split(/\r?\n/)).map(x => String(x).trim().slice(0, 600)).filter(Boolean).slice(0, 30);
    return fila_([i + 1, String(g.icono || '📘').trim().slice(0, 8), t, p.join('\n'), g.visible === false ? 'NO' : 'SI']);
  });
  if (filas.length > 60) throw new Error('Máximo 60 guías.');
  const h = hojaAyuda_(), lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (h.getLastRow() >= 2) h.getRange(2, 1, h.getLastRow() - 1, 5).clearContent();
    if (filas.length) h.getRange(2, 1, filas.length, 5).setValues(filas);
  } finally { lock.releaseLock(); }
  log_(u.nombre, 'AYUDA_EDITADA', filas.length + ' guías');
  return ayudaLeer_(true);
}
