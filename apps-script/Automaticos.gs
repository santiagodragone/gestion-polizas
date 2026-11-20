/**
 * AUTOMÁTICOS — Seguros Dragone
 * 1. Copia de seguridad semanal de toda la planilla (domingos a la madrugada, guarda las últimas 8).
 * 2. Resumen diario por mail a cada persona del equipo (lunes a viernes a las 8):
 *    sus tareas de Mi día, impagos a reclamar, cupones que vencen hoy, pólizas para enviar y cumpleaños.
 *
 * Quién recibe el resumen y qué secciones se arma en la hoja "Equipo" (se crea sola).
 * Activación: menú "Sistema Dragone" > "Activar copia semanal y resumen diario" (pide autorización una vez).
 * En el Apps Script en vivo este código está al final de cobranza.gs.
 */

var AUT_TZ = 'America/Argentina/Buenos_Aires';
var AUT_URL = 'https://santiagodragone.github.io/gestion-polizas/';
var AUT_CARPETA = 'Copias de seguridad — Sistema Dragone';
var AUT_COPIAS = 8;                 // copias semanales que se conservan
var AUT_HORA_RESUMEN = 8;           // hora del resumen diario
var AUT_DIAS_ENVIO = 30, AUT_DIAS_ATRAS = 45;      // misma regla que Renovaciones en el sistema
var AUT_EQUIPO = { nombre: 'Equipo', cols: ['Nombre', 'Mail', 'Recibe resumen', 'Cobranza', 'Renovaciones'] };
var AUT_VIGENCIAS = { mensual: 1, bimestral: 2, trimestral: 3, cuatrimestral: 4, semestral: 6, anual: 12 };
var AUT_MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
var AUT_DIAS = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];

// ── Activación ─────────────────────────────────────────────────────────────
function instalarAutomaticos() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (['backupSemanal', 'resumenDiario'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('backupSemanal').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(3).inTimezone(AUT_TZ).create();
  ScriptApp.newTrigger('resumenDiario').timeBased().everyDays(1).atHour(AUT_HORA_RESUMEN).inTimezone(AUT_TZ).create();
  autHojaEquipo_();
  autAviso_('Listo. La copia de seguridad corre todos los domingos y el resumen sale de lunes a viernes a las ' + AUT_HORA_RESUMEN +
    '. Completá los mails en la hoja "Equipo".');
}

function autAviso_(texto) {
  try { SpreadsheetApp.getUi().alert(texto); } catch (e) { Logger.log(texto); }
}

// ── 1. Copia de seguridad ─────────────────────────────────────────────────
function backupSemanal() {
  var props = PropertiesService.getScriptProperties(), carpeta = null, id = props.getProperty('AUT_CARPETA_ID');
  if (id) { try { carpeta = DriveApp.getFolderById(id); if (carpeta.isTrashed()) carpeta = null; } catch (e) { carpeta = null; } }
  if (!carpeta) { carpeta = DriveApp.createFolder(AUT_CARPETA); props.setProperty('AUT_CARPETA_ID', carpeta.getId()); }

  var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(props.getProperty('AUT_SS_ID'));
  var nombre = ss.getName() + ' — copia ' + Utilities.formatDate(new Date(), AUT_TZ, 'yyyy-MM-dd HH.mm');
  DriveApp.getFileById(ss.getId()).makeCopy(nombre, carpeta);

  // Conserva las últimas AUT_COPIAS; las más viejas van a la papelera de Drive (se pueden recuperar 30 días)
  var archivos = [], it = carpeta.getFiles();
  while (it.hasNext()) archivos.push(it.next());
  archivos.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  archivos.slice(AUT_COPIAS).forEach(function (f) { f.setTrashed(true); });
  return nombre;
}

function backupAhora() {
  var n = backupSemanal();
  autAviso_('Copia guardada en Drive, carpeta "' + AUT_CARPETA + '":\n' + n);
}

// ── 2. Resumen diario ─────────────────────────────────────────────────────
function autHojaEquipo_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(AUT_EQUIPO.nombre);
  PropertiesService.getScriptProperties().setProperty('AUT_SS_ID', ss.getId());
  if (sh) return sh;
  sh = ss.insertSheet(AUT_EQUIPO.nombre);
  sh.getRange(1, 1, 200, AUT_EQUIPO.cols.length).setNumberFormat('@');
  sh.getRange(1, 1, 1, AUT_EQUIPO.cols.length).setValues([AUT_EQUIPO.cols]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange(2, 1, 4, 5).setValues([['Santiago', '', 'Sí', 'Sí', 'Sí'], ['Bruno', '', 'Sí', 'Sí', 'Sí'], ['Leo', '', 'Sí', 'Sí', 'Sí'], ['Roberto', '', 'Sí', 'Sí', 'Sí']]);
  sh.getRange('C2:E200').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Sí', 'No'], true).build());
  sh.getRange('B1').setNote('Mail donde recibe el resumen diario. Sin mail, no se le manda.');
  sh.getRange('D1').setNote('Si incluye impagos y cupones que vencen hoy.');
  sh.getRange('E1').setNote('Si incluye las pólizas para enviar.');
  sh.autoResizeColumns(1, 5);
  return sh;
}

function resumenDiario() {
  var hoy = autHoy_();
  if (hoy.getUTCDay() === 0 || hoy.getUTCDay() === 6) return;       // fines de semana no
  var datos = autDatos_(hoy);
  autDestinatarios_().forEach(function (p) {
    if (!p.mail) return;
    var r = autResumenDe_(p, datos, hoy);
    MailApp.sendEmail({ to: p.mail, subject: r.asunto, htmlBody: r.html, name: 'Sistema Dragone' });
  });
}

/** Manda a tu propio mail el resumen de hoy de cada persona del equipo, para ver cómo queda. */
function resumenPrueba() {
  var hoy = autHoy_(), datos = autDatos_(hoy), yo = Session.getEffectiveUser().getEmail(), n = 0;
  autDestinatarios_().forEach(function (p) {
    var r = autResumenDe_(p, datos, hoy);
    MailApp.sendEmail({ to: yo, subject: '[Prueba] ' + r.asunto, htmlBody: r.html, name: 'Sistema Dragone' }); n++;
  });
  autAviso_('Te mandé ' + n + ' resumen(es) de prueba a ' + yo + '.');
}

function autDestinatarios_() {
  var sh = autHojaEquipo_(), v = sh.getDataRange().getValues(), enc = v.shift().map(function (h) { return String(h).trim(); });
  return v.map(function (f) { var o = {}; enc.forEach(function (h, i) { o[h] = f[i]; }); return o; })
    .filter(function (o) { return String(o['Nombre']).trim() && autSi_(o['Recibe resumen'], true); })
    .map(function (o) {
      return { nombre: String(o['Nombre']).trim(), mail: String(o['Mail']).trim(),
               cobranza: autSi_(o['Cobranza'], true), renovaciones: autSi_(o['Renovaciones'], true) };
    });
}

// ── Lectura de datos (las mismas reglas que el sistema web) ────────────────
function autSi_(v, porDefecto) {
  var s = String(v == null ? '' : v).trim().toLowerCase();
  if (s === '') return !!porDefecto;
  return ['sí', 'si', 'true', 'yes', 'x'].indexOf(s) >= 0;
}
function autDia_(y, m, d) { return new Date(Date.UTC(y, m, d)); }
function autHoy_() {
  var s = Utilities.formatDate(new Date(), AUT_TZ, 'yyyy-MM-dd').split('-');
  return autDia_(+s[0], +s[1] - 1, +s[2]);
}
function autSumar_(d, n) { return new Date(d.getTime() + n * 86400000); }
function autDiasMes_(y, m) { return new Date(Date.UTC(y, m + 1, 0)).getUTCDate(); }
function autISO_(d) { return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd'); }
function autCorta_(d) { return Utilities.formatDate(d, 'UTC', 'dd/MM'); }
function autFecha_(v) {
  if (v instanceof Date) { var s = Utilities.formatDate(v, AUT_TZ, 'yyyy-MM-dd').split('-'); return autDia_(+s[0], +s[1] - 1, +s[2]); }
  var t = String(v == null ? '' : v).trim(), x;
  if ((x = t.match(/^(\d{4})-(\d{2})-(\d{2})/))) return autDia_(+x[1], +x[2] - 1, +x[3]);
  if ((x = t.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/))) return autDia_(+x[3] < 100 ? 2000 + +x[3] : +x[3], +x[2] - 1, +x[1]);
  return null;
}
function autHoja_(nombre) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
  if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getDataRange().getValues(), enc = v.shift().map(function (h) { return String(h).trim(); });
  return v.filter(function (f) { return f.some(function (c) { return String(c).trim() !== ''; }); })
    .map(function (f) { var o = {}; enc.forEach(function (h, i) { o[h] = f[i]; }); return o; });
}
function autPer_(y, m) { return y + '-' + ('0' + (m + 1)).slice(-2); }

function autDatos_(hoy) {
  return { tareas: autTareas_(hoy), cobranza: autCobranza_(hoy), renovaciones: autRenovaciones_(hoy), cumples: autCumples_(hoy), companias: autImpagosCompania_(hoy) };
}

function autTareas_(hoy) {
  return autHoja_('Pendientes').filter(function (o) { return String(o['ID']).trim() && String(o['Título']).trim() && String(o['Estado']).trim() !== 'Hecha'; })
    .map(function (o) {
      return { titulo: String(o['Título']).trim(), responsable: String(o['Responsable']).trim(), fecha: autFecha_(o['Fecha']),
               importante: autSi_(o['Importante'], false), notas: String(o['Notas'] || '').trim() };
    })
    .filter(function (t) { return t.fecha && t.fecha <= hoy; })
    .sort(function (a, b) { return (b.importante - a.importante) || (a.fecha - b.fecha); });
}

function autCobranza_(hoy) {
  var est = {};
  autHoja_('CuponesEstado').forEach(function (o) {
    var per = (String(o['Periodo']).match(/(\d{4})-(\d{2})/) || [])[0];
    if (o['Periodo'] instanceof Date) per = Utilities.formatDate(o['Periodo'], AUT_TZ, 'yyyy-MM');
    if (per) est[String(o['ID']) + '|' + per] = { aviso: autSi_(o['Aviso'], false), pagado: autSi_(o['Pagado'], false) };
  });
  var cupones = autHoja_('Cupones').filter(function (o) { return String(o['ID']).trim() && autSi_(o['Activo'], true); });
  var y = hoy.getUTCFullYear(), m = hoy.getUTCMonth(), meses = [[m === 0 ? y - 1 : y, m === 0 ? 11 : m - 1], [y, m]];
  var impagos = [], vencenHoy = [];
  meses.forEach(function (ym) {
    var per = autPer_(ym[0], ym[1]);
    cupones.forEach(function (o) {
      var alta = autFecha_(o['Alta']);
      if (alta && autPer_(alta.getUTCFullYear(), alta.getUTCMonth()) > per) return;
      var dia = Math.max(1, Math.min(31, parseInt(o['Día'], 10) || 1));
      var fecha = autDia_(ym[0], ym[1], Math.min(dia, autDiasMes_(ym[0], ym[1])));
      var st = est[String(o['ID']) + '|' + per] || { aviso: false, pagado: false };
      if (!st.pagado && !st.aviso && alta && fecha < alta) return;
      if (st.pagado) return;
      var item = { cliente: String(o['Cliente']).trim(), compania: String(o['Compañía'] || '').trim(), ramo: String(o['Ramo'] || '').trim(), fecha: fecha, aviso: st.aviso };
      if (fecha < hoy) impagos.push(item); else if (+fecha === +hoy) vencenHoy.push(item);
    });
  });
  impagos.sort(function (a, b) { return a.fecha - b.fecha; });
  return { impagos: impagos, vencenHoy: vencenHoy };
}

function autRenovaciones_(hoy) {
  var est = {};
  autHoja_('RenovacionesEstado').forEach(function (o) {
    var f = autFecha_(o['Fecha']); if (f && autSi_(o['Enviada'], false)) est[String(o['ID']) + '|' + autISO_(f)] = true;
  });
  var desde = autSumar_(hoy, -AUT_DIAS_ATRAS), hasta = autSumar_(hoy, AUT_DIAS_ENVIO), out = [];
  autHoja_('Vencimientos').forEach(function (o) {
    var id = String(o['ID']).trim(), vence = autFecha_(o['Vencimiento']);
    if (!id || !vence || !String(o['Cliente']).trim() || !autSi_(o['Activo'], true)) return;
    var meses = AUT_VIGENCIAS[String(o['Vigencia']).trim().toLowerCase()] || 0;
    for (var d = autDia_(desde.getUTCFullYear(), desde.getUTCMonth(), 1); d <= hasta; d = autDia_(d.getUTCFullYear(), d.getUTCMonth() + 1, 1)) {
      var y = d.getUTCFullYear(), m = d.getUTCMonth(), dif = (y - vence.getUTCFullYear()) * 12 + (m - vence.getUTCMonth());
      if (dif < 0 || (meses ? dif % meses !== 0 : dif !== 0)) continue;
      var fecha = autDia_(y, m, Math.min(vence.getUTCDate(), autDiasMes_(y, m)));
      if (fecha < desde || fecha > hasta || est[id + '|' + autISO_(fecha)]) continue;
      out.push({ cliente: String(o['Cliente']).trim(), compania: String(o['Compañía'] || '').trim(), riesgo: String(o['Riesgo'] || '').trim(), fecha: fecha });
    }
  });
  return out.sort(function (a, b) { return a.fecha - b.fecha; });
}

function autCumples_(hoy) {
  return autHoja_('Clientes').filter(function (o) {
    var n = autFecha_(o['Nacimiento']);
    return n && String(o['Nombre']).trim() && n.getUTCDate() === hoy.getUTCDate() && n.getUTCMonth() === hoy.getUTCMonth();
  }).map(function (o) { return String(o['Nombre']).trim(); });
}

// Cuotas impagas que informan las compañías (hoja ImpagosCompania, se carga desde "Compañías" en el sistema)
function autImpagosCompania_(hoy) {
  return autHoja_('ImpagosCompania').map(function (o) {
    return { cliente: String(o['Cliente']).trim(), compania: String(o['Compañía']).trim(), ramo: String(o['Ramo'] || '').trim(),
             poliza: String(o['Póliza'] || '').trim(), fecha: autFecha_(o['Vencimiento']) };
  }).filter(function (o) { return o.cliente && o.fecha && o.fecha < hoy; })
    .sort(function (a, b) { return b.fecha - a.fecha; });
}

// ── Armado del mail ───────────────────────────────────────────────────────
function autEsc_(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }

function autSeccion_(titulo, cantidad, filas, link, textoLink, maximo) {
  if (!cantidad) return '';
  var max = maximo || 12, li = filas.slice(0, max).map(function (f) {
    return '<tr><td style="padding:7px 0;border-bottom:1px solid #eef0f4;font-size:14px;color:#1f2937">' + f + '</td></tr>';
  }).join('');
  if (filas.length > max) li += '<tr><td style="padding:7px 0;font-size:13px;color:#6b7280">y ' + (filas.length - max) + ' más</td></tr>';
  return '<div style="margin:0 0 22px">' +
    '<div style="font-size:13px;font-weight:600;color:#1e3a5f;text-transform:uppercase;letter-spacing:.04em;margin:0 0 6px">' +
    autEsc_(titulo) + ' <span style="background:#e8eef7;color:#1e3a5f;border-radius:10px;padding:1px 8px;font-size:12px">' + cantidad + '</span></div>' +
    '<table style="width:100%;border-collapse:collapse">' + li + '</table>' +
    (link ? '<a href="' + AUT_URL + link + '" style="display:inline-block;margin-top:8px;font-size:13px;color:#1e3a5f;font-weight:600;text-decoration:none">' + autEsc_(textoLink) + ' →</a>' : '') +
    '</div>';
}

function autResumenDe_(p, d, hoy) {
  var gris = function (t) { return ' <span style="color:#6b7280">· ' + autEsc_(t) + '</span>'; };
  var tareas = d.tareas.filter(function (t) { return t.responsable === p.nombre || !t.responsable; });
  var filasTar = tareas.map(function (t) {
    return (t.importante ? '<span style="color:#b45309">★</span> ' : '') + autEsc_(t.titulo) +
      (t.fecha < hoy ? gris('atrasada desde el ' + autCorta_(t.fecha)) : '') + (!t.responsable ? gris('sin responsable') : '');
  });
  var cob = p.cobranza ? d.cobranza : { impagos: [], vencenHoy: [] };
  var ren = p.renovaciones ? d.renovaciones : [];
  var filasImp = cob.impagos.map(function (c) { return autEsc_(c.cliente) + gris([c.compania, 'venció el ' + autCorta_(c.fecha), c.aviso ? 'avisado' : 'sin avisar'].filter(String).join(' · ')); });
  var filasHoy = cob.vencenHoy.map(function (c) { return autEsc_(c.cliente) + gris([c.compania, c.ramo, c.aviso ? 'avisado' : 'sin avisar'].filter(String).join(' · ')); });
  var filasRen = ren.map(function (r) { return autEsc_(r.cliente) + gris([r.compania, r.riesgo, (r.fecha < hoy ? 'venció el ' : 'vence el ') + autCorta_(r.fecha)].filter(String).join(' · ')); });
  var cia = p.cobranza ? (d.companias || []) : [];
  var filasCia = cia.map(function (c) { return autEsc_(c.cliente) + gris([c.compania, c.ramo + ' ' + c.poliza, 'venció el ' + autCorta_(c.fecha)].join(' · ')); });
  var filasCum = d.cumples.map(function (n) { return autEsc_(n); });

  var cuerpo = autSeccion_('Tus tareas de hoy', tareas.length, filasTar, 'tareas.html', 'Abrir Mi día') +
    autSeccion_('Cupones que vencen hoy', cob.vencenHoy.length, filasHoy, 'gestion_polizas.html#avisar', 'Avisar por WhatsApp') +
    autSeccion_('Impagos a reclamar', cob.impagos.length, filasImp, 'gestion_polizas.html#impagos', 'Ver impagos') +
    autSeccion_('Cuotas impagas en compañías', cia.length, filasCia, 'cartera.html#impagos', 'Ver y reclamar') +
    autSeccion_('Pólizas para enviar', ren.length, filasRen, 'renovaciones.html', 'Abrir Renovaciones') +
    autSeccion_('Cumpleaños de clientes', d.cumples.length, filasCum, 'clientes.html#cumples', 'Saludar');
  if (!cuerpo) cuerpo = '<p style="font-size:15px;color:#1f2937">Todo al día: no hay nada pendiente para hoy.</p>';

  var fecha = AUT_DIAS[hoy.getUTCDay()] + ' ' + hoy.getUTCDate() + ' de ' + AUT_MESES[hoy.getUTCMonth()];
  var partes = [];
  if (tareas.length) partes.push(tareas.length + (tareas.length === 1 ? ' tarea' : ' tareas'));
  if (cob.impagos.length) partes.push(cob.impagos.length + (cob.impagos.length === 1 ? ' impago' : ' impagos'));
  if (cob.vencenHoy.length) partes.push(cob.vencenHoy.length + ' vence' + (cob.vencenHoy.length === 1 ? '' : 'n') + ' hoy');
  if (cia.length) partes.push(cia.length + ' impagas en compañías');
  if (ren.length) partes.push(ren.length + ' para enviar');
  var asunto = 'Tu día · ' + fecha + (partes.length ? ' — ' + partes.join(', ') : ' — todo al día');

  var html = '<div style="background:#f5f6f8;padding:24px 12px;font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif">' +
    '<div style="max-width:560px;margin:0 auto;background:#ffffff;border-radius:12px;padding:26px 24px;border:1px solid #e5e7eb">' +
    '<div style="font-size:12px;color:#6b7280;margin:0 0 2px">Seguros Dragone</div>' +
    '<div style="font-size:21px;font-weight:600;color:#111827;margin:0 0 4px">Buen día, ' + autEsc_(p.nombre) + '</div>' +
    '<div style="font-size:14px;color:#6b7280;margin:0 0 22px">' + fecha.charAt(0).toUpperCase() + fecha.slice(1) + '</div>' +
    cuerpo +
    '<a href="' + AUT_URL + '" style="display:inline-block;margin-top:6px;background:#1e3a5f;color:#ffffff;text-decoration:none;font-size:14px;font-weight:600;padding:10px 18px;border-radius:8px">Abrir el sistema</a>' +
    '</div><div style="max-width:560px;margin:10px auto 0;font-size:11px;color:#9ca3af;text-align:center">Resumen automático. Para dejar de recibirlo, poné "No" en la hoja Equipo.</div></div>';
  return { asunto: asunto, html: html };
}
