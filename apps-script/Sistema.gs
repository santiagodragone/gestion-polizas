/**
 * SISTEMA — Seguros Dragone
 * Renovaciones (calendario de vencimientos) + Tareas ("Mi día").
 *
 * INSTALACIÓN (una sola vez)
 * 1. Archivo nuevo en el Apps Script, nombre "Sistema", y pegar todo este archivo.
 * 2. En doPost(e), debajo de la línea de cobranza, agregar:
 *        var _sis = sistemaPost_(e); if (_sis) return _sis;
 * 3. Ejecutar una vez la función "setupSistema".
 * 4. Implementar > Gestionar implementaciones > lápiz > Versión nueva > Implementar.
 */

var SIS_HOJAS = {
  venc:   { nombre: 'Vencimientos',       cols: ['ID', 'Cliente', 'WhatsApp', 'Compañía', 'Póliza', 'Riesgo', 'Vencimiento', 'Vigencia', 'Activo', 'Notas'] },
  estado: { nombre: 'RenovacionesEstado', cols: ['Clave', 'ID', 'Fecha', 'Enviada', 'FechaEnvio'] },
  tareas: { nombre: 'Pendientes',         cols: ['ID', 'Título', 'Responsable', 'Fecha', 'Repite', 'Estado', 'FechaHecha', 'Importante', 'Notas', 'Creada'] }
};
var SIS_VIGENCIAS = ['Mensual', 'Bimestral', 'Trimestral', 'Cuatrimestral', 'Semestral', 'Anual'];

function sisHoja_(def) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(def.nombre);
  if (!sh) {
    sh = ss.insertSheet(def.nombre);
    sh.getRange(1, 1, sh.getMaxRows(), def.cols.length).setNumberFormat('@');     // texto plano
    sh.getRange(1, 1, 1, def.cols.length).setValues([def.cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
    if (def.nombre === 'Vencimientos') {
      // Pensada para cargar a mano: la fecha es una fecha de verdad y la vigencia se elige de una lista
      sh.getRange('G2:G').setNumberFormat('dd/MM/yyyy');
      sh.getRange('H2:H').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(SIS_VIGENCIAS, true).setAllowInvalid(true).build());
      sh.getRange('I2:I').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Sí', 'No'], true).setAllowInvalid(true).build());
      sh.getRange('A1').setNote('No completar: el sistema asigna el ID solo.');
      sh.getRange('H1').setNote('Cada cuánto se renueva la póliza. Si queda vacío, el vencimiento no se repite.');
      sh.setColumnWidth(2, 240);
    }
  }
  return sh;
}

function sisUpsert_(sh, clave, fila) {
  var ultima = sh.getLastRow(), destino = ultima + 1;
  if (ultima >= 2) {
    var claves = sh.getRange(2, 1, ultima - 1, 1).getValues();
    for (var i = 0; i < claves.length; i++) if (String(claves[i][0]) === String(clave)) { destino = i + 2; break; }
  }
  sh.getRange(destino, 1, 1, fila.length).setValues([fila]);
}

function sisBorrar_(sh, clave) {
  var n = sh.getLastRow();
  if (n < 2) return;
  var ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === String(clave)) sh.deleteRow(i + 2);
}

function sisJson_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

/** Completa el ID de las pólizas cargadas a mano en Vencimientos. Devuelve cuántas completó. */
function sisAsignarIds_() {
  var sh = sisHoja_(SIS_HOJAS.venc), n = sh.getLastRow(), hechos = 0;
  if (n < 2) return 0;
  var datos = sh.getRange(2, 1, n - 1, 2).getValues();
  for (var i = 0; i < datos.length; i++) {
    if (String(datos[i][0]).trim() === '' && String(datos[i][1]).trim() !== '') {
      sh.getRange(i + 2, 1).setValue('REN' + new Date().getTime() + '_' + (i + 2));
      hechos++;
    }
  }
  return hechos;
}

/** Devuelve null si la acción no es de este módulo, para que doPost siga con las demás. */
function sistemaPost_(e) {
  var data;
  try { data = JSON.parse(e.postData.contents); } catch (err) { return null; }
  var acciones = ['asignar_ids_venc', 'guardar_venc', 'importar_venc', 'estado_venc', 'guardar_tarea', 'eliminar_tarea'];
  if (!data || acciones.indexOf(data.action) === -1) return null;

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (data.action === 'asignar_ids_venc') return sisJson_({ ok: true, asignados: sisAsignarIds_() });

    if (data.action === 'guardar_venc') {
      var p = data.poliza || {};
      if (!p.id || !p.cliente) return sisJson_({ ok: false, error: 'Faltan datos de la póliza' });
      var fecha = p.vencimiento ? Utilities.parseDate(String(p.vencimiento), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd') : '';
      sisUpsert_(sisHoja_(SIS_HOJAS.venc), p.id, [String(p.id), String(p.cliente), String(p.whatsapp || ''), String(p.compania || ''),
        String(p.poliza || ''), String(p.riesgo || ''), fecha, String(p.vigencia || ''), p.activo === false ? 'No' : 'Sí', String(p.notas || '')]);
      return sisJson_({ ok: true });
    }

    if (data.action === 'importar_venc') {
      // Carga masiva: actualiza las pólizas que traen un ID existente y agrega el resto de una sola vez
      var shV = sisHoja_(SIS_HOJAS.venc), lista = data.polizas || [], tz = ss.getSpreadsheetTimeZone();
      var nV = shV.getLastRow(), filaDe = {};
      if (nV >= 2) { var idsV = shV.getRange(2, 1, nV - 1, 1).getValues(); for (var iV = 0; iV < idsV.length; iV++) filaDe[String(idsV[iV][0])] = iV + 2; }
      var nuevas = [], actualizadas = 0, base = new Date().getTime();
      lista.forEach(function (q, k) {
        if (!q.cliente) return;
        var existe = q.id && filaDe[String(q.id)];
        var fila = [existe ? String(q.id) : 'REN' + base + '_' + k, String(q.cliente), String(q.whatsapp || ''), String(q.compania || ''),
          String(q.poliza || ''), String(q.riesgo || ''), q.vencimiento ? Utilities.parseDate(String(q.vencimiento), tz, 'yyyy-MM-dd') : '',
          String(q.vigencia || ''), q.activo === false ? 'No' : 'Sí', String(q.notas || '')];
        if (existe) { shV.getRange(existe, 1, 1, fila.length).setValues([fila]); actualizadas++; }
        else nuevas.push(fila);
      });
      if (nuevas.length) shV.getRange(shV.getLastRow() + 1, 1, nuevas.length, nuevas[0].length).setValues(nuevas);
      return sisJson_({ ok: true, nuevas: nuevas.length, actualizadas: actualizadas });
    }

    if (data.action === 'estado_venc') {
      var clave = data.id + '|' + data.fecha;
      sisUpsert_(sisHoja_(SIS_HOJAS.estado), clave, [clave, String(data.id), String(data.fecha), data.enviada ? 'Sí' : 'No', String(data.fechaEnvio || '')]);
      return sisJson_({ ok: true });
    }

    if (data.action === 'guardar_tarea') {
      var t = data.tarea || {};
      if (!t.id || !t.titulo) return sisJson_({ ok: false, error: 'Falta el título de la tarea' });
      sisUpsert_(sisHoja_(SIS_HOJAS.tareas), t.id, [String(t.id), String(t.titulo), String(t.responsable || ''), String(t.fecha || ''),
        String(t.repite || 'No'), t.estado === 'Hecha' ? 'Hecha' : 'Pendiente', String(t.fechaHecha || ''), t.importante ? 'Sí' : 'No',
        String(t.notas || ''), String(t.creada || '')]);
      return sisJson_({ ok: true });
    }

    if (data.action === 'eliminar_tarea') { sisBorrar_(sisHoja_(SIS_HOJAS.tareas), data.id); return sisJson_({ ok: true }); }
  } catch (err) {
    return sisJson_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (err2) {}
  }
  return null;
}

/**
 * Ejecutar UNA vez. Crea las hojas Vencimientos, RenovacionesEstado y Pendientes,
 * y pasa a Pendientes las tareas sin terminar del tablero viejo (hoja "tareas"). No borra nada.
 */
function setupSistema() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  sisHoja_(SIS_HOJAS.venc);
  sisHoja_(SIS_HOJAS.estado);
  var shT = sisHoja_(SIS_HOJAS.tareas);
  if (shT.getLastRow() > 1) { Logger.log('Pendientes ya tiene datos: no se migra nada.'); return; }

  var vieja = ss.getSheetByName('Tareas');
  if (!vieja || vieja.getLastRow() < 2) { Logger.log('Hojas creadas. No hay tareas viejas para migrar.'); return; }
  // Hoja vieja: dia | titulo | descripcion | responsable | prioridad | fecha limite | estado | notas
  var datos = vieja.getDataRange().getValues(), tz = ss.getSpreadsheetTimeZone(), hoy = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'), filas = [];
  for (var r = 1; r < datos.length; r++) {
    var f = datos[r], titulo = String(f[1]).trim();
    if (!titulo || String(f[6]).trim() === 'Hecho') continue;
    var limite = f[5] instanceof Date ? Utilities.formatDate(f[5], tz, 'yyyy-MM-dd') : String(f[5]).slice(0, 10);
    var resp = String(f[3]).trim();
    if (/^santi$/i.test(resp)) resp = 'Santiago';
    var notas = [String(f[2]).trim(), String(f[7]).trim()].filter(String).join(' — ');
    filas.push(['TAR' + new Date().getTime() + '_' + r, titulo, resp, limite, 'No', 'Pendiente', '', String(f[4]).trim() === 'Alta' ? 'Sí' : 'No', notas, hoy]);
  }
  if (filas.length) shT.getRange(2, 1, filas.length, filas[0].length).setValues(filas);
  Logger.log('Hojas creadas. Tareas migradas: ' + filas.length);
}
