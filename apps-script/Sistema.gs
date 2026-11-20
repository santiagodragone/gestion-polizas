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
  tareas: { nombre: 'Pendientes',         cols: ['ID', 'Título', 'Responsable', 'Fecha', 'Repite', 'Estado', 'FechaHecha', 'Importante', 'Notas', 'Creada'] },
  clientes: { nombre: 'Clientes',         cols: ['ID', 'Nombre', 'Documento', 'CUIT', 'WhatsApp', 'Teléfono', 'Mail', 'Dirección', 'Localidad', 'Nacimiento', 'Productor', 'Notas', 'Alta'] }
};

function sisFilaCliente_(c, id) {
  return [id, String(c.nombre), String(c.documento || ''), String(c.cuit || ''), String(c.whatsapp || ''), String(c.telefono || ''),
    String(c.mail || ''), String(c.direccion || ''), String(c.localidad || ''), String(c.nacimiento || ''), String(c.productor || ''),
    String(c.notas || ''), String(c.alta || '')];
}
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
  var acciones = ['asignar_ids_venc', 'guardar_venc', 'importar_venc', 'estado_venc', 'guardar_tarea', 'eliminar_tarea',
                  'guardar_cliente', 'eliminar_cliente', 'importar_clientes',
                  'importar_comisiones', 'cerrar_mes', 'reabrir_mes', 'importar_compania'];
  if (!data || acciones.indexOf(data.action) === -1) return null;

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var _com = comisionesPost_(data); if (_com) return _com;
    var _cia = companiaPost_(data); if (_cia) return _cia;

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

    if (data.action === 'guardar_cliente') {
      var c = data.cliente || {};
      if (!c.id || !c.nombre) return sisJson_({ ok: false, error: 'Falta el nombre del cliente' });
      sisUpsert_(sisHoja_(SIS_HOJAS.clientes), c.id, sisFilaCliente_(c, String(c.id)));
      return sisJson_({ ok: true });
    }

    if (data.action === 'eliminar_cliente') { sisBorrar_(sisHoja_(SIS_HOJAS.clientes), data.id); return sisJson_({ ok: true }); }

    if (data.action === 'importar_clientes') {
      // Carga masiva: actualiza los clientes que traen un ID existente y agrega el resto de una sola vez
      var shC = sisHoja_(SIS_HOJAS.clientes), listaC = data.clientes || [];
      var nC = shC.getLastRow(), filaC = {};
      if (nC >= 2) { var idsC = shC.getRange(2, 1, nC - 1, 1).getValues(); for (var iC = 0; iC < idsC.length; iC++) filaC[String(idsC[iC][0])] = iC + 2; }
      var altas = [], cambios = 0, baseC = new Date().getTime();
      listaC.forEach(function (q, k) {
        if (!q.nombre) return;
        var existe = q.id && filaC[String(q.id)];
        var fila = sisFilaCliente_(q, existe ? String(q.id) : 'CLI' + baseC + '_' + k);
        if (existe) { shC.getRange(existe, 1, 1, fila.length).setValues([fila]); cambios++; }
        else altas.push(fila);
      });
      if (altas.length) shC.getRange(shC.getLastRow() + 1, 1, altas.length, altas[0].length).setValues(altas);
      return sisJson_({ ok: true, nuevas: altas.length, actualizadas: cambios });
    }
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

// ═══════════════════════════════════════════════════════════════════════════
// COMISIONES — importar la liquidación de una compañía y cerrar / reabrir un mes
// ═══════════════════════════════════════════════════════════════════════════

var COM_CIERRE = { nombre: 'ComisionesCierre', cols: ['Clave', 'Periodo', 'Firma', 'Productor', 'Origen', 'Comisión', 'Cerrado'] };

/** Saca de una hoja todas las filas cuyo Periodo (columna B) es el indicado. */
function comQuitarPeriodo_(sh, periodo) {
  var n = sh.getLastRow(), cols = sh.getLastColumn();
  if (n < 2) return;
  var datos = sh.getRange(2, 1, n - 1, cols).getValues();
  var quedan = datos.filter(function (f) { return String(f[1]) !== String(periodo); });
  if (quedan.length === datos.length) return;
  sh.getRange(2, 1, n - 1, cols).clearContent();
  if (quedan.length) sh.getRange(2, 1, quedan.length, cols).setValues(quedan);
}

/** Devuelve null si la acción no es de comisiones. */
function comisionesPost_(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (data.action === 'importar_comisiones') {
    // Filas en el orden de la hoja Comisiones: Año | Mes | Compañía | Asegurado | Ramo | Póliza |
    // Fecha emisión | Fecha vigencia | Moneda | Prima | Premio | Comisión total | Cobrador
    var sh = ss.getSheetByName('Comisiones');
    if (!sh) return sisJson_({ ok: false, error: 'No existe la hoja Comisiones' });
    var filas = (data.filas || []).filter(function (f) { return f && f.length === 13 && String(f[3]).trim() !== ''; });
    if (filas.length) sh.getRange(sh.getLastRow() + 1, 1, filas.length, 13).setValues(filas);
    return sisJson_({ ok: true, agregadas: filas.length });
  }

  if (data.action === 'cerrar_mes') {
    // Guarda a quién quedó asignada cada comisión del mes, para que cambios posteriores no la modifiquen
    var shC = sisHoja_(COM_CIERRE), periodo = String(data.periodo || '');
    if (!/^\d{4}-\d{2}$/.test(periodo)) return sisJson_({ ok: false, error: 'Mes inválido' });
    comQuitarPeriodo_(shC, periodo);
    var hoy = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    var filasC = (data.filas || []).map(function (f) {
      return [periodo + '|' + f.firma, periodo, String(f.firma), String(f.productor || ''), String(f.origen || ''), Number(f.comision) || 0, hoy];
    });
    if (filasC.length) shC.getRange(shC.getLastRow() + 1, 1, filasC.length, filasC[0].length).setValues(filasC);
    return sisJson_({ ok: true, guardadas: filasC.length, fecha: hoy });
  }

  if (data.action === 'reabrir_mes') {
    comQuitarPeriodo_(sisHoja_(COM_CIERRE), String(data.periodo || ''));
    return sisJson_({ ok: true });
  }
  return null;
}

// ── Cartera de compañías (interfaz de Federación Patronal) ───────────────────
// El sistema lee el .zip en la computadora y manda acá solo lo que hay que guardar.
// Impagos y siniestros se reemplazan enteros por compañía + código de productor en cada importación.
var CIA_HOJAS = {
  impagos:    { nombre: 'ImpagosCompania',    cols: ['Compañía', 'Código', 'Sección', 'Ramo', 'Póliza', 'Certificado', 'Cliente', 'WhatsApp', 'Cuota', 'Vencimiento', 'Premio', 'Recibo', 'Actualizado'] },
  siniestros: { nombre: 'SiniestrosCompania', cols: ['Compañía', 'Código', 'Siniestro', 'Ramo', 'Póliza', 'Certificado', 'Cliente', 'Ocurrido', 'Denuncia', 'Estado', 'Fecha estado', 'Tipo', 'Lugar', 'Descripción', 'Actualizado'] },
  log:        { nombre: 'Importaciones',      cols: ['Compañía', 'Código', 'Datos al', 'Importado', 'Resumen'] }
};

function ciaReemplazar_(def, compania, codigo, filas) {
  var sh = sisHoja_(def), n = sh.getLastRow(), quedan = [];
  if (n >= 2) {
    quedan = sh.getRange(2, 1, n - 1, def.cols.length).getValues().filter(function (f) {
      return !(String(f[0]) === compania && String(f[1]) === codigo);
    });
    sh.getRange(2, 1, n - 1, def.cols.length).clearContent();
  }
  var todas = quedan.concat(filas);
  if (todas.length) sh.getRange(2, 1, todas.length, def.cols.length).setValues(todas);
  return filas.length;
}

function companiaPost_(data) {
  if (data.action !== 'importar_compania') return null;
  var cia = String(data.compania || '').trim(), cod = String(data.codigo || '').trim();
  if (!cia || !cod) return sisJson_({ ok: false, error: 'Falta la compañía o el código de productor' });
  var ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  var imp = (data.impagos || []).map(function (q) {
    return [cia, cod, String(q.seccion || ''), String(q.ramo || ''), String(q.poliza || ''), String(q.certificado || ''), String(q.cliente || ''),
      String(q.whatsapp || ''), String(q.cuota || ''), String(q.vencimiento || ''), Number(q.premio) || 0, String(q.recibo || ''), ahora];
  });
  var sin = (data.siniestros || []).map(function (q) {
    return [cia, cod, String(q.numero || ''), String(q.ramo || ''), String(q.poliza || ''), String(q.certificado || ''), String(q.cliente || ''),
      String(q.ocurrido || ''), String(q.denuncia || ''), String(q.estado || ''), String(q.fechaEstado || ''), String(q.tipo || ''),
      String(q.lugar || ''), String(q.descripcion || '').slice(0, 2000), ahora];
  });
  var r = { ok: true, impagos: ciaReemplazar_(CIA_HOJAS.impagos, cia, cod, imp), siniestros: ciaReemplazar_(CIA_HOJAS.siniestros, cia, cod, sin) };
  var log = sisHoja_(CIA_HOJAS.log);
  log.appendRow([cia, cod, String(data.datosAl || ''), ahora, String(data.resumen || '').slice(0, 500)]);
  return sisJson_(r);
}
