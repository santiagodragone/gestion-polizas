/**
 * COBRANZA — Seguros Dragone
 * Cupones recurrentes (calendario) + reporte mensual de cobranza.
 *
 * INSTALACIÓN (una sola vez)
 * 1. En el Apps Script del Sheet "Sistema Dragone Seguros": Archivo nuevo > Secuencia de comandos,
 *    nombre "Cobranza", y pegar todo este archivo.
 * 2. En tu función doPost(e) existente, agregar como PRIMERA línea dentro de la función:
 *
 *        var _cob = cobranzaPost_(e); if (_cob) return _cob;
 *
 * 3. Elegir la función "setupCobranza" en el menú de arriba y tocar Ejecutar.
 *    Crea las hojas Cupones, CuponesEstado y ReporteCobranza, y pasa a Cupones
 *    las pólizas que ya estaban en la hoja Cobranzas (no borra nada de Cobranzas).
 * 4. Implementar > Administrar implementaciones > lápiz > Versión: "Nueva versión" > Implementar.
 *    (Así la URL del sistema sigue siendo la misma.)
 */

var COB_HOJAS = {
  cupones:  { nombre: 'Cupones',         cols: ['ID', 'Cliente', 'WhatsApp', 'Compañía', 'Póliza', 'Ramo', 'Día', 'Activo', 'Notas', 'Alta'] },
  estado:   { nombre: 'CuponesEstado',   cols: ['Clave', 'ID', 'Periodo', 'Aviso', 'FechaAviso', 'Pagado', 'FechaPago'] },
  reporte:  { nombre: 'ReporteCobranza', cols: ['Clave', 'Periodo', 'Compañía', 'Emitida', 'Cobrada', 'Notas'] }
};

/** Devuelve la hoja; si no existe la crea con encabezados y formato de texto. */
function cobHoja_(def) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(def.nombre);
  if (!sh) {
    sh = ss.insertSheet(def.nombre);
    // Texto plano en todo: evita que Sheets convierta "2026-11" o un teléfono en fecha/número científico
    sh.getRange(1, 1, sh.getMaxRows(), def.cols.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, def.cols.length).setValues([def.cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
    if (def.nombre === 'ReporteCobranza') sh.getRange('D2:E').setNumberFormat('#,##0.00');
    if (def.nombre === 'Cupones') sh.getRange('G2:G').setNumberFormat('0');
  }
  return sh;
}

/** Inserta o actualiza una fila buscando por el valor de la columna A. */
function cobUpsert_(sh, clave, fila) {
  var ultima = sh.getLastRow();
  var destino = ultima + 1;
  if (ultima >= 2) {
    var claves = sh.getRange(2, 1, ultima - 1, 1).getValues();
    for (var i = 0; i < claves.length; i++) {
      if (String(claves[i][0]) === String(clave)) { destino = i + 2; break; }
    }
  }
  sh.getRange(destino, 1, 1, fila.length).setValues([fila]);
  return destino;
}

function cobJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function cobHoy_() {
  return Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
}

/**
 * Atiende las acciones de cobranza. Devuelve null si la acción no es de este módulo,
 * para que el doPost original siga funcionando igual que siempre.
 */
function cobranzaPost_(e) {
  var data;
  try { data = JSON.parse(e.postData.contents); } catch (err) { return null; }
  var acciones = ['guardar_cupon', 'eliminar_cupon', 'estado_cupon', 'guardar_reporte'];
  if (!data || acciones.indexOf(data.action) === -1) return null;

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    if (data.action === 'guardar_cupon') {
      var c = data.cupon || {};
      if (!c.id || !c.cliente || !c.dia) return cobJson_({ ok: false, error: 'Faltan datos del cupón' });
      var dia = Math.max(1, Math.min(31, parseInt(c.dia, 10) || 1));
      cobUpsert_(cobHoja_(COB_HOJAS.cupones), c.id, [
        String(c.id), String(c.cliente), String(c.whatsapp || ''), String(c.compania || ''),
        String(c.poliza || ''), String(c.ramo || ''), dia,
        c.activo === false ? 'No' : 'Sí', String(c.notas || ''), String(c.alta || cobHoy_())
      ]);
      return cobJson_({ ok: true });
    }

    if (data.action === 'eliminar_cupon') {
      var shC = cobHoja_(COB_HOJAS.cupones);
      var n = shC.getLastRow();
      if (n >= 2) {
        var ids = shC.getRange(2, 1, n - 1, 1).getValues();
        for (var i = ids.length - 1; i >= 0; i--) {
          if (String(ids[i][0]) === String(data.id)) shC.deleteRow(i + 2);
        }
      }
      return cobJson_({ ok: true });
    }

    if (data.action === 'estado_cupon') {
      // Guarda el estado completo del cupón en ese mes (aviso + pagado)
      var clave = data.id + '|' + data.periodo;
      cobUpsert_(cobHoja_(COB_HOJAS.estado), clave, [
        clave, String(data.id), String(data.periodo),
        data.aviso ? 'Sí' : 'No', String(data.fechaAviso || ''),
        data.pagado ? 'Sí' : 'No', String(data.fechaPago || '')
      ]);
      return cobJson_({ ok: true });
    }

    if (data.action === 'guardar_reporte') {
      var shR = cobHoja_(COB_HOJAS.reporte);
      (data.filas || []).forEach(function (f) {
        var k = data.periodo + '|' + f.compania;
        cobUpsert_(shR, k, [k, String(data.periodo), String(f.compania),
          Number(f.emitida) || 0, Number(f.cobrada) || 0, String(f.notas || '')]);
      });
      return cobJson_({ ok: true });
    }
  } catch (err) {
    return cobJson_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (err2) {}
  }
  return null;
}

/**
 * Ejecutar UNA vez desde el editor.
 * Crea las hojas y migra las pólizas de la hoja Cobranzas a Cupones (solo si Cupones está vacía).
 */
function setupCobranza() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shC = cobHoja_(COB_HOJAS.cupones);
  cobHoja_(COB_HOJAS.estado);
  cobHoja_(COB_HOJAS.reporte);

  if (shC.getLastRow() > 1) { Logger.log('Cupones ya tiene datos: no se migra nada.'); return; }
  var vieja = ss.getSheetByName('Cobranzas');
  if (!vieja || vieja.getLastRow() < 2) { Logger.log('Hojas creadas. No hay datos en Cobranzas para migrar.'); return; }

  var datos = vieja.getDataRange().getValues();
  var enc = datos[0].map(function (h) { return String(h).trim(); });
  function col(nombre) { return enc.indexOf(nombre); }
  var iNom = col('Nombre completo'), iWpp = col('WhatsApp'), iComp = col('Compañía'),
      iPol = col('Nro. póliza'), iVen = col('Vencimiento');
  var tz = ss.getSpreadsheetTimeZone(), hoy = cobHoy_(), filas = [], sinDia = 0;

  for (var r = 1; r < datos.length; r++) {
    var f = datos[r];
    var nombre = iNom >= 0 ? String(f[iNom]).trim() : '';
    if (!nombre) continue;
    var v = iVen >= 0 ? f[iVen] : '', dia = 0;
    if (v instanceof Date) dia = parseInt(Utilities.formatDate(v, tz, 'd'), 10);
    else { var m = String(v).match(/^(\d{1,2})[\/\-]/); if (m) dia = parseInt(m[1], 10); }
    if (!dia) { dia = 1; sinDia++; }
    filas.push(['CUP' + new Date().getTime() + '_' + r, nombre,
      iWpp >= 0 ? String(f[iWpp]).replace(/\D/g, '') : '',
      iComp >= 0 ? String(f[iComp]).trim().replace(/^federaci[oó]n patronal$/i, 'Federación Patronal') : '',
      iPol >= 0 ? String(f[iPol]).trim() : '', '',
      dia, 'Sí', '', hoy]);
  }
  if (filas.length) shC.getRange(2, 1, filas.length, filas[0].length).setValues(filas);
  Logger.log('Migradas ' + filas.length + ' pólizas a Cupones.' +
    (sinDia ? ' ' + sinDia + ' no tenían fecha de vencimiento legible y quedaron con día 1: revisarlas.' : ''));
}
