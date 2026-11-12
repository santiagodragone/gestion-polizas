/**
 * SEGURIDAD — Seguros Dragone
 * Sin sesión válida no se puede leer ni escribir nada de la planilla.
 *
 * Cómo funciona
 * - La clave se define desde la planilla: menú "Sistema Dragone" > "Definir clave de acceso".
 *   Se guarda cifrada (no se puede leer), en las propiedades del script.
 * - El sistema manda la clave al iniciar sesión y recibe un pase firmado que vence a las 12 horas.
 * - Mientras no haya una clave definida, la protección está apagada (todo funciona como antes).
 *
 * Instalación: primera línea de doGet(e)   →  var _seg = seguridadGet_(e);  if (_seg) return _seg;
 *              primera línea de doPost(e)  →  var _seg = seguridadPost_(e); if (_seg) return _seg;
 */

var SEG_HORAS = 12;        // duración de la sesión
var SEG_INTENTOS = 8;      // claves equivocadas seguidas antes de bloquear 10 minutos

function segProps_() { return PropertiesService.getScriptProperties(); }
function segJson_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function segActiva_() { return !!segProps_().getProperty('SEG_CLAVE_HASH'); }

function segHash_(sal, clave) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, sal + '|' + clave, Utilities.Charset.UTF_8));
}
function segFirma_(texto) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(texto, segProps_().getProperty('SEG_SECRETO')));
}
function segPase_() {
  var vence = new Date().getTime() + SEG_HORAS * 3600000;
  return vence + '.' + segFirma_(String(vence));
}
function segValido_(pase) {
  var p = String(pase || '').split('.');
  return p.length === 2 && Number(p[0]) > new Date().getTime() && segFirma_(p[0]) === p[1];
}

/** Lecturas: devuelve null si puede seguir, o la respuesta de "sin sesión". */
function seguridadGet_(e) {
  if (!segActiva_()) return null;
  if (segValido_(e && e.parameter && e.parameter.token)) return null;
  return segJson_({ error: 'sesion' });
}

/** Escrituras e inicio de sesión. */
function seguridadPost_(e) {
  var data = null;
  try { data = JSON.parse(e.postData.contents); } catch (err) {}

  if (data && data.action === 'login') {
    if (!segActiva_()) return segJson_({ ok: false, error: 'sin_configurar' });
    var cache = CacheService.getScriptCache(), fallos = Number(cache.get('seg_fallos') || 0);
    if (fallos >= SEG_INTENTOS) return segJson_({ ok: false, error: 'bloqueado' });
    var p = segProps_();
    if (segHash_(p.getProperty('SEG_SAL'), String(data.clave || '')) === p.getProperty('SEG_CLAVE_HASH')) {
      cache.remove('seg_fallos');
      return segJson_({ ok: true, token: segPase_(), horas: SEG_HORAS });
    }
    cache.put('seg_fallos', String(fallos + 1), 600);
    return segJson_({ ok: false, error: 'clave' });
  }

  if (!segActiva_()) return null;
  if (data && segValido_(data.token)) return null;
  return segJson_({ ok: false, error: 'sesion' });
}

/** Menú en la planilla. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Sistema Dragone')
    .addItem('Definir clave de acceso', 'segDefinirClave')
    .addToUi();
}

/** Pide la clave nueva en la planilla y la guarda cifrada. Cambiarla cierra todas las sesiones abiertas. */
function segDefinirClave() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Clave de acceso al sistema',
    'Escribí la clave nueva (mínimo 8 caracteres). Es la que va a usar el equipo para entrar al sistema.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var clave = r.getResponseText().trim();
  if (clave.length < 8) { ui.alert('La clave tiene que tener al menos 8 caracteres. No se cambió nada.'); return; }
  var p = segProps_(), sal = Utilities.getUuid();
  p.setProperty('SEG_SAL', sal);
  p.setProperty('SEG_CLAVE_HASH', segHash_(sal, clave));
  p.setProperty('SEG_SECRETO', Utilities.getUuid() + Utilities.getUuid());
  ui.alert('Listo. La clave quedó guardada. Desde ahora el sistema la pide para entrar y las sesiones abiertas se cerraron.');
}
