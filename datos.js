/* DATOS — Seguros Dragone
   Conexión con la planilla y reglas de negocio compartidas por Inicio, Renovaciones y Tareas.
   Todo queda dentro del objeto global DS. */
(function () {
  var SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxFydbgOT0KDGjMaiE8UrJF7LCVZH3hZzwCAO4QriMnDMjfRB5PVYzqSZMlDFwv8rt8/exec";

  var hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  var MESES = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  var DIAS = ['domingo','lunes','martes','miércoles','jueves','viernes','sábado'];

  // ── Utilidades ───────────────────────────────────────────────────────────
  function p2(n) { return String(n).padStart(2, '0'); }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function norm(s) { return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim(); }
  function esSi(v) { return ['sí', 'si', 'true', 'yes', 'x'].indexOf(String(v).trim().toLowerCase()) >= 0; }
  function fechaISO(d) { return d.getFullYear() + '-' + p2(d.getMonth() + 1) + '-' + p2(d.getDate()); }
  function fechaCorta(d) { return p2(d.getDate()) + '/' + p2(d.getMonth() + 1); }
  function fechaLarga(d) { var t = DIAS[d.getDay()] + ' ' + d.getDate() + ' de ' + MESES[d.getMonth()]; return t.charAt(0).toUpperCase() + t.slice(1); }
  function diasDelMes(y, m) { return new Date(y, m + 1, 0).getDate(); }
  function sumarDias(d, n) { var r = new Date(d); r.setDate(r.getDate() + n); return r; }
  function diasEntre(a, b) { return Math.round((b - a) / 86400000); }

  // Acepta 2026-11-10, la fecha ISO que devuelve Sheets (2026-11-10T03:00:00.000Z) y 10/11/2026.
  function parseFecha(v) {
    var s = String(v == null ? '' : v).trim(), x;
    if ((x = s.match(/^(\d{4})-(\d{2})-(\d{2})/))) return new Date(+x[1], +x[2] - 1, +x[3]);
    if ((x = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/))) return new Date(+x[3] < 100 ? 2000 + +x[3] : +x[3], +x[2] - 1, +x[1]);
    return null;
  }
  function wppNumero(v) {               // deja el número listo para wa.me (549 + área + número)
    var n = String(v || '').replace(/\D/g, '').replace(/^0+/, '');
    if (!n) return '';
    if (n.indexOf('54') === 0 && n.length >= 12) return n.indexOf('549') === 0 ? n : '549' + n.slice(2);
    if (n.length === 12) for (var k = 2; k <= 4; k++) if (n.slice(k, k + 2) === '15') { n = n.slice(0, k) + n.slice(k + 2); break; }
    return n.length === 10 ? '549' + n : n;
  }
  function wppLink(numero, texto) { var n = wppNumero(numero); return n ? 'https://wa.me/' + n + '?text=' + encodeURIComponent(texto) : ''; }
  function iniciales(nombre) { return String(nombre || '?').trim().split(/\s+/).slice(0, 2).map(function (p) { return p.charAt(0).toUpperCase(); }).join(''); }
  function nuevoId(pre) { return pre + Date.now() + Math.floor(Math.random() * 1000); }
  function toast(msg, err) {
    var t = document.getElementById('toast');
    if (!t) { t = document.createElement('div'); t.id = 'toast'; document.body.appendChild(t); }
    t.textContent = msg; t.className = 'show' + (err ? ' err' : '');
    clearTimeout(toast._t); toast._t = setTimeout(function () { t.className = ''; }, err ? 5000 : 2200);
  }
  function recordar(clave, valor) {     // preferencias de esta computadora (persona elegida, etc.)
    try { if (valor === undefined) return localStorage.getItem(clave); localStorage.setItem(clave, valor); } catch (e) { return null; }
  }

  // ── Planilla ─────────────────────────────────────────────────────────────
  // Devuelve las filas como objetos {encabezado: valor}. Si la hoja no existe, lanza error.
  async function hoja(nombre) {
    var r = await fetch(SCRIPT_URL + '?hoja=' + encodeURIComponent(nombre) + '&t=' + Date.now());
    var raw = await r.json();
    if (!Array.isArray(raw) || !raw.length) throw new Error('Hoja ' + nombre + ' no disponible');
    var enc = raw[0].map(function (h) { return String(h).trim(); });
    return raw.slice(1)
      .filter(function (f) { return f.some(function (c) { return String(c).trim() !== ''; }); })
      .map(function (f) { var o = {}; enc.forEach(function (h, i) { o[h] = f[i] === undefined ? '' : f[i]; }); return o; });
  }
  async function hojaOpcional(nombre) { try { return await hoja(nombre); } catch (e) { return null; } }
  async function post(payload) {
    var r = await fetch(SCRIPT_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain' }, body: JSON.stringify(payload) });
    var j = null; try { j = await r.json(); } catch (e) {}
    if (!j || j.ok !== true) throw new Error((j && j.error) || 'El servidor no confirmó el guardado');
    return j;
  }

  // ── Cobranza (misma regla que el módulo de Cobranza) ─────────────────────
  // Cada cupón vence todos los meses el mismo día; si el mes es más corto, el último día.
  function periodo(y, m) { return y + '-' + p2(m + 1); }
  function leerCobranza(cupones, estados) {
    var est = {};
    (estados || []).forEach(function (o) {
      var per = (String(o['Periodo']).match(/(\d{4})-(\d{2})/) || [])[0]; if (!per) return;
      est[String(o['ID']) + '|' + per] = { aviso: esSi(o['Aviso']), pagado: esSi(o['Pagado']) };
    });
    var cup = (cupones || []).filter(function (o) { return String(o['ID']).trim() !== ''; }).map(function (o) {
      return { id: String(o['ID']), cliente: String(o['Cliente']), dia: Math.max(1, Math.min(31, parseInt(o['Día'], 10) || 1)),
               activo: String(o['Activo']).trim() === '' ? true : esSi(o['Activo']), alta: parseFecha(o['Alta']) };
    });
    return { cupones: cup, estados: est };
  }
  function cobranzaDelMes(cob, y, m) {
    var per = periodo(y, m), out = [];
    cob.cupones.forEach(function (c) {
      if (!c.activo) return;
      if (c.alta && periodo(c.alta.getFullYear(), c.alta.getMonth()) > per) return;
      var fecha = new Date(y, m, Math.min(c.dia, diasDelMes(y, m)));
      var st = cob.estados[c.id + '|' + per] || { aviso: false, pagado: false };
      if (!st.pagado && !st.aviso && c.alta && fecha < c.alta) return;      // venció antes de cargarla
      out.push({ c: c, fecha: fecha, st: st });
    });
    return out;
  }
  // Lo que hay que hacer hoy en cobranza: impagos a reclamar (mes actual y anterior) y cupones que vencen hoy.
  function cobranzaHoy(cob) {
    var y = hoy.getFullYear(), m = hoy.getMonth(), ant = new Date(y, m - 1, 1);
    var todos = cobranzaDelMes(cob, ant.getFullYear(), ant.getMonth()).concat(cobranzaDelMes(cob, y, m));
    return {
      impagos: todos.filter(function (o) { return !o.st.pagado && o.fecha < hoy; }).length,
      vencenHoy: todos.filter(function (o) { return !o.st.pagado && +o.fecha === +hoy; }).length,
      sinAvisar: todos.filter(function (o) { return !o.st.pagado && !o.st.aviso && +o.fecha === +hoy; }).length
    };
  }

  // ── Renovaciones ─────────────────────────────────────────────────────────
  // Cada póliza se carga una vez con su vencimiento y su vigencia; el vencimiento se repite solo.
  var VIGENCIAS = { 'mensual': 1, 'bimestral': 2, 'trimestral': 3, 'cuatrimestral': 4, 'semestral': 6, 'anual': 12 };
  var DIAS_ENVIO = 30;      // desde cuántos días antes del vencimiento una póliza cuenta como "para enviar"
  var DIAS_ATRAS = 45;      // hasta cuántos días después de vencida se sigue reclamando el envío

  function leerRenovaciones(polizas, estados) {
    var est = {};
    (estados || []).forEach(function (o) {
      var f = parseFecha(o['Fecha']); if (!f) return;
      est[String(o['ID']) + '|' + fechaISO(f)] = { enviada: esSi(o['Enviada']), fechaEnvio: String(o['FechaEnvio'] || '').slice(0, 10) };
    });
    var lista = (polizas || []).map(function (o) {
      var vig = norm(o['Vigencia']);
      return { id: String(o['ID']).trim(), cliente: String(o['Cliente']).trim(), whatsapp: String(o['WhatsApp']).trim(),
               compania: String(o['Compañía']).trim(), poliza: String(o['Póliza']).trim(), riesgo: String(o['Riesgo']).trim(),
               vence: parseFecha(o['Vencimiento']), vigencia: String(o['Vigencia']).trim(), meses: VIGENCIAS[vig] || 0,
               activo: String(o['Activo']).trim() === '' ? true : esSi(o['Activo']), notas: String(o['Notas']).trim() };
    }).filter(function (p) { return p.cliente; });
    return { polizas: lista, estados: est, sinId: lista.filter(function (p) { return !p.id; }).length };
  }
  // Vencimientos de un mes. Una póliza sin vigencia cargada vence una sola vez, en su fecha.
  function renovacionesDelMes(ren, y, m) {
    var out = [], limite = sumarDias(hoy, DIAS_ENVIO);
    ren.polizas.forEach(function (p) {
      if (!p.activo || !p.vence || !p.id) return;
      var dif = (y - p.vence.getFullYear()) * 12 + (m - p.vence.getMonth());
      if (dif < 0 || (p.meses ? dif % p.meses !== 0 : dif !== 0)) return;
      var fecha = new Date(y, m, Math.min(p.vence.getDate(), diasDelMes(y, m)));
      var st = ren.estados[p.id + '|' + fechaISO(fecha)] || { enviada: false, fechaEnvio: '' };
      var estado = st.enviada ? 'enviada' : fecha < hoy ? 'vencida' : fecha <= limite ? 'enviar' : 'futura';
      out.push({ p: p, fecha: fecha, dia: fecha.getDate(), clave: fechaISO(fecha), st: st, estado: estado });
    });
    return out.sort(function (a, b) { return a.fecha - b.fecha || a.p.cliente.localeCompare(b.p.cliente, 'es'); });
  }
  // Pólizas que hay que enviar ahora: sin enviar, que vencen en los próximos DIAS_ENVIO días o vencieron hace poco.
  function renovacionesParaEnviar(ren) {
    var desde = sumarDias(hoy, -DIAS_ATRAS), hasta = sumarDias(hoy, DIAS_ENVIO), out = [];
    for (var d = new Date(desde.getFullYear(), desde.getMonth(), 1); d <= hasta; d = new Date(d.getFullYear(), d.getMonth() + 1, 1)) {
      renovacionesDelMes(ren, d.getFullYear(), d.getMonth()).forEach(function (o) {
        if (!o.st.enviada && o.fecha >= desde && o.fecha <= hasta) out.push(o);
      });
    }
    return out.sort(function (a, b) { return a.fecha - b.fecha; });
  }

  // ── Tareas ───────────────────────────────────────────────────────────────
  var EQUIPO = ['Santiago', 'Bruno', 'Leo', 'Roberto'];
  var REPITE = ['No', 'Todos los días', 'Todas las semanas', 'Todos los meses'];

  function leerTareas(filas) {
    return (filas || []).filter(function (o) { return String(o['ID']).trim() !== '' && String(o['Título']).trim() !== ''; }).map(function (o) {
      return { id: String(o['ID']), titulo: String(o['Título']), responsable: String(o['Responsable']).trim(),
               fecha: parseFecha(o['Fecha']), repite: REPITE.indexOf(String(o['Repite']).trim()) > 0 ? String(o['Repite']).trim() : 'No',
               hecha: String(o['Estado']).trim() === 'Hecha', fechaHecha: parseFecha(o['FechaHecha']),
               importante: esSi(o['Importante']), notas: String(o['Notas']), creada: String(o['Creada']).slice(0, 10) };
    });
  }
  function tareaAFila(t) {
    return { id: t.id, titulo: t.titulo, responsable: t.responsable, fecha: t.fecha ? fechaISO(t.fecha) : '', repite: t.repite,
             estado: t.hecha ? 'Hecha' : 'Pendiente', fechaHecha: t.fechaHecha ? fechaISO(t.fechaHecha) : '',
             importante: !!t.importante, notas: t.notas || '', creada: t.creada || fechaISO(new Date()) };
  }
  // Próxima fecha de una tarea que se repite (las diarias saltean sábado y domingo).
  function proximaFecha(t) {
    var base = t.fecha && t.fecha > hoy ? t.fecha : hoy, d;
    if (t.repite === 'Todos los días') { d = sumarDias(base, 1); while (d.getDay() === 0 || d.getDay() === 6) d = sumarDias(d, 1); return d; }
    if (t.repite === 'Todas las semanas') return sumarDias(t.fecha || hoy, 7 * Math.max(1, Math.ceil((diasEntre(t.fecha || hoy, hoy) + 1) / 7)));
    if (t.repite === 'Todos los meses') {
      var o = t.fecha || hoy, y = o.getFullYear(), m = o.getMonth();
      do { m++; d = new Date(y, m, Math.min(o.getDate(), diasDelMes(y, m))); } while (d <= hoy);
      return d;
    }
    return null;
  }
  // Marca o desmarca una tarea. Si se repite, al completarla pasa sola a su próxima fecha.
  function completar(t, hecha) {
    if (hecha && t.repite !== 'No') { t.fechaHecha = new Date(hoy); t.fecha = proximaFecha(t); t.hecha = false; }
    else if (!hecha && t.repite !== 'No') { t.fechaHecha = null; t.fecha = new Date(hoy); t.hecha = false; }   // se desmarcó: vuelve a hoy
    else { t.hecha = hecha; t.fechaHecha = hecha ? new Date(hoy) : null; }
    return t;
  }
  function grupoDe(t) {
    if (t.hecha) return 'hechas';
    if (!t.fecha) return 'sinfecha';
    if (t.fecha < hoy) return 'atrasadas';
    if (+t.fecha === +hoy) return 'hoy';
    var finSemana = sumarDias(hoy, (7 - hoy.getDay()) % 7);      // hasta el domingo de esta semana
    return t.fecha <= finSemana ? 'semana' : 'despues';
  }
  function hechaHoy(t) { return t.fechaHecha && +t.fechaHecha === +hoy; }

  window.DS = {
    SCRIPT_URL: SCRIPT_URL, hoy: hoy, MESES: MESES, DIAS: DIAS, EQUIPO: EQUIPO, REPITE: REPITE,
    VIGENCIAS: Object.keys(VIGENCIAS), DIAS_ENVIO: DIAS_ENVIO,
    p2: p2, esc: esc, norm: norm, esSi: esSi, fechaISO: fechaISO, fechaCorta: fechaCorta, fechaLarga: fechaLarga, parseFecha: parseFecha,
    diasDelMes: diasDelMes, sumarDias: sumarDias, diasEntre: diasEntre, wppLink: wppLink, wppNumero: wppNumero,
    iniciales: iniciales, nuevoId: nuevoId, toast: toast, recordar: recordar,
    hoja: hoja, hojaOpcional: hojaOpcional, post: post,
    leerCobranza: leerCobranza, cobranzaHoy: cobranzaHoy,
    leerRenovaciones: leerRenovaciones, renovacionesDelMes: renovacionesDelMes, renovacionesParaEnviar: renovacionesParaEnviar,
    leerTareas: leerTareas, tareaAFila: tareaAFila, completar: completar, grupoDe: grupoDe, hechaHoy: hechaHoy
  };
})();
