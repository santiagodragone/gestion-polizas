/* IMPORTAR — Seguros Dragone
   Lectura de listados de las compañías (Excel, CSV o filas pegadas) y reconocimiento de columnas.
   Queda en el objeto global IMP. Lo usa Clientes. */
(function () {
  function norm(s) { return String(s == null ? '' : s).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim(); }

  function cargarLector() {                 // la librería que lee Excel se baja solo cuando hace falta
    if (window.XLSX) return Promise.resolve();
    return new Promise(function (ok, mal) {
      var s = document.createElement('script');
      s.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
      s.onload = ok; s.onerror = function () { mal(new Error('No se pudo bajar el lector de Excel. Revisá la conexión.')); };
      document.head.appendChild(s);
    });
  }

  // Devuelve la tabla (filas de celdas) de la hoja con más filas del archivo
  async function tablaDeArchivo(archivo) {
    await cargarLector();
    var libro = XLSX.read(await archivo.arrayBuffer(), { type: 'array' });
    var hojas = libro.SheetNames.map(function (n) { return XLSX.utils.sheet_to_json(libro.Sheets[n], { header: 1, raw: true, defval: '' }); });
    return hojas.sort(function (a, b) { return b.length - a.length; })[0] || [];
  }
  function tablaDeTexto(texto) { return texto.replace(/\r/g, '').split('\n').map(function (l) { return l.split('\t'); }); }

  // ¿El título de la columna contiene esa palabra entera? ("cia" no tiene que coincidir con "vigencia")
  function contiene(titulo, palabra) {
    return (' ' + norm(titulo).replace(/[^a-z0-9°]+/g, ' ') + ' ').indexOf(' ' + palabra.replace(/[^a-z0-9°]+/g, ' ').trim() + ' ') >= 0;
  }

  // campos: [{id, nombre, claves:[títulos posibles]}]. Encuentra la fila de títulos y propone qué columna es cada dato.
  function reconocer(tabla, campos) {
    tabla = tabla.filter(function (f) { return f.some(function (c) { return String(c).trim() !== ''; }); });
    if (tabla.length < 2) return null;
    function puntos(f) { return f.filter(function (c) { return campos.some(function (k) { return k.claves.some(function (p) { return contiene(c, p); }); }); }).length; }
    var iEnc = 0, mejor = -1;
    tabla.slice(0, 15).forEach(function (f, i) { var p = puntos(f); if (p > mejor) { mejor = p; iEnc = i; } });
    var cols = tabla[iEnc].map(function (c, i) { return String(c).trim() || 'Columna ' + (i + 1); });
    var mapa = {}, usadas = {};
    campos.forEach(function (k) {
      var col = -1, p;
      for (p = 0; p < k.claves.length && col < 0; p++) col = cols.findIndex(function (c, i) { return !usadas[i] && norm(c) === k.claves[p]; });
      for (p = 0; p < k.claves.length && col < 0; p++) col = cols.findIndex(function (c, i) { return !usadas[i] && contiene(c, k.claves[p]); });
      if (col >= 0) usadas[col] = true;
      mapa[k.id] = col;
    });
    return { cols: cols, filas: tabla.slice(iEnc + 1), mapa: mapa };
  }

  // Fecha de una celda: número de serie de Excel, fecha, 10/11/2026 o 2026-11-10
  function fecha(v) {
    if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
    if (typeof v === 'number' && v > 3000 && v < 90000) {
      var d = new Date(Date.UTC(1899, 11, 30) + Math.floor(v) * 86400000);
      return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
    }
    return DS.parseFecha(String(v == null ? '' : v).trim().split(' ')[0]);
  }

  window.IMP = { tablaDeArchivo: tablaDeArchivo, tablaDeTexto: tablaDeTexto, reconocer: reconocer, fecha: fecha, norm: norm };
})();
