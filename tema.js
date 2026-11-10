/* TEMA — Seguros Dragone
   Aplica el modo claro u oscuro antes de que se dibuje la página y agrega el botón
   para cambiarlo en el encabezado. Se carga en el <head> de cada módulo. */
(function () {
  var CLAVE = 'dragone-tema';
  var usado = false;
  var SOL = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/></svg>';
  var LUNA = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 14.5A8.5 8.5 0 0 1 9.5 4a8.5 8.5 0 1 0 10.5 10.5z"/></svg>';

  function leer() { try { return localStorage.getItem(CLAVE); } catch (e) { return null; } }
  function guardar(v) { try { localStorage.setItem(CLAVE, v); } catch (e) {} }

  var tema = leer() === 'dark' ? 'dark' : 'light';
  document.documentElement.setAttribute('data-theme', tema);

  function pintar(btn) {
    btn.innerHTML = tema === 'dark' ? SOL : LUNA;
    btn.setAttribute('aria-label', tema === 'dark' ? 'Cambiar a modo claro' : 'Cambiar a modo oscuro');
    btn.title = btn.getAttribute('aria-label');
  }

  document.addEventListener('DOMContentLoaded', function () {
    var cabecera = document.querySelector('header');
    if (!cabecera) return;
    var donde = cabecera.querySelector('.header-right');
    if (!donde) {            // módulos cuyo encabezado solo tiene el botón de volver
      donde = document.createElement('div');
      donde.className = 'header-right';
      var volver = cabecera.querySelector('.back-btn');
      cabecera.appendChild(donde);
      if (volver) donde.appendChild(volver);
    }
    if (window.Chart) {      // los gráficos usan la misma letra y color de texto que el resto
      Chart.defaults.font.family = "'Inter', system-ui, sans-serif";
      Chart.defaults.color = window.temaColor('--text-2');
      Chart.defaults.borderColor = window.temaColor('--line');
    }
    var btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'tema-btn';
    pintar(btn);
    btn.addEventListener('click', function () {
      tema = tema === 'dark' ? 'light' : 'dark';
      document.documentElement.setAttribute('data-theme', tema);
      guardar(tema);
      pintar(btn);
      // Los colores de los gráficos se fijan al dibujarlos: hay que redibujarlos.
      // Si el módulo define window.temaRedibujar se usa eso; si no, se recarga la página.
      if (window.Chart) {
        Chart.defaults.color = window.temaColor('--text-2');
        Chart.defaults.borderColor = window.temaColor('--line');
      }
      if (typeof window.temaRedibujar === 'function') window.temaRedibujar();
      else if (usado) location.reload();
    });
    donde.insertBefore(btn, donde.firstChild);
  });

  // Color actual de una variable del tema, para usar en gráficos: temaColor('--accent')
  window.temaColor = function (nombre) {
    usado = true;
    return getComputedStyle(document.documentElement).getPropertyValue(nombre).trim();
  };
})();
