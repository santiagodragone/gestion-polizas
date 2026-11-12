/* TEMA — Seguros Dragone
   Se carga en el <head> de cada módulo. Hace tres cosas:
   1. aplica el modo claro u oscuro antes de que se dibuje la página;
   2. arma el menú lateral (el mismo en todas las pantallas);
   3. expone temaColor() para que los gráficos usen los colores del tema. */
(function () {
  var CLAVE = 'dragone-tema';
  var SESION = 'sd_sesion';

  // ── Sesión ──
  // El inicio guarda acá el pase que entrega la planilla al validar la clave. Sin pase vigente,
  // cualquier pantalla vuelve al inicio, y todos los pedidos a la planilla lo llevan adjunto.
  function sesion() {
    try { var s = JSON.parse(localStorage.getItem(SESION) || 'null'); return s && s.vence > Date.now() ? s : null; } catch (e) { return null; }
  }
  function cerrarSesion() {
    try { localStorage.removeItem(SESION); } catch (e) {}
    if (esInicio) location.reload(); else location.replace('index.html');
  }
  var esInicio = /(^|\/)(index\.html)?$/i.test(location.pathname);
  if (!esInicio && !sesion()) location.replace('index.html');

  var pedir = window.fetch;
  window.fetch = function (url, opciones) {
    var u = typeof url === 'string' ? url : (url && url.url) || '';
    if (u.indexOf('script.google.com/macros/') < 0) return pedir.apply(this, arguments);
    var s = sesion(), pase = s ? s.token : '';
    opciones = opciones || {};
    if (String(opciones.method || 'GET').toUpperCase() === 'POST') {
      try {
        var d = JSON.parse(opciones.body);
        if (d && d.action !== 'login' && !d.token) { d.token = pase; opciones = Object.assign({}, opciones, { body: JSON.stringify(d) }); }
      } catch (e) {}
    } else if (pase && u.indexOf('token=') < 0) url = u + (u.indexOf('?') < 0 ? '?' : '&') + 'token=' + encodeURIComponent(pase);
    return pedir.call(this, url, opciones).then(function (r) {
      // Si la planilla dice que el pase ya no vale (venció o se cambió la clave), se vuelve a pedir la clave
      r.clone().json().then(function (j) { if (j && j.error === 'sesion' && s) cerrarSesion(); }).catch(function () {});
      return r;
    });
  };
  window.temaSesion = { leer: sesion, cerrar: cerrarSesion,
    guardar: function (token, horas) { try { localStorage.setItem(SESION, JSON.stringify({ token: token || '', vence: Date.now() + horas * 3600000 - 60000 })); } catch (e) {} } };
  var usado = false;

  function ic(d) { return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">' + d + '</svg>'; }
  var SOL = ic('<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/>');
  var LUNA = ic('<path d="M20 14.5A8.5 8.5 0 0 1 9.5 4a8.5 8.5 0 1 0 10.5 10.5z"/>');
  var SALIR = ic('<path d="M15 4h3a2 2 0 0 1 2 2v12a2 2 0 0 1-2 2h-3M10 16l-4-4 4-4M6 12h10"/>');
  var MENU = ic('<path d="M4 7h16M4 12h16M4 17h16"/>');

  // ── Menú: para sumar un módulo, agregar una línea acá ──
  var MODULOS = [
    { grupo: 'Hoy' },
    { href: 'index.html', nombre: 'Inicio', icono: ic('<path d="M4 11l8-7 8 7v8a1 1 0 0 1-1 1h-4v-6h-6v6H5a1 1 0 0 1-1-1z"/>') },
    { href: 'tareas.html', nombre: 'Tareas', icono: ic('<rect x="3" y="3" width="18" height="18" rx="4"/><path d="M8 12.5l2.8 2.8L16 9.5"/>') },
    { grupo: 'Gestión' },
    { href: 'clientes.html', nombre: 'Clientes', icono: ic('<circle cx="9" cy="8" r="3.5"/><path d="M2.5 20a6.5 6.5 0 0 1 13 0M16 4.6a3.5 3.5 0 0 1 0 6.8M18 14.5a6.5 6.5 0 0 1 3.5 5.5"/>') },
    { href: 'gestion_polizas.html', nombre: 'Cobranza', icono: ic('<rect x="2.5" y="5" width="19" height="14" rx="3"/><path d="M2.5 10h19M6.5 15h4"/>') },
    { href: 'renovaciones.html', nombre: 'Renovaciones', icono: ic('<path d="M20 11a8 8 0 0 0-14.3-4.3L4 8.5M4 4v4.5h4.5M4 13a8 8 0 0 0 14.3 4.3L20 15.5M20 20v-4.5h-4.5"/>') },
    { href: 'siniestros.html', nombre: 'Siniestros', icono: ic('<path d="M5 11l1.6-4.2A2 2 0 0 1 8.5 5.5h7a2 2 0 0 1 1.9 1.3L19 11"/><rect x="3" y="11" width="18" height="6" rx="2"/><path d="M6.5 17v2M17.5 17v2M7 14h.01M17 14h.01"/>') },
    { href: 'ventas.html', nombre: 'Ventas', icono: ic('<path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/>') },
    { grupo: 'Consulta' },
    { href: 'comisiones.html', nombre: 'Comisiones', icono: ic('<circle cx="12" cy="12" r="9"/><path d="M14.8 9.2c-.5-.9-1.5-1.4-2.8-1.4-1.7 0-2.9.9-2.9 2.1 0 2.9 5.9 1.3 5.9 4.2 0 1.2-1.2 2.1-3 2.1-1.4 0-2.5-.6-3-1.5M12 6v1.8M12 16.2V18"/>') },
    { href: 'contactos.html', nombre: 'Contactos', icono: ic('<rect x="5" y="3" width="14" height="18" rx="2"/><path d="M9 7.5h1.5M13.5 7.5H15M9 11.5h1.5M13.5 11.5H15M10 21v-4h4v4"/>') },
    { href: 'indicadores.html', nombre: 'Indicadores', icono: ic('<path d="M3 17l6-6 4 4 8-8"/><path d="M15 7h6v6"/>') }
  ];

  function leer() { try { return localStorage.getItem(CLAVE); } catch (e) { return null; } }
  function guardar(v) { try { localStorage.setItem(CLAVE, v); } catch (e) {} }

  var tema = leer() === 'dark' ? 'dark' : 'light';
  document.documentElement.setAttribute('data-theme', tema);

  function pintar(btn) {
    var claro = tema === 'dark';
    btn.innerHTML = (claro ? SOL : LUNA) + '<span>' + (claro ? 'Modo claro' : 'Modo oscuro') + '</span>';
  }

  function armarMenu() {
    var actual = (location.pathname.split('/').pop() || 'index.html').toLowerCase();
    var nav = document.createElement('nav');
    nav.className = 'nav';
    nav.setAttribute('aria-label', 'Módulos');
    var html = '<a class="nav-brand" href="index.html"><span class="nav-logo">D</span><span><strong>Seguros Dragone</strong><small>Sistema de gestión</small></span></a><div class="nav-links">';
    MODULOS.forEach(function (m) {
      if (m.grupo) html += '<div class="nav-grupo">' + m.grupo + '</div>';
      else html += '<a href="' + m.href + '"' + (m.href === actual ? ' class="active" aria-current="page"' : '') + '>' + m.icono + '<span>' + m.nombre + '</span></a>';
    });
    html += '</div><div class="nav-pie"><button type="button" class="nav-btn" id="nav-tema"></button><button type="button" class="nav-btn" id="nav-salir">' + SALIR + '<span>Salir</span></button></div>';
    nav.innerHTML = html;

    var barra = document.createElement('div');       // barra superior, solo en celular
    barra.className = 'nav-barra';
    barra.innerHTML = '<button type="button" class="nav-abrir" aria-label="Abrir menú">' + MENU + '</button><strong>Seguros Dragone</strong>';
    var fondo = document.createElement('div');
    fondo.className = 'nav-fondo';

    document.body.classList.add('con-nav');
    document.body.insertBefore(fondo, document.body.firstChild);
    document.body.insertBefore(nav, document.body.firstChild);
    document.body.insertBefore(barra, document.body.firstChild);

    function cerrar() { document.body.classList.remove('nav-abierto'); }
    barra.querySelector('.nav-abrir').addEventListener('click', function () { document.body.classList.toggle('nav-abierto'); });
    fondo.addEventListener('click', cerrar);

    var btn = nav.querySelector('#nav-tema');
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
    nav.querySelector('#nav-salir').addEventListener('click', function () {
      cerrarSesion();
    });
  }

  document.addEventListener('DOMContentLoaded', function () {
    armarMenu();
    // El encabezado viejo de cada módulo: la marca y "volver" ahora están en el menú.
    // Si no le queda ningún botón propio (Actualizar, + Nuevo...), se oculta entero.
    var cabecera = document.querySelector('header');
    if (cabecera) {
      var propios = cabecera.querySelectorAll('button, a:not(.brand):not(.back-btn), input, select');
      if (!propios.length) cabecera.classList.add('sin-acciones');
    }
    if (window.Chart) {      // los gráficos usan la misma letra y color de texto que el resto
      Chart.defaults.font.family = "'Inter', system-ui, sans-serif";
      Chart.defaults.color = window.temaColor('--text-2');
      Chart.defaults.borderColor = window.temaColor('--line');
    }
  });

  // Color actual de una variable del tema, para usar en gráficos: temaColor('--accent')
  window.temaColor = function (nombre) {
    usado = true;
    return getComputedStyle(document.documentElement).getPropertyValue(nombre).trim();
  };
})();
