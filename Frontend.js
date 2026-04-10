// Frontend.js - Código JavaScript compartido para el frontend
// Incluye funciones comunes para llamadas al servidor, manejo de UI con Materialize

/* eslint-env browser */

// 1. Inicialización centralizada
// - Materialize
// - Ruteo basado en ?p=...
if (typeof document !== 'undefined') {
  document.addEventListener('DOMContentLoaded', () => {
    M.AutoInit();
    initRoutes();
  });
}

// 2. Función de comunicación (Optimizada)
function llamarServidor(nombreFuncion, parametros = [], callbackExito) {
  const preloader = document.getElementById('preloader');
  if (preloader) preloader.style.display = 'block';

  google.script.run
    .withSuccessHandler((result) => {
      if (preloader) preloader.style.display = 'none';
      if (result.success && callbackExito) {
        callbackExito(result);
      } else {
        M.toast({html: `⚠️ ${result.message || 'Error'}`, classes: 'red'});
      }
    })
    .withFailureHandler((error) => {
      if (preloader) preloader.style.display = 'none';
      M.toast({html: '❌ Error de conexión', classes: 'red'});
      console.error("Error GAS:", error);
    })[nombreFuncion](...parametros);
}

// 3. Obtener parámetros (Moderno y sin "Citations")
const getQueryParam = (name) => {
  if (typeof window === 'undefined') return null;
  return new URLSearchParams(window.location.search).get(name);
};

// 4. Ruteo flexible basado en query params
const routeHandlers = {
  super: () => {
    // Inicializaciones específicas para la vista de Super Usuario
    // Por ejemplo: cargar datos, configurar tablas, etc.
  },
  mesa: () => {
    // Inicializaciones específicas para la vista de Mesa de Entrada
  },
};

function initRoutes() {
  const page = (getQueryParam('p') || '').toLowerCase();
  const handler = routeHandlers[page] || (() => {});
  handler();
}

// 5. Cerrar sesión limpio (limpia query params de la URL)
function cerrarSesion({ redirectTo = '/' } = {}) {
  const url = new URL(window.location.href);
  url.search = '';
  window.history.replaceState({}, document.title, url.toString());
  // Opcional: redirigir a una página de login/portada
  if (redirectTo) {
    window.location.href = redirectTo;
  }
}

// 4. Lógica de Negocio (cliente)
// Nota: No debe colisionar con la función server-side `loginPersonal`.
function loginPersonalUI() {
  const dni = document.getElementById('dni')?.value.trim();
  const pass = document.getElementById('pass')?.value.trim();

  if (!dni || !pass) return M.toast({html: 'Completa los campos', classes: 'orange'});

  llamarServidor('loginPersonal', [dni, pass], (res) => {
    window.location.href = res.perfil === 'Admin' ? '?p=super' : '?p=mesa';
  });
}

// Ejemplo específico: Procesar Examen
function enviarExamen() {
  // Recopilar respuestas: Asumiendo inputs con name="respuesta_0", "respuesta_1", etc.
  const inputs = document.querySelectorAll('input[name^="respuesta_"]');

  if (!inputs.length) {
    return M.toast({html: 'No se detectaron respuestas para procesar', classes: 'orange'});
  }

  const respuestas = typeof data[i][5] === "string"
  ? JSON.parse(data[i][5])
  : data[i][5];
  inputs.forEach(input => {
    const nameParts = (input.name || '').split('_');
    if (nameParts.length < 2) return;
    const index = nameParts[1];
    respuestas[index] = (input.value || '').toLowerCase();
  });

  const dniRaw = getQueryParam('dni');
  const dni = dniRaw ? dniRaw.trim() : '';
  if (!dni) {
    M.toast({html: 'DNI no encontrado', classes: 'red'});
    return;
  }

  llamarServidor('procesarExamenDesdeCliente', [respuestas, dni], function(result) {
    // Mostrar resultado en modal (seguro contra nulls)
    const notaFinalEl = document.getElementById('notaFinal');
    const estadoFinalEl = document.getElementById('estadoFinal');

    if (notaFinalEl) notaFinalEl.textContent = `${result?.nota ?? 0}%`;
    if (estadoFinalEl) estadoFinalEl.textContent = result?.aprobado ? 'Aprobado' : 'Desaprobado';

    const modalEl = document.getElementById('modalResultado');
    const modal = modalEl ? M.Modal.getInstance(modalEl) : null;
    if (modal) modal.open();
  });
}

// Función para mostrar modal de confirmación
function confirmarAccion(mensaje, callback) {
  const mensajeEl = document.getElementById('mensajeConfirmacion');
  if (mensajeEl) mensajeEl.textContent = mensaje;

  const modalEl = document.getElementById('modalConfirmacion');
  const modal = modalEl ? M.Modal.getInstance(modalEl) : null;
  if (modal) modal.open();
  // Asignar callback al botón confirmar (asumiendo id="btnConfirmar")
  const btnConfirmar = document.getElementById('btnConfirmar');
  if (btnConfirmar) btnConfirmar.onclick = callback;
}