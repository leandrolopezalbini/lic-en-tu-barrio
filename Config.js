//Config.gs

/**Config pública para el frontend */
function getPublicConfig() {
  return {
    tiempoExamen: EXAM_CONFIG.totalTimeMinutes,
    aprobacion: EXAM_CONFIG.puntajeAprobacion
  };
}