//Config.gs
const EXAM_CONFIG = {
  totalTimeMinutes: 30,
  puntajeAprobacion: 75,
  claveInstructor: "12345" 
};

const TOTAL_CLASES = 3;


// CONFIGURACIÓN DE COLUMNAS DE LA HOJA RESPUESTAS
const COL_RESP = {
  FECHA: 0,
  DNI: 1,
  NOMBRE: 2,
  NOTA: 3,
  ESTADO: 4,
  RESPUESTAS: 5,
  TIEMPO: 6
};

// Columnas de la hoja INSCRIPCIONES
const COL_INS = {
  ID: 0,              // A
  NOMBRE: 1,          // B
  APELLIDO: 2,        // C
  DNI: 3,             // D
  FECHA_NAC: 4,       // E
  TELEFONO: 5,        // F
  EMAIL: 6,           // G
  CATEGORIA: 7,       // H
  BARRIO: 8,          // I
  INSTITUCION: 9,     // J
  CURSADA1: 10,       // K
  CURSADA2: 11,       // L
  FECHA_EXAMEN: 12,   // M
  ASISTENCIA: 13,     // N
  NOTA: 14,           // O
  ESTADO_TRAMITE: 15, // P
  OPERADOR: 16        // Q
};

/**
 * Config pública para el frontend
 */
function getPublicConfig() {
  return {
    tiempoExamen: EXAM_CONFIG.totalTimeMinutes,
    aprobacion: EXAM_CONFIG.puntajeAprobacion
  };
}