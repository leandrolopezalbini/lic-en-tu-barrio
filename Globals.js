//Globals.js
const SHEETS = {
  INSCRIPCIONES: "Inscripciones",
  BARRIOS: "BarriosInstituciones",
  PERSONAL: "Personal",
  LOGS: "HistorialAccesos",
  ASISTENCIA: "HistorialAsistencia",
  PREGUNTAS: "PreguntasExamen",
  RESPUESTAS: "RespuestasExamen",
  CONFIG: "Configuracion"
};

const EXAM_CONFIG = {
  totalTimeMinutes: 30,
  puntajeAprobacion: 75,
  claveInstructor: "12345" 
};

const TOTAL_CLASES = 3;

// 📊 Columnas PErsonal
const COL_PER = {
  NOMBRE: 0,
  APELLIDO: 1,
  DNI: 2,
  PERFIL: 3,
  EMAIL: 4,
  TELEFONO: 5,
  PASSWORD: 6,
  REQUIERE_CAMBIO: 7
};

// 📊 COLUMNAS INSCRIPCIONES
const COL_INS = {
  ID: 0,
  NOMBRE: 1,
  APELLIDO: 2,
  DNI: 3,
  FECHA_NAC: 4,
  TELEFONO: 5,
  EMAIL: 6,
  CATEGORIA: 7,
  BARRIO: 8,
  INSTITUCION: 9,
  CURSADA1: 10,
  CURSADA2: 11,
  FECHA_EXAMEN: 12,
  ASISTENCIA: 13,
  NOTA: 14,
  ESTADO_TRAMITE: 15,
  OPERADOR: 16
};

// 🏫 COLUMNAS SEDES
const COL_BAR = {
  BARRIO: 0,
  INST: 1,
  DIR: 2,
  CUPO: 3,
  F1: 4,
  F2: 5,
  F_EX: 6,
  ESTADO: 7,
  ACTUALES: 8,
  HORARIO: 9
};

// 📝 preguntas y RESPUESTAS EXAMEN
const COL_PREG = {
  PREGUNTA: 0,
  OPC1: 1,
  OPC2: 2,
  OPC3: 3,
  CORRECTA: 4,
  IMAGEN: 5,   // 👈 NUEVA
  PUNTOS: 6,
  TIEMPO: 7,
  EXCLUYENTE: 8
};


const COL_RESP = {
  FECHA: 0,
  DNI: 1,
  NOMBRE: 2,
  NOTA: 3,
  ESTADO: 4,
  RESPUESTAS: 5,
  TIEMPO: 6
};