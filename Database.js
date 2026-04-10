// DATABASE.GS 

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

const COL_INS = {
  ID: 0,
  NOM: 1,
  APE: 2,
  DNI: 3,
  FNAC: 4,
  TEL: 5,
  EMAIL: 6,
  CAT: 7,
  BARRIO: 8,
  INST: 9,
  F1: 10,
  F2: 11,
  F_EX: 12,
  ASIST: 13,
  NOTA: 14,
  ESTADO: 15
};

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

// --- 1. NÚCLEO Y AUDITORÍA ---

function getSheet(nombre) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(nombre);
  if (!sheet) throw new Error("No se encontró la hoja: " + nombre);
  return sheet;
}

function registrarAccion(dniOp, accion, detalle = "") {
  try {
    const sheet = getSheet(SHEETS.LOGS);
    const personal = buscarPersonaPorDni(dniOp);
    const nombreOp = personal ? `${personal.nombre} ${personal.apellido}` : "Sistema";
    sheet.appendRow([new Date(), dniOp, nombreOp, accion, detalle]);
  } catch (e) { console.error("Error Log: " + e.message); }
}

function obtenerLogs() {
  try {
    const data = getSheet(SHEETS.LOGS).getDataRange().getValues();
    if (data.length <= 1) return [];
    return data.slice(1).reverse().slice(0, 50).map(f => ({
      fecha: Utilities.formatDate(new Date(f[0]), "GMT-3", "dd/MM/yyyy HH:mm"),
      dni: f[1], operador: f[2], accion: f[3], detalle: f[4]
    }));
  } catch (e) { return []; }
}

// --- 2. GESTIÓN DE ALUMNOS ---

function buscarAlumno(query) {
  try {
    const data = getSheet(SHEETS.INSCRIPCIONES).getDataRange().getValues();
    const q = query.toString().toLowerCase().trim();
    
    return data.slice(1)
      .filter(f => {
        const dni = f[COL.DNI] ? f[COL.DNI].toString() : "";
        const ape = f[COL.APE] ? f[COL.APE].toString().toLowerCase() : "";
        return dni.includes(q) || ape.includes(q);
      })
      .slice(0, 10)
      .map(f => ({
        nombre: f[COL.NOM], apellido: f[COL.APE], dni: f[COL.DNI],
        institucion: f[COL.INST] || "Sin Sede",
        asistencia: f[COL.ASIST], nota: f[COL.NOTA], estado: f[COL.ESTADO],
        email: f[COL.EMAIL], categoria: f[COL.CAT], barrio: f[COL.BARRIO],
        fechaNac: f[COL.FNAC] instanceof Date ? f[COL.FNAC].toISOString().split('T')[0] : f[COL.FNAC]
      }));
  } catch (e) { return []; }
}

function obtenerDatosMesa(dni) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetInsc = getSheet(SHEETS.INSCRIPCIONES);
  const data = sheetInsc.getDataRange().getValues();
  const d = dni.toString().replace(/\D/g, "");
  
  let alumno = null;

  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][COL.DNI].toString().replace(/\D/g, "") === d) {
      // Mapeo de datos basado en tus constantes COL
      alumno = {
        nombre: data[i][COL.NOM], 
        apellido: data[i][COL.APE], 
        dni: data[i][COL.DNI],
        email: data[i][COL.EMAIL], 
        categoria: data[i][COL.CAT], 
        institucion: data[i][COL.INST],
        asistencia: parseFloat(data[i][COL.ASIST]) || 0, // Columna 13 (N)
        
        // IMPORTANTE: Aquí definimos qué leer para validar el acceso
        // Usamos COL.ESTADO porque ahí es donde escribirás "HABILITADO"
        estadoExamen: data[i][COL.ESTADO], // Columna 15 (P)
        
        notaRegistrada: data[i][COL.NOTA], // Columna 14 (O) por si ya rindió
        
        fechaExamen: data[i][COL.F_EX] instanceof Date 
          ? Utilities.formatDate(data[i][COL.F_EX], "GMT-3", "yyyy-MM-dd") 
          : ""
      };
      break;
    }
  }
  
  if (!alumno) return { success: false, message: "DNI no encontrado." };
  
  return { success: true, data: alumno };
}

function obtenerDatosEdicionCompleta(dni) {
  try {
    // 1. Reutilizamos la lógica de búsqueda de alumno
    const resMesa = obtenerDatosMesa(dni);
    if (!resMesa.success) return resMesa;

    // 2. Obtenemos la lista de sedes de la hoja BARRIOS (o SEDES)
    // Asumiendo que la Columna B (índice 1) tiene el nombre de la Institución
    const dataSedes = getSheet(SHEETS.BARRIOS).getDataRange().getValues();
    const sedesUnicas = [...new Set(dataSedes.slice(1)
      .map(fila => fila[1]) // Columna B
      .filter(nombre => nombre && nombre !== ""))];

    return { 
      success: true, 
      data: resMesa.data, 
      sedes: sedesUnicas.map(s => ({ nombre: s })) 
    };

  } catch (e) {
    return { success: false, message: "Error en servidor: " + e.toString() };
  }
}

function registrarAsistenciaFila(dniAlumno, presente, dniOperador) {
  const sheet = getSheet(SHEETS.INSCRIPCIONES);
  const data = sheet.getDataRange().getValues();
  const dAlu = dniAlumno.toString().replace(/\D/g, "");

  for (let i = 1; i < data.length; i++) {
    if (data[i][COL.DNI].toString().replace(/\D/g, "") === dAlu) {
      let actual = parseFloat(data[i][COL.ASIST]) || 0;
      const incremento = 100 / (typeof TOTAL_CLASES !== 'undefined' ? TOTAL_CLASES : 2);
      let nuevo = presente ? Math.min(100, actual + incremento) : Math.max(0, actual - incremento);
      
      sheet.getRange(i + 1, COL.ASIST + 1).setValue(nuevo);
      registrarAccion(dniOperador, presente ? "PRESENTE" : "QUITÓ ASISTENCIA", `DNI Alumno: ${dAlu}`);
      if (presente) getSheet(SHEETS.ASISTENCIA).appendRow([new Date(), dAlu, dniOperador]);
      return { success: true, nuevoValor: nuevo };
    }
  }
  return { success: false, message: "Alumno no encontrado" };
}

function obtenerAlumnosPorFiltro(sede, fechaSeleccionada) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Obtener la fecha de examen configurada para esa sede (Hoja BARRIOS)
    const sheetBarrios = ss.getSheetByName(SHEETS.BARRIOS);
    const dataBarrios = sheetBarrios.getDataRange().getValues();
    // Col B (1) es Institución, Col G (6) es Fecha Examen
    const infoSede = dataBarrios.find(fila => fila[1] === sede);
    
    // Formateamos la fecha del Excel (Date) a string YYYY-MM-DD para comparar en el HTML
    let fechaExamenSede = "";
    if (infoSede && infoSede[6]) {
      const d = new Date(infoSede[6]);
      fechaExamenSede = d.toISOString().split('T')[0];
    }

    // 2. Obtener los alumnos (Hoja INSCRIPCIONES)
    const sheetInsc = ss.getSheetByName(SHEETS.INSCRIPCIONES);
    const dataInsc = sheetInsc.getDataRange().getValues();
    dataInsc.shift(); // Quitar cabecera

    const alumnosFiltrados = dataInsc
      .filter(fila => fila[9] === sede) // Col J (9): Institución/Sede
      .map(fila => {
        return {
          dni: fila[3],           // Col D
          nombre: fila[1],        // Col B
          apellido: fila[2],      // Col C
          asistencia: fila[13],   // Col N (Porcentaje de asistencia)
          nota: fila[14],         // Col O (Estado/Nota de examen)
          finalizado: fila[15] === "FINALIZADO" // Col P
        };
      });

    // 3. Devolvemos el objeto tal cual lo espera el SuccessHandler del HTML
    return {
      alumnos: alumnosFiltrados,
      fechaExamenSede: fechaExamenSede
    };

  } catch (e) {
    console.error("Error en obtenerAlumnosPorFiltro: " + e.toString());
    return { alumnos: [], fechaExamenSede: "", error: e.toString() };
  }
}
// --- 3. PROCESO DE INSCRIPCIÓN ---

function obtenerOpcionesCursada() {
  const sedes = getSheet(SHEETS.BARRIOS).getDataRange().getValues();
  const inscripciones = getSheet(SHEETS.INSCRIPCIONES).getDataRange().getValues();
  const conteo = inscripciones.slice(1).reduce((acc, f) => {
    acc[f[COL.INST]] = (acc[f[COL.INST]] || 0) + 1;
    return acc;
  }, {});

  return sedes.slice(1).filter(r => r[1]).map(r => {
    const inscritos = conteo[r[1]] || 0;
    const cupoMax = parseInt(r[3]) || 0;
    const agotado = (cupoMax > 0 && inscritos >= cupoMax);
    return {
      barrio: r[0], institucion: r[1],
      texto: `${r[0]} - ${r[1]} (${inscritos}/${cupoMax || '∞'})${agotado ? ' [AGOTADO]' : ''}`,
      deshabilitado: agotado
    };
  });
}

function procesarNuevaInscripcion(datos) {
  try {
    const sheet = getSheet(SHEETS.INSCRIPCIONES);
    const data = sheet.getDataRange().getValues();
    const dniLimpio = datos.dni.toString().replace(/\D/g, "");

    // 1. VALIDACIÓN DE DUPLICADOS
    const yaExiste = data.some(fila => fila[COL.DNI].toString().replace(/\D/g, "") === dniLimpio);
    if (yaExiste) {
      return { success: false, message: "Ya existe una inscripción activa para el DNI " + dniLimpio };
    }

    // 2. BUSCAR INFO DE SEDE
    const sedes = getSheet(SHEETS.BARRIOS).getDataRange().getValues();
    let infoSede = sedes.find(s => s[1] === datos.inst) || [];
    const barrioEncontrado = infoSede[0] || "No especificado";

    // 3. ARMAR FILA
    const fila = [
      sheet.getLastRow() + 1,
      datos.nombre,
      datos.apellido,
      dniLimpio,
      datos.fechaNac,
      datos.tel,
      datos.email,
      datos.cat,
      barrioEncontrado,
      datos.inst,
      infoSede[4] || "",
      infoSede[5] || "",
      infoSede[6] || "",
      0,
      "",
      "INSCRIPTO"
    ];

    sheet.appendRow(fila);

    // 4. REGISTRAR EN HISTORIAL
    registrarAccion(
      dniLimpio,
      "NUEVA INSCRIPCIÓN",
      `Alumno: ${datos.apellido}, ${datos.nombre} | Cat: ${datos.cat} | Inst: ${datos.inst}`
    );

    // 5. ENVÍO DE MAIL
    if (datos.email && datos.email.includes("@")) {
      try {
        const fechasObj = {
          fecha1: infoSede[4],
          fecha2: infoSede[5],
          fechaExamen: infoSede[6]
        };

        enviarMailConfirmacion(datos.email, datos, fechasObj);

      } catch (eMail) {
        console.warn("Fallo envío mail: " + eMail.message);
      }
    }

    return { success: true };

  } catch (e) {
    return { success: false, message: "Error en servidor: " + e.toString() };
  }
}

function cancelarInscripcion(dni, dniOperador = "SISTEMA/AUTO") {
  try {
    const sheet = getSheet(SHEETS.INSCRIPCIONES);
    const data = sheet.getDataRange().getValues();
    const dStr = dni.toString().replace(/\D/g, "");

    for (let i = 1; i < data.length; i++) {
      // Usamos COL.DNI para ser fieles a tu estructura
      if (data[i][COL.DNI].toString().replace(/\D/g, "") === dStr) {
        sheet.deleteRow(i + 1);
        registrarAccion(dniOperador, "ELIMINACIÓN/CANCELACIÓN", `DNI: ${dStr}`);
        return "Inscripción cancelada exitosamente.";
      }
    }
    return "No se encontró inscripción activa.";
  } catch (e) { 
    console.error("Error al cancelar: " + e.toString());
    return "Error en el servidor: " + e.toString(); 
  }
}

// --- 4. EXAMEN Y NOTAS ---

function obtenerSedesUnicas() {
  try {
    const sheet = getSheet(SHEETS.BARRIOS);
    const data = sheet.getDataRange().getValues();
    // Quitamos el encabezado y extraemos la columna de Institución (índice 1)
    const sedes = data.slice(1).map(fila => fila[1]);
    
    // Filtramos para que no haya repetidos y quitamos vacíos
    return [...new Set(sedes)].filter(s => s);
  } catch (e) {
    console.error("Error en obtenerSedesUnicas: " + e.toString());
    return [];
  }
}

function obtenerDetalleExamen(dni) {

  try {

    const sheet = getSheet(SHEETS.RESPUESTAS);
    const data = sheet.getDataRange().getValues();
    const dStr = dni.toString().replace(/\D/g,"");

    for (let i = data.length - 1; i >= 1; i--) {

      if (data[i][COL_RESP.DNI].toString().replace(/\D/g,"") === dStr) {

        const respuestas = typeof data[i][COL_RESP.RESPUESTAS] === "string"
          ? JSON.parse(data[i][COL_RESP.RESPUESTAS])
          : data[i][COL_RESP.RESPUESTAS];

        return {
          success:true,
          fecha:Utilities.formatDate(data[i][COL_RESP.FECHA],"GMT-3","dd/MM/yyyy HH:mm"),
          nota:data[i][COL_RESP.NOTA],
          estado:data[i][COL_RESP.ESTADO],
          respuestas:respuestas
        };

      }
    }

    return {success:false};

  } catch(e) {

    return {
      success:false,
      error:e.toString()
    };

  }

}

function puedeRendirExamen(dni, maxIntentos = 1) {
  try {

    const sheet = getSheet(SHEETS.RESPUESTAS);
    const data = sheet.getDataRange().getValues();
    const dStr = dni.toString().replace(/\D/g, "");

    let intentos = 0;

    for (let i = 1; i < data.length; i++) {

      const dniFila = data[i][COL_RESP.DNI].toString().replace(/\D/g, "");
      const estado = data[i][COL_RESP.ESTADO];

      if (dniFila === dStr) {

        // 🔥 SOLO CUENTA INTENTOS NO APROBADOS (PRO)
        if (estado !== "APROBADO") {
          intentos++;
        }

      }
    }

    return {
      permitido: intentos < maxIntentos,
      intentos: intentos
    };

  } catch (e) {

    console.error("Error en puedeRendirExamen:", e);

    return {
      permitido: false,
      intentos: 0,
      error: e.toString()
    };
  }
}

function sincronizarExamenConInscriptos() {

  const ss = SpreadsheetApp.getActive();
  const sheetResp = ss.getSheetByName('RespuestasExamen');
  const sheetIns = ss.getSheetByName('Inscriptos');
  const dataResp = sheetResp.getDataRange().getValues();
  const dataIns = sheetIns.getDataRange().getValues();

  const mapaIns = {};

  // 📌 indexar inscriptos por DNI
  for (let i = 1; i < dataIns.length; i++) {
    const dni = String(dataIns[i][COL_INS.DNI]).replace(/\D/g,"");
    mapaIns[dni] = i;
  }

  // 🔄 recorrer respuestas
  for (let i = 1; i < dataResp.length; i++) {

    const dni = String(dataResp[i][COL_RESP.DNI]).replace(/\D/g,"");
    const nota = dataResp[i][COL_RESP.NOTA];
    const estado = dataResp[i][COL_RESP.ESTADO];

    const filaIns = mapaIns[dni];

    if (filaIns !== undefined) {

      // ✅ actualizar nota
      sheetIns.getRange(filaIns + 1, COL_INS.NOTA + 1).setValue(nota);

      // ✅ actualizar estado trámite
      let nuevoEstado = 'PENDIENTE';

      if (estado === 'APROBADO') {
        nuevoEstado = 'APROBADO';
      } else if (estado === 'DESAPROBADO') {
        nuevoEstado = 'DESAPROBADO';
      }

      sheetIns.getRange(filaIns + 1, COL_INS.ESTADO_TRAMITE + 1).setValue(nuevoEstado);

    }
  }

}

function marcarAsistencia(dniAlumno, esSuma, dniOperador) {
  try {
    const sheet = getSheet(SHEETS.INSCRIPCIONES);
    const data = sheet.getDataRange().getValues();
    const dAlu = dniAlumno.toString().replace(/\D/g, "");

    for (let i = 1; i < data.length; i++) {
      if (data[i][COL.DNI].toString().replace(/\D/g, "") === dAlu) {
        let actual = parseFloat(data[i][COL.ASIST]) || 0;
        
        // Lógica de incremento (50% por clase para 2 clases totales)
        const incremento = 50; 
        let nuevo = esSuma ? Math.min(100, actual + incremento) : Math.max(0, actual - incremento);
        
        sheet.getRange(i + 1, COL.ASIST + 1).setValue(nuevo);
        
        // Registro de auditoría
        registrarAccion(dniOperador, esSuma ? "PRESENTE" : "QUITÓ ASISTENCIA", `DNI Alumno: ${dAlu}`);
        
        // Registro en historial específico
        if (esSuma) {
          getSheet(SHEETS.ASISTENCIA).appendRow([new Date(), dAlu, dniOperador]);
        }
        
        return { success: true, nuevoValor: nuevo };
      }
    }
    return { success: false, message: "Alumno no encontrado" };
  } catch (e) {
    return { success: false, message: "Error Servidor: " + e.toString() };
  }
}

function guardarPreguntaServidor(datos, dniOp) {
  const sheet = getSheet(SHEETS.PREGUNTAS);
  const fila = [datos.pregunta, datos.opciones[0], datos.opciones[1], datos.opciones[2], datos.correcta, datos.tiempo, 10, datos.excluyente];
  
  if (datos.id) {
    sheet.getRange(datos.id, 1, 1, 8).setValues([fila]);
    registrarAccion(dniOp, "EDITÓ PREGUNTA ID: " + datos.id);
  } else {
    sheet.appendRow(fila);
    registrarAccion(dniOp, "CREÓ PREGUNTA");
  }
  return { success: true };
}

function eliminarPreguntaServidor(indice, dniOp) {
  try {
    getSheet(SHEETS.PREGUNTAS).deleteRow(indice + 2);
    registrarAccion(dniOp, "ELIMINAR PREGUNTA", `Fila: ${indice + 2}`);
    return { success: true };
  } catch (e) { return { success: false, message: e.toString() }; }
}

// --- 6. PERSONAL Y LOGIN ---

function crearNuevoPersonal(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.PERSONAL);
    const data = sheet.getDataRange().getValues();
    
    // 1. Limpiar DNI (quitar puntos o espacios)
    const dniNuevo = datos.dni.toString().replace(/\D/g, "");

    // 2. Validar duplicados en la columna C (índice 2)
    const existe = data.some(fila => fila[2].toString().replace(/\D/g, "") === dniNuevo);
    
    if (existe) {
      return { success: false, message: "El DNI " + dniNuevo + " ya está registrado en el sistema." };
    }

    // 3. Preparar la fila con la clave por defecto "12345"
    // Col A: Nombre | B: Apellido | C: DNI | D: Perfil | E: Email | F: Tel | G: Pass | H: RequiereCambio
    const nuevaFila = [
      datos.nombre, 
      datos.apellido, 
      dniNuevo, 
      datos.perfil, 
      datos.email, 
      datos.telefono, 
      "12345", // Password inicial
      "SI"    // Forzamos el cambio de clave en el primer ingreso
    ];

    sheet.appendRow(nuevaFila);
    registrarAccion(dniNuevo, "ADMIN CREÓ USUARIO", datos.perfil);
    
    return { success: true, message: "Usuario creado exitosamente con clave '1234'" };
    
  } catch (e) {
    return { success: false, message: "Error: " + e.toString() };
  }
}

function loginPersonal(dni, password) {
  const data = getSheet(SHEETS.PERSONAL).getDataRange().getValues();
  const d = dni.toString().replace(/\D/g, "");
  
  for (let i = 1; i < data.length; i++) {
    // Si coinciden DNI y Password
    if (data[i][2].toString().replace(/\D/g, "") === d && data[i][6].toString() === password) {
      return { 
        success: true, 
        perfil: data[i][3], 
        requiereCambio: (data[i][7] === "SI"), // <--- ESTO ES CLAVE (Columna H)
        nombre: data[i][0]
      };
    }
  }
  return { success: false, message: "DNI o Contraseña incorrectos" };
}

function actualizarPasswordPersonal(dni, nuevaPass) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.PERSONAL);
  const data = sheet.getDataRange().getValues();
  const d = dni.toString().replace(/\D/g, "");

  // VALIDACIÓN DE SEGURIDAD LADO SERVIDOR
  if (nuevaPass === "12345" || nuevaPass.length < 4) {
    return { success: false, message: "La clave es muy débil o es la de defecto." };
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][2].toString().replace(/\D/g, "") === d) {
      // Escribimos en la fila i+1
      sheet.getRange(i + 1, 7).setValue(nuevaPass); // Columna G (Password)
      sheet.getRange(i + 1, 8).setValue("NO");      // Columna H (RequiereCambio)
      
      registrarAccion(d, "ACTUALIZÓ SU PASSWORD", "SISTEMA");
      return { success: true };
    }
  }
  return { success: false, message: "Usuario no encontrado." };
}

function resetearPasswordPorAdmin(dniNuevo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.PERSONAL);
  const data = sheet.getDataRange().getValues();
  const d = dniNuevo.toString().replace(/\D/g, "");
  const CLAVE_DEFECTO = "12345"; // <--- Esta es tu clave por defecto

  for (let i = 1; i < data.length; i++) {
    if (data[i][2].toString().replace(/\D/g, "") === d) {
      sheet.getRange(i + 1, 7).setValue(CLAVE_DEFECTO); 
      sheet.getRange(i + 1, 8).setValue("SI"); // Marcamos que debe cambiarla
      return "Clave reseteada a '" + CLAVE_DEFECTO + "' exitosamente.";
    }
  }
  return "Error: No se encontró el DNI solicitado.";
}

function buscarPersonaPorDni(dni) {
  const data = getSheet(SHEETS.PERSONAL).getDataRange().getValues();
  const d = dni.toString().replace(/\D/g, "");
  for (let i = 1; i < data.length; i++) {
    if (data[i][2].toString().replace(/\D/g, "") === d) {
      return { nombre: data[i][0], apellido: data[i][1], cargo: data[i][3] };
    }
  }
  return null;
}

function buscarAlumnoParaEdicion(query) {
  try {
    const sheet = getSheet(SHEETS.INSCRIPCIONES);
    const data = sheet.getDataRange().getValues();
    const q = query.toString().toLowerCase().trim().replace(/\D/g, ""); // Versión numérica para DNI
    const qTexto = query.toString().toLowerCase().trim(); // Versión texto para Apellido

    const encontrados = data.slice(1) // Quitamos cabecera
      .filter(f => {
        const dni = f[COL.DNI] ? f[COL.DNI].toString().replace(/\D/g, "") : "";
        const ape = f[COL.APE] ? f[COL.APE].toString().toLowerCase() : "";
        // Busca coincidencia parcial en DNI o Apellido
        return dni.includes(q) || ape.includes(qTexto);
      })
      .map(f => {
        return {
          nombre: f[COL.NOM],
          apellido: f[COL.APE],
          dni: f[COL.DNI],
          institucion: f[COL.INST] || "Sin Sede Asignada",
          asistencia: f[COL.ASIST] || 0,
          estado: f[COL.ESTADO]
        };
      });

    if (encontrados.length === 0) {
      return { success: false, message: "No se encontraron alumnos con el criterio: " + query };
    }

    return { 
      success: true, 
      alumnos: encontrados.slice(0, 15) // Limitamos a 15 resultados para no saturar el panel
    };

  } catch (e) {
    console.error("Error en buscarAlumnoParaEdicion: " + e.toString());
    return { success: false, message: "Error en la búsqueda: " + e.toString() };
  }
}
// --- 7. AUXILIARES ---
function marcarTramiteFinalizado(dni, dniOperador) {

  try {

    const sheet = getSheet(SHEETS.INSCRIPCIONES);
    const data = sheet.getDataRange().getValues();
    const dAlu = dni.toString().replace(/\D/g, "");

    for (let i = 1; i < data.length; i++) {

      if (data[i][COL_INS.DNI].toString().replace(/\D/g, "") === dAlu) {

        // NO tocamos la nota (columna O)

        // Estado del trámite → columna P
        sheet.getRange(i + 1, COL_INS.ESTADO_TRAMITE + 1)
          .setValue("FINALIZADO");

        // Operador que entregó certificado → columna Q
        if (dniOperador) {
          sheet.getRange(i + 1, COL_INS.OPERADOR + 1)
            .setValue(dniOperador);
        }

        return true;
      }

    }

    return false;

  } catch (e) {

    console.error("Error en marcarTramiteFinalizado: " + e.toString());
    return false;

  }

}

