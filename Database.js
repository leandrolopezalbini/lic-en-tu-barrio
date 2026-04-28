// Database.js

// ================================
// 🔹 CORE
// ================================

function limpiarDni(dni) {
  return (dni || "").toString().replace(/\D/g, "");
}

function logErrorGS(contexto, e) {
  Logger.log(`❌ ${contexto}: ${e}`);
}

// ================================
// 🔹 HELPERS INSCRIPCION
// ================================

function existeDni(dni) {
  const data = getSheet(SHEETS.INSCRIPCIONES).getDataRange().getValues();
  const d = limpiarDni(dni);

  return data.some(f =>
    f[COL_INS.DNI] &&
    limpiarDni(f[COL_INS.DNI]) === d
  );
}

function obtenerInfoSede(nombreSede) {
  const sedes = getSheet(SHEETS.BARRIOS).getDataRange().getValues();
  return sedes.find(s => s[COL_BAR.INST] === nombreSede) || null;
}

function armarFilaInscripcion(datos, infoSede) {
  return [
    getSheet(SHEETS.INSCRIPCIONES).getLastRow() + 1,
    datos.nombre,
    datos.apellido,
    limpiarDni(datos.dni),
    datos.fechaNac,
    datos.tel,
    datos.email,
    datos.cat,
    infoSede?.[COL_BAR.BARRIO] || "No especificado",
    datos.inst,
    infoSede?.[COL_BAR.F1] || "",
    infoSede?.[COL_BAR.F2] || "",
    infoSede?.[COL_BAR.F_EX] || "",
    0,
    "",
    "INSCRIPTO"
  ];
}

function insertarInscripcion(fila) {
  getSheet(SHEETS.INSCRIPCIONES).appendRow(fila);
}




///////////////////////////////////
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
      fecha: Utilities.formatDate(new Date(f[COL_LOGS.FECHA]), "GMT-3", "dd/MM/yyyy HH:mm"),
      dni: f[COL_LOGS.DNI], operador: f[COL_LOGS.OPERADOR], accion: f[COL_LOGS.ACCION], detalle: f[COL_LOGS.DETALLE]
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
        const dni = f[COL_INS.DNI] ? f[COL_INS.DNI].toString() : "";
        const ape = f[COL_INS.APELLIDO] ? f[COL_INS.APELLIDO].toString().toLowerCase() : "";
        return dni.includes(q) || ape.includes(q);
      })
      .slice(0, 10)
      .map(f => ({
        nombre: f[COL_INS.NOMBRE], apellido: f[COL_INS.APELLIDO], dni: f[COL_INS.DNI],
        institucion: f[COL_BAR.INST] || "Sin Sede",
        asistencia: f[COL_INS.ASISTENCIA], nota: f[COL_INS.NOTA], estado: f[COL_INS.ESTADO_TRAMITE],
        email: f[COL_INS.EMAIL], categoria: f[COL_INS.CATEGORIA], barrio: f[COL_INS.BARRIO],
        fechaNac: f[COL_INS.FNAC] instanceof Date ? f[COL_INS.FNAC].toISOString().split('T')[COL_INS.FECHA_NAC] : f[COL_INS.FNAC]
      }));
  } catch (e) { return []; }
}

function obtenerDatosMesa(dni) {
  const sheetInsc = getSheet(SHEETS.INSCRIPCIONES);
  const data = sheetInsc.getDataRange().getValues();
  const d = dni.toString().replace(/\D/g, "");
  
  let alumno = null;

  for (let i = data.length - 1; i >= 1; i--) {

    const fila = data[i];

    if (!fila || !fila[COL_INS.DNI]) continue;

    if (fila[COL_INS.DNI].toString().replace(/\D/g, "") === d) {
      alumno = {
        nombre: fila[COL_INS.NOMBRE],
        apellido: fila[COL_INS.APELLIDO],
        dni: fila[COL_INS.DNI],
        email: fila[COL_INS.EMAIL],
        categoria: fila[COL_INS.CATEGORIA],
        institucion: fila[COL_INS.INSTITUCION],
        asistencia: parseFloat(fila[COL_INS.ASISTENCIA]) || 0,
        estadoExamen: fila[COL_INS.ESTADO_TRAMITE],
        notaRegistrada: fila[COL_INS.NOTA],
        fechaExamen: fila[COL_INS.FECHA_EXAMEN] instanceof Date 
          ? Utilities.formatDate(fila[COL_INS.FECHA_EXAMEN], "GMT-3", "yyyy-MM-dd") 
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
    const dataSedes = getSheet(SHEETS.BARRIOS).getDataRange().getValues();
    const sedesUnicas = [...new Set(dataSedes.slice(1)
      .map(fila => fila[COL_BAR.INST])
      
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
  try {

    const sheet = getSheet(SHEETS.INSCRIPCIONES);
    const data = sheet.getDataRange().getValues();
    const dAlu = dniAlumno.toString().replace(/\D/g, "");

    for (let i = 1; i < data.length; i++) {

      if (!data[i][COL_INS.DNI]) continue;

      if (data[i][COL_INS.DNI].toString().replace(/\D/g, "") === dAlu) {

        let actual = parseFloat(data[i][COL_INS.ASISTENCIA]) || 0;

        const incremento = 100 / (typeof TOTAL_CLASES !== 'undefined' ? TOTAL_CLASES : 2);

        let nuevo = presente
          ? Math.min(100, actual + incremento)
          : Math.max(0, actual - incremento);

        sheet.getRange(i + 1, COL_INS.ASISTENCIA + 1).setValue(nuevo);

        registrarAccion(
          dniOperador,
          presente ? "PRESENTE" : "QUITÓ ASISTENCIA",
          `DNI Alumno: ${dAlu}`
        );

        if (presente) {
          getSheet(SHEETS.ASISTENCIA).appendRow([new Date(), dAlu, dniOperador]);
        }

        return { success: true, nuevoValor: nuevo };
      }
    }

    return { success: false, message: "Alumno no encontrado" };

  } catch (error) {

    logError("registrarAsistenciaFila", error, {
      dniAlumno,
      presente,
      dniOperador
    });

    return {
      success: false,
      message: "Error interno del servidor"
    };
  }
}

function obtenerAlumnosPorFiltro(sede) {

  const sheet = getSheet(SHEETS.INSCRIPCIONES);
  const data = sheet.getDataRange().getValues();

  const sedeFiltro = (sede || "").toString().trim().toLowerCase();

  // 📅 HOY
  const hoy = normalizarFecha(new Date());

  const alumnos = data.slice(1)
    .filter(r => {

      const sedeFila = (r[COL_INS.INSTITUCION] || "")
        .toString()
        .trim()
        .toLowerCase();

      const f1 = normalizarFecha(r[COL_INS.CURSADA1]);
      const f2 = normalizarFecha(r[COL_INS.CURSADA2]);
      const fEx = normalizarFecha(r[COL_INS.FECHA_EXAMEN]);

      // 🎯 coincide si HOY es cualquiera de las 3 fechas
      const coincideFecha = (hoy === f1 || hoy === f2 || hoy === fEx);

      return sedeFila === sedeFiltro && coincideFecha;

    })
    .map(r => ({

      nombre: r[COL_INS.NOMBRE],
      apellido: r[COL_INS.APELLIDO],
      dni: r[COL_INS.DNI],
      asistencia: r[COL_INS.ASISTENCIA],
      nota: r[COL_INS.NOTA],

      finalizado: r[COL_INS.ESTADO_TRAMITE] === "FINALIZADO",

      // 👇 info útil para UI
      esExamen: hoy === normalizarFecha(r[COL_INS.FECHA_EXAMEN])

    }));

  return {
    alumnos: alumnos,
    fechaHoy: hoy
  };
}

// --- 3. PROCESO DE INSCRIPCIÓN ---

function obtenerOpcionesCursada() {
  try {
    const sheetBarrios = getSheet(SHEETS.BARRIOS);
    const sheetInsc = getSheet(SHEETS.INSCRIPCIONES);

    if (!sheetBarrios || !sheetInsc) {
      throw new Error("No se encontraron las hojas");
    }

    const sedes = sheetBarrios.getDataRange().getValues();
    const inscripciones = sheetInsc.getDataRange().getValues();

    if (sedes.length <= 1) return [];

    // ===============================
    // 🔹 CONTEO DE INSCRIPTOS
    // ===============================
    const conteo = {};

    inscripciones.slice(1).forEach(f => {
      const sede = (f[COL_INS.INSTITUCION] || "").toString().trim();
      if (!sede) return;

      conteo[sede] = (conteo[sede] || 0) + 1;
    });

    // ===============================
    // 🔹 ARMADO DE OPCIONES
    // ===============================
    const resultado = sedes.slice(1)
      .filter(r => r[COL_BAR.INST])
      .map(r => {

        const barrio = (r[COL_BAR.BARRIO] || "").toString().trim();
        const sedeNombre = (r[COL_BAR.INST] || "").toString().trim();

        if (!sedeNombre) return null;

        const inscritos = conteo[sedeNombre] || 0;

        const cupoRaw = r[COL_BAR.CUPO];
        const cupoMax = Number(cupoRaw) || 0;

        const agotado = (cupoMax > 0 && inscritos >= cupoMax);

        return {
          barrio,
          institucion: sedeNombre,
          texto: `${barrio} - ${sedeNombre} (${inscritos}/${cupoMax || '∞'})${agotado ? ' [AGOTADO]' : ''}`,
          deshabilitado: agotado
        };
      })
      .filter(x => x !== null);
      // 👇 👉 ACÁ EXACTAMENTE
    Logger.log("RESULTADO OPCIONES:");
    Logger.log(JSON.stringify(resultado, null, 2));
    Logger.log("SEDES RAW:");
    Logger.log(JSON.stringify(sedes.slice(0,5)));
        
    Logger.log("INSCRIPCIONES RAW:");
    Logger.log(JSON.stringify(inscripciones.slice(0,5)));
    return resultado;    

  } catch (err) {
    Logger.log("ERROR obtenerOpcionesCursada: " + err.message);
    
    return [];
  }
}

function procesarNuevaInscripcion(datos) {
  try {

    const dni = limpiarDni(datos.dni);

    // 1. Validar duplicado
    if (existeDni(dni)) {
      return {
        success: false,
        message: "Ya existe una inscripción para ese DNI " + dni
      };
    }

    // 2. Obtener sede
    const infoSede = obtenerInfoSede(datos.inst);

    // 3. Armar fila
    const fila = armarFilaInscripcion(datos, infoSede);

    // 4. Insertar
    insertarInscripcion(fila);

    // 5. Log
    registrarAccion(
      dni,
      "NUEVA INSCRIPCIÓN",
      `Alumno: ${datos.apellido}, ${datos.nombre} | Cat: ${datos.cat} | Inst: ${datos.inst}`
    );

    // 6. Mail (igual que antes)
    if (datos.email && datos.email.includes("@")) {
      try {
        const fechasObj = {
          fecha1: infoSede?.[COL_BAR.F1] || "",
          fecha2: infoSede?.[COL_BAR.F2] || "",
          fechaExamen: infoSede?.[COL_BAR.F_EX] || ""
        };

        enviarMailConfirmacion(datos.email, datos, fechasObj);

      } catch (eMail) {
        Logger.log("⚠️ Error mail: " + eMail.message);
      }
    }

    return { success: true };

  } catch (e) {
    logErrorGS("procesarNuevaInscripcion", e);
    return {
      success: false,
      message: "Error en servidor: " + e.toString()
    };
  }
}

function cancelarInscripcion(dni, dniOperador = "SISTEMA/AUTO") {
  try {
    const sheet = getSheet(SHEETS.INSCRIPCIONES);
    const data = sheet.getDataRange().getValues();
    const dStr = dni.toString().replace(/\D/g, "");

    for (let i = 1; i < data.length; i++) {
      // Usamos COL_INS.DNI para ser fieles a tu estructura
      if (data[i][COL_INS.DNI].toString().replace(/\D/g, "") === dStr) {
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
    const sedes = data.slice(1).map(fila => fila[COL_BAR.BARRIO]);
    
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
      if (data[i][COL_INS.DNI].toString().replace(/\D/g, "") === dAlu) {
        let actual = parseFloat(data[i][COL_INS.ASISTENCIA]) || 0;
        
        // Lógica de incremento (50% por clase para 2 clases totales)
        const incremento = 50; 
        let nuevo = esSuma ? Math.min(100, actual + incremento) : Math.max(0, actual - incremento);
        
        sheet.getRange(i + 1, COL_INS.ASISTENCIA + 1).setValue(nuevo);
        
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
  const fila = [datos.pregunta, datos.opciones[COL_PREG.OPC1], datos.opciones[COL_PREG.OPC2], datos.opciones[COL_PREG.OPC3], datos.correcta, datos.tiempo, 10, datos.excluyente];
  
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
    limpiarDni(datos.dni);

    // 2. Validar duplicados en la columna C (índice 2)
    const existe = data.some(fila => fila[COL_PER.DNI].toString().replace(/\D/g, "") === dniNuevo);
    
    if (existe) {
      return { success: false, message: "El DNI " + dniNuevo + " ya está registrado en el sistema." };
    }

    // 3. Preparar la fila con la clave por defecto "12345"
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
  try {
    const data = getSheet(SHEETS.PERSONAL).getDataRange().getValues();
    const d = dni.toString().replace(/\D/g, "");
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // 🛑 evitar undefined
      if (!row[COL_PER.DNI] || !row[COL_PER.PASSWORD]) continue;

      const dniFila = row[COL_PER.DNI].toString().replace(/\D/g, "");
      const passFila = row[COL_PER.PASSWORD].toString();

      if (dniFila === d && passFila === password) {
        Logger.log("LOGIN OK → PERFIL: " + row[COL_PER.PERFIL]);

        return { 
          success: true, 
          perfil: row[COL_PER.PERFIL], 
          requiereCambio: (row[COL_PER.REQUIERE_CAMBIO] === "SI"),
          nombre: row[COL_PER.NOMBRE],
          apellido: row[COL_PER.APELLIDO]
        };
      }
    }

    return { success: false, message: "DNI o Contraseña incorrectos" };

  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function buscarPersonaPorDni(dni) {
  const data = getSheet(SHEETS.PERSONAL).getDataRange().getValues();
  const d = dni.toString().replace(/\D/g, "");
  for (let i = 1; i < data.length; i++) {
    if (data[i][COL_PER.DNI].toString().replace(/\D/g, "") === d) {
      return { nombre: data[i][COL_PER.NOMBRE], apellido: data[i][COL_PER.APELLIDO], cargo: data[i][COL_PER.PERFIL] };
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
        const dni = f[COL_INS.DNI] ? f[COL_INS.DNI].toString().replace(/\D/g, "") : "";
        const ape = f[COL_INS.APELLIDO] ? f[COL_INS.APELLIDO].toString().toLowerCase() : "";
        // Busca coincidencia parcial en DNI o Apellido
        return dni.includes(q) || ape.includes(qTexto);
      })
      .map(f => {
        return {
          nombre: f[COL_INS.NOMBRE],
          apellido: f[COL_INS.APELLIDO],
          dni: f[COL_INS.DNI],
          institucion: f[COL_INS.INSTITUCION] || "Sin Sede Asignada",
          asistencia: f[COL_INS.ASISTENCIA] || 0,
          estado: f[COL_INS.ESTADO_TRAMITE]
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

function resetearPasswordPorAdmin(dniNuevo) {
  const sheet = getSheet(SHEETS.PERSONAL);
  const data = sheet.getDataRange().getValues();
  const d = dniNuevo.toString().replace(/\D/g, "");
  const CLAVE_DEFECTO = "12345";

  for (let i = 1; i < data.length; i++) {
    if (data[i][COL_PER.DNI].toString().replace(/\D/g, "") === d) {
      sheet.getRange(i + 1, COL_PER.PASSWORD + 1).setValue(CLAVE_DEFECTO);
      sheet.getRange(i + 1, COL_PER.REQUIERE_CAMBIO + 1).setValue("SI");
      return { success: true };
    }
  }

  return { success: false, message: "Usuario no encontrado." };
}

function actualizarPasswordPersonal(dni, nuevaPass) {
  const sheet = getSheet(SHEETS.PERSONAL);
  const data = sheet.getDataRange().getValues();
  const d = dni.toString().replace(/\D/g, "");

  if (nuevaPass === "12345" || nuevaPass.length < 4) {
    return { success: false, message: "La clave es muy débil o es la de defecto." };
  }

  for (let i = 1; i < data.length; i++) {
    const dniFila = data[i][COL_PER.DNI].toString().replace(/\D/g, "");

    if (dniFila === d) {
      sheet.getRange(i + 1, COL_PER.PASSWORD + 1).setValue(nuevaPass);
      sheet.getRange(i + 1, COL_PER.REQUIERE_CAMBIO + 1).setValue("NO");

      registrarAccion(d, "ACTUALIZÓ SU PASSWORD", "SISTEMA");
      return { success: true };
    }
  }

  return { success: false, message: "Usuario no encontrado." };
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

function limpiarImagenesPreguntas() {
  const sheet = getSheet(SHEETS.PREGUNTAS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {

    let texto = (data[i][0] || "").toString();
    let imgCol = (data[i][5] || "").toString();

    // 🔥 extraer ID desde texto viejo
    const match = texto.match(/obtenerUrlImagen\('(.+?)'\)/);

    if (match) {
      const id = match[1];

      // limpiar texto
      texto = texto.replace(match[0], "").trim();

      sheet.getRange(i + 1, 1).setValue(texto); // columna A
      sheet.getRange(i + 1, 6).setValue(id);    // columna F

      Logger.log(`Fila ${i+1} migrada`);
    }

    // 🔥 limpiar URLs completas
    if (imgCol.includes("drive.google.com")) {
      const idMatch = imgCol.match(/[-\w]{25,}/);
      if (idMatch) {
        sheet.getRange(i + 1, 6).setValue(idMatch[0]);
        Logger.log(`ID limpiado fila ${i+1}`);
      }
    }
  }

  Logger.log("✔ Limpieza terminada");
}