// ExamenLogic.gs
function examenEstaHabilitado() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const conf = ss.getSheetByName("Configuracion") || ss.getSheetByName("Config");
    
    // Asumiendo que en B6 tienes la fecha/hora de inicio y en B7 la de fin
    const inicio = conf.getRange("B6").getValue(); 
    const fin = conf.getRange("B7").getValue();
    const ahora = new Date();

    const habilitado = (ahora >= new Date(inicio) && ahora <= new Date(fin));

    return {
      habilitado: habilitado,
      inicio: Utilities.formatDate(new Date(inicio), "GMT-3", "dd/MM HH:mm"),
      fin: Utilities.formatDate(new Date(fin), "GMT-3", "dd/MM HH:mm")
    };
  } catch (e) {
    // Si no hay fechas configuradas, devolvemos deshabilitado por seguridad
    return { habilitado: false, inicio: "No definida", fin: "No definida" };
  }
}

function procesarExamenDesdeCliente(respuestasDelAlumno, dni) {
  try {
    const preguntasDB = obtenerTodasLasPreguntas();
    let puntosLogrados = 0;
    let totalPuntosPosibles = 0;
    let falloExcluyente = false;
    let detalleParaDB = [];

    preguntasDB.forEach((p, index) => {
      // Ahora buscamos la respuesta por el índice (0, 1, 2...) que envía el cliente
      const respuestaDada = respuestasDelAlumno[index] ? respuestasDelAlumno[index].toLowerCase() : "skip";
      const acerto = (respuestaDada === p.correcta);
      
      const puntosPregunta = p.puntos; 
      totalPuntosPosibles += puntosPregunta;

      if (acerto) {
        puntosLogrados += puntosPregunta;
      } else if (p.excluyente) {
        falloExcluyente = true;
      }
      
      detalleParaDB.push({
        pregunta: p.texto.replace(/<[^>]*>?/gm, ''), // Guardamos solo texto en la DB sin HTML
        respuesta: respuestaDada,
        correcta: p.correcta,
        estado: acerto ? "OK" : (p.excluyente && !acerto ? "FALLO EXCLUYENTE" : "ERROR")
      });
    });

    let porcentajeLogrado = totalPuntosPosibles > 0 
      ? Math.round((puntosLogrados / totalPuntosPosibles) * 100) 
      : 0;

    // Nota 0 si falló excluyente, sino el porcentaje
    let notaFinal = falloExcluyente ? 0 : porcentajeLogrado;
    const aprobado = (notaFinal >= 75 && !falloExcluyente);

    const registro = finalizarExamen(dni, notaFinal, JSON.stringify(detalleParaDB), falloExcluyente);

    if (registro.success) {
      return {
        success: true,
        nota: notaFinal,
        aprobado: aprobado,
        excluyente: falloExcluyente
      };
    } else {
      throw new Error(registro.message);
    }
    
  } catch (e) {
    console.error("Error procesarExamen: " + e.toString());
    return { success: false, message: "Error en el servidor: " + e.toString() };
  }
}

function finalizarExamen(dni, nota, respuestasJson, huboFalloExcluyente) {
  try {
    const sheet = getSheet(SHEETS.INSCRIPCIONES);
    const data = sheet.getDataRange().getValues();
    const dStr = dni.toString().replace(/\D/g, "");

    let alumnoEncontrado = false;
    let emailAlumno = "";
    let nombreCompleto = "";
    let filaIndex = -1;

    // 🔍 Buscar alumno
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][COL_INS.DNI].toString().replace(/\D/g, "") === dStr) {
        filaIndex = i + 1;
        nombreCompleto = data[i][COL_INS.NOMBRE] + " " + data[i][COL_INS.APELLIDO];
        emailAlumno = data[i][COL_INS.EMAIL];
        alumnoEncontrado = true;
        break;
      }
    }

    if (!alumnoEncontrado) {
      return { success: false, message: "No se encontró el alumno con DNI: " + dStr };
    }

    // 🔴 NORMALIZAR RESPUESTAS (CLAVE)
    const respuestas = typeof respuestasJson === "string"
      ? JSON.parse(respuestasJson)
      : respuestasJson;

    const respuestasEnriquecidas = respuestas.map(r => ({
      ...r,
      excluyente: r.excluyente === true
    }));

    const respuestasFinal = JSON.stringify(respuestasEnriquecidas);

    // 🔴 IMPORTANTE: NO SOBREESCRIBIR NOTA SI YA EXISTE
    const notaActual = data[filaIndex - 1][COL_INS.NOTA];

    if (!notaActual || notaActual === "") {
      sheet.getRange(filaIndex, COL_INS.NOTA + 1).setValue(nota);
    }

    // ✔ SOLO CAMBIAMOS ESTADO
    sheet.getRange(filaIndex, COL_INS.ESTADO_TRAMITE + 1).setValue("FINALIZADO");

    SpreadsheetApp.flush();

    // 🔴 LÓGICA APROBACIÓN
    const aprobado = (nota >= 75 && !huboFalloExcluyente);
    const estadoTexto = huboFalloExcluyente
      ? "REPROBADO (EXCLUYENTE)"
      : (aprobado ? "APROBADO" : "REPROBADO");

    // 📝 GUARDAR RESPUESTAS
    const sheetResp = getSheet(SHEETS.RESPUESTAS);
    sheetResp.appendRow([
      new Date(),
      dStr,
      nombreCompleto,
      nota,
      estadoTexto,
      respuestasFinal // 🔥 ACA VA EL FIX
    ]);

    // 📧 MAIL
    if (emailAlumno) {
      enviarCorreoResultado(emailAlumno, nombreCompleto, nota, aprobado, huboFalloExcluyente);
    }

    // 📜 LOG
    registrarAccion(
      "SISTEMA",
      "CIERRE EXAMEN",
      `DNI: ${dStr} - Nota: ${nota}% - ${estadoTexto}`
    );

    return {
      success: true,
      aprobado: aprobado,
      nota: nota,
      excluyente: huboFalloExcluyente
    };

  } catch (e) {
    console.error("Error en finalizarExamen: " + e.toString());
    return { success: false, message: e.toString() };
  }
}

function actualizarNotaEnBD(dni, nota, detalle) {
  const sheet = getSheet(SHEETS.INSCRIPCIONES);
  const data = sheet.getDataRange().getValues();
  const dStr = dni.toString().replace(/\D/g, "");

  for (let i = 1; i < data.length; i++) {
    if (data[i][COL_INS.DNI].toString().replace(/\D/g, "") === dStr) {
      // Guardamos la nota numérica
      sheet.getRange(i + 1, COL_INS.NOTA + 1).setValue(nota);
      
      // Opcional: Guardar el JSON del detalle en una columna oculta para auditoría
      // sheet.getRange(i + 1, COL_INS.ESTADO_TRAMITE + 1).setValue("RENDIDO");
      
      registrarAccion("SISTEMA", "EXAMEN FINALIZADO", `DNI: ${dStr} - Nota: ${nota}`);
      break;
    }
  }
}

function habilitarExamen(dniAlumno, dniOperador) {
  try {
    const sheet = getSheet(SHEETS.INSCRIPCIONES);
    const data = sheet.getDataRange().getValues();
    const dAlu = dniAlumno.toString().replace(/\D/g, "");

    for (let i = 1; i < data.length; i++) {
      if (data[i][COL_INS.DNI].toString().replace(/\D/g, "") === dAlu) {
        
        // Escribimos HABILITADO en la Columna P (COL_INS.ESTADO_TRAMITE)
        sheet.getRange(i + 1, COL_INS.ESTADO_TRAMITE + 1).setValue("HABILITADO");
        
        // NO LIMPIAMOS COL_INS.NOTA: Así mantenemos el registro de la nota anterior 
        // hasta que el sistema la sobreescriba al finalizar el nuevo examen.

        registrarAccion(dniOperador, "HABILITÓ EXAMEN", `DNI Alumno: ${dAlu}`);
        return { success: true, message: "Alumno habilitado correctamente." };
      }
    }
    return { success: false, message: "No se encontró el alumno." };
  } catch (e) {
    console.error("Error en habilitarExamen: " + e.toString());
    return { success: false, message: e.toString() };
  }
}

function validarAccesoExamen(dni) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dniLimpio = dni.toString().replace(/\D/g, "");
    
    // 1. Buscar al alumno para saber su fecha de examen
    const hojaInscritos = ss.getSheetByName(SHEETS.INSCRIPCIONES);
    const datosInscritos = hojaInscritos.getDataRange().getValues();
    const alumno = datosInscritos.find(fila => fila[COL_INS.DNI].toString().replace(/\D/g, "") === dniLimpio);

    if (!alumno) return { habilitado: false, mensaje: "DNI no encontrado en el sistema." };
    
    // 2. Verificar Asistencia (Opcional, pero recomendado)
    if (alumno[COL_INS.ASISTENCIA] < 2) { // Columna N es Asistencia
       return { habilitado: false, mensaje: "No posees la asistencia mínima requerida para rendir." };
    }

    // 3. Obtener la fecha de examen del alumno (Columna M - índice 12)
    const fechaExamenAlumno = alumno[COL_INS.BARRIOF_EXAMEN]; 
    if (!(fechaExamenAlumno instanceof Date)) {
      return { habilitado: false, mensaje: "No tienes una fecha de examen asignada aún." };
    }

    // 4. Obtener franja horaria de Configuración (B1 y B2)
    const conf = ss.getSheetByName("Configuracion") || ss.getSheetByName("Config");
    const horaInicioStr = conf.getRange("B1").getDisplayValue(); // Ej: "08:00"
    const horaFinStr = conf.getRange("B2").getDisplayValue();    // Ej: "18:00"

    // 5. Construir objetos de fecha completos para comparar
    const ahora = new Date();
    
    // Fecha de inicio permitida: Fecha del examen + hora de B1
    const inicioHabilitado = new Date(fechaExamenAlumno);
    const [hI, mI] = horaInicioStr.split(":");
    inicioHabilitado.setHours(parseInt(hI), parseInt(mI), 0);

    // Fecha de fin permitida: Fecha del examen + hora de B2
    const finHabilitado = new Date(fechaExamenAlumno);
    const [hF, mF] = horaFinStr.split(":");
    finHabilitado.setHours(parseInt(hF), parseInt(mF), 0);

    const estaEnRango = (ahora >= inicioHabilitado && ahora <= finHabilitado);

    return {
      habilitado: estaEnRango,
      mensaje: estaEnRango ? "OK" : `El examen estará disponible el ${Utilities.formatDate(fechaExamenAlumno, "GMT-3", "dd/MM")} de ${horaInicioStr} a ${horaFinStr} hs.`,
      alumno: { nombre: alumno[COL_INS.NOMBREBRE], apellido: alumno[COL_INS.APELLIDOLLIDO] }
    };

  } catch (e) {
    return { habilitado: false, mensaje: "Error al validar acceso: " + e.toString() };
  }
}

function obtenerTodasLasPreguntas() {
  const sheet = getSheet(SHEETS.PREGUNTAS);
  const data = sheet.getDataRange().getValues();

  return data.slice(1)
    .filter(r => r[COL_PREG.PREGUNTA])
    .map((r, i) => {

      let imagen = null;
      const imgIdRaw = r[COL_PREG.IMAGEN];

      if (imgIdRaw) {
        const imgId = imgIdRaw.toString().trim();

        if (
          imgId.length >= 20 &&
          /^[a-zA-Z0-9_-]+$/.test(imgId) &&
          !imgId.includes("obtenerUrlImagen")
        ) {
          //imagen = `https://drive.google.com/uc?export=view&id=${imgId}`;
          imagen = `https://lh3.googleusercontent.com/d/${imgId}`;
        } else {
          Logger.log(`⚠ Imagen inválida en fila ${i + 2}: ${imgId}`);
        }
      }

      return {
        id: i + 1,
        texto: r[COL_PREG.PREGUNTA],
        imagen: imagen,

        opciones: {
          a: r[COL_PREG.OPC1],
          b: r[COL_PREG.OPC2],
          c: r[COL_PREG.OPC3]
        },

        correcta: (r[COL_PREG.CORRECTA] || "")
          .toString().toLowerCase().trim(),

        puntos: Number(r[COL_PREG.PUNTOS]) || 1,
        tiempo: Number(r[COL_PREG.TIEMPO]) || 30,
        excluyente:
          (r[COL_PREG.EXCLUYENTE] || "")
            .toString()
            .toUpperCase() === "SI"
      };
    });
}