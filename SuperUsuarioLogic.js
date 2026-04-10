// superUsuario.gs

function crearNuevoCurso(datos, dniOperador) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("BarriosInstituciones");

  // A:Barrio | B:Inst | C:Dir | D:Cupo | E:F1 | F:F2 | G:FEx | H:Estado | I:Actuales | J:Horario
  sheet.appendRow([
    datos.barrio,
    datos.institucion,
    datos.direccion,
    datos.cupo,
    datos.fecha1,
    datos.fecha2,
    datos.fechaEx,
    "Activo",
    0, 
    datos.horario
  ]);

  registrarAccion(dniOperador, `ALTA SEDE: ${datos.institucion} - Horario: ${datos.horario}`);
  return { success: true };
}

function actualizarSede(nombreOriginal, datosNuevos, dniOperador) {
  const sheet = getSheet(SHEETS.BARRIOS);
  const data = sheet.getDataRange().getValues();
  let encontrada = false;

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === nombreOriginal) {
      const fila = i + 1;
      
      // Actualizamos los datos básicos de la sede en Barrios/Instituciones
      // Nota: Usamos setValues en una sola línea para ser más eficientes
      const valoresSede = [[
        datosNuevos.barrio,
        datosNuevos.institucion,
        datosNuevos.direccion,
        datosNuevos.cupo
      ]];
      sheet.getRange(fila, 1, 1, 4).setValues(valoresSede);
      
      // Actualizamos horario (Col J = 10) y fechas si vienen en datosNuevos
      sheet.getRange(fila, 10).setValue(datosNuevos.horario);
      if(datosNuevos.fecha1) sheet.getRange(fila, 5).setValue(datosNuevos.fecha1);
      if(datosNuevos.fecha2) sheet.getRange(fila, 6).setValue(datosNuevos.fecha2);
      if(datosNuevos.fechaEx) sheet.getRange(fila, 7).setValue(datosNuevos.fechaEx);

      encontrada = true;
      break; 
    }
  }

  if (encontrada) {
    // AHORA SÍ: Sincronizamos a los alumnos después de salir del bucle
    actualizarNombreSedeEnAlumnos(nombreOriginal, datosNuevos.institucion, {
      fecha1: datosNuevos.fecha1,
      fecha2: datosNuevos.fecha2,
      fechaEx: datosNuevos.fechaEx
    });

    registrarAccion(dniOperador, "ACTUALIZACIÓN SEDE Y SINCRONIZACIÓN", `De ${nombreOriginal} a ${datosNuevos.institucion}`);
    return { success: true };
  }

  return { success: false, error: "No se encontró la sede original." };
}

function eliminarSedeServidor(institucion, dniOperador) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSedes = ss.getSheetByName("BarriosInstituciones");
  const sheetIns = ss.getSheetByName("Inscripciones");
  
  // 1. Verificar si hay alumnos inscriptos en esa sede
  const alumnos = sheetIns.getDataRange().getValues();
  const tieneAlumnos = alumnos.some(fila => fila[9].toString().trim() === institucion.trim());
  
  if (tieneAlumnos) {
    return { 
      success: false, 
      message: "No se puede eliminar: Hay alumnos inscriptos en esta sede. Primero cámbialos de sede o elimínalos." 
    };
  }

  // 2. Si está vacía, proceder a eliminar la fila
  const sedes = sheetSedes.getDataRange().getValues();
  for (let i = 1; i < sedes.length; i++) {
    if (sedes[i][1].toString().trim() === institucion.trim()) {
      sheetSedes.deleteRow(i + 1);
      registrarAccion(dniOperador, `ELIMINÓ SEDE PERMANENTEMENTE: ${institucion}`);
      return { success: true, message: "Sede eliminada correctamente." };
    }
  }
  
  return { success: false, message: "Sede no encontrada." };
}

function obtenerTodasLasSedes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("BarriosInstituciones");
  const data = sheet.getDataRange().getValues();
  data.shift(); // Quitar cabeceras

  // Obtener inscriptos reales para cada sede desde la hoja Inscripciones
  const inscripciones = ss.getSheetByName("Inscripciones").getDataRange().getValues();
  const conteo = {};
  inscripciones.shift();
  inscripciones.forEach(r => {
    conteo[r[9]] = (conteo[r[9]] || 0) + 1;
  });

  return data.map(r => ({
    barrio: r[0],
    institucion: r[1],
    direccion: r[2],
    cupo: r[3],
    fecha1: r[4],
    fecha2: r[5],
    fechaEx: r[6],
    estado: r[7],
    actuales: conteo[r[1]] || 0,
    horario: r[9]
  }));
}

function obtenerInscriptosPorSede(nombreSede) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Inscripciones");
  const data = sheet.getDataRange().getValues();
  
  // Filtrar alumnos que pertenezcan a la sede (Columna J - índice 9)
  const filtrados = data.filter((fila, index) => {
    if (index === 0) return false; // Omitir cabecera
    return fila[9].toString().trim() === nombreSede.trim();
  });

  if (filtrados.length === 0) return { success: false, message: "No hay alumnos inscriptos en esta sede." };

  // Formatear datos para el reporte
  const reporte = filtrados.map(f => ({
    Apellido: f[2],
    Nombre: f[1],
    DNI: f[3],
    Telefono: f[5],
    Categoria: f[7],
    Asistencia: f[13] + "%"
  }));

  return { success: true, datos: reporte, sede: nombreSede };
}

function obtenerEstadoCupos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sedesData = ss.getSheetByName("BarriosInstituciones").getDataRange().getValues();
  const alumnosData = ss.getSheetByName("Inscripciones").getDataRange().getValues();
  
  sedesData.shift(); // Quitar cabecera de sedes
  alumnosData.shift(); // Quitar cabecera de alumnos
  
  // Crear un mapa de conteo: { "Nombre Sede": cantidad }
  const conteoAlumnos = {};
  alumnosData.forEach(fila => {
    const sede = fila[9]; // Columna J: Institución
    if (sede) {
      conteoAlumnos[sede] = (conteoAlumnos[sede] || 0) + 1;
    }
  });

  // Mapear los resultados finales
  return sedesData.map(r => {
    const sedeNombre = r[1];
    const cupoMax = parseInt(r[3]) || 0;
    const inscriptosReal = conteoAlumnos[sedeNombre] || 0;
    
    return {
      barrio: r[0],
      sede: sedeNombre,
      max: cupoMax,
      actual: inscriptosReal,
      disponible: cupoMax - inscriptosReal
    };
  });
}

function actualizarNombreSedeEnAlumnos(viejoNombre, nuevoNombre, nuevasFechas = null) {
  const sheet = getSheet(SHEETS.INSCRIPCIONES);
  const range = sheet.getDataRange();
  const data = range.getValues();
  let huboCambios = false;

  // Recorremos los datos en memoria (empezando desde la fila 1 para saltar cabecera)
  for (let i = 1; i < data.length; i++) {
    if (data[i][COL_INS.INST] === viejoNombre) {
      // 1. Actualizamos el nombre de la institución
      data[i][COL_INS.INST] = nuevoNombre;
      
      // 2. Si se pasaron nuevas fechas, las sincronizamos de una vez
      if (nuevasFechas) {
        data[i][COL_INS.F1] = nuevasFechas.fecha1;
        data[i][COL_INS.F2] = nuevasFechas.fecha2;
        data[i][COL_INS.F_EX] = nuevasFechas.fechaEx;
      }
      huboCambios = true;
    }
  }

  // Solo escribimos en la hoja si realmente encontramos alumnos afectados
  if (huboCambios) {
    range.setValues(data); 
    console.log(`Sincronización completada para la sede: ${nuevoNombre}`);
  }
}

function obtenerSedesActivas() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("BarriosInstituciones");
    const data = sheet.getDataRange().getValues();
    data.shift(); 

    return data
      .filter(r => r[7] && r[7].toString().toUpperCase().trim() === "ACTIVO")
      .map(r => ({
        barrio: r[0],
        institucion: r[1]
      }));
  } catch (e) {
    console.error("Error en obtenerSedesActivas: " + e.message);
    return [];
  }
}

function obtenerDatosEdicionCompleta(dni) {
  try {
    const resMesa = obtenerDatosMesa(dni); // Reutiliza tu lógica de búsqueda
    if (!resMesa.success) return resMesa;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sedesData = ss.getSheetByName("BarriosInstituciones").getDataRange().getValues();
    const sedes = sedesData.slice(1)
      .filter(r => r[1] !== "")
      .map(r => ({ barrio: r[0], nombre: r[1] }));

    return { success: true, data: resMesa.data, sedes: sedes };
  } catch (e) {
    return { success: false, message: "Error al cargar ficha: " + e.toString() };
  }
}

function actualizarDatosAlumno(dniOriginal, datos, dniOperador) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetIns = getSheet(SHEETS.INSCRIPCIONES);
    const dataIns = sheetIns.getDataRange().getValues();
    const dniBusqueda = dniOriginal.toString().trim().replace(/\D/g, "");

    for (let i = 1; i < dataIns.length; i++) {
      if (dataIns[i][COL_INS.DNI].toString().trim().replace(/\D/g, "") === dniBusqueda) {
        const fila = i + 1;
        let logDetalle = `DNI: ${dniBusqueda}`;

        // Caso 1: Edición integral desde el formulario
        if (datos.nombre) {
          const sedesData = getSheet(SHEETS.BARRIOS).getDataRange().getValues();
          const configSede = sedesData.find(r => 
            r[COL_BAR.INST].toString().trim() === datos.institucion.toString().trim()
          );
          
          let barrio = dataIns[i][COL_INS.BARRIO];
          let fechasExamen = [
            dataIns[i][COL_INS.F1],
            dataIns[i][COL_INS.F2],
            dataIns[i][COL_INS.F_EX]
          ];

          // Si el admin cambió la sede, actualizamos automáticamente las fechas de esa cursada
          if (configSede) {
            barrio = configSede[COL_BAR.BARRIO];
            fechasExamen = [
              configSede[COL_BAR.F1],
              configSede[COL_BAR.F2],
              configSede[COL_BAR.F_EX]
            ];
          }

          let fechaNac = dataIns[i][COL_INS.FNAC];

          if (datos.fechaNac && datos.fechaNac !== "") {
            const f = new Date(datos.fechaNac);
            if (!isNaN(f.getTime())) {
              fechaNac = f;
            }
          }

        const dniLimpio = datos.dni 
          ? datos.dni.toString().replace(/\D/g, "")
          : dataIns[i][COL_INS.DNI];

          // ESCRITURA EN BLOQUE 1: Datos Personales (Columnas B a J)          
          sheetIns.getRange(fila, COL_INS.NOM + 1, 1, 9).setValues([[
            datos.nombre, 
            datos.apellido, 
            dniLimpio,
            fechaNac, 
            dataIns[i][COL_INS.TEL],
            dataIns[i][COL_INS.EMAIL],
            dataIns[i][COL_INS.CAT],
            barrio,
            datos.institucion
          ]]);

          // ESCRITURA EN BLOQUE 2: Cursada y Notas (Columnas iNSCRIPCION K a P)
          sheetIns.getRange(fila, COL_INS.F1 + 1, 1, 6).setValues([[
            fechasExamen[0], 
            fechasExamen[1], 
            fechasExamen[2], 
            datos.asistencia || 0, 
            datos.nota || dataIns[i][COL_INS.NOTA], // Mantenemos nota previa si no viene nueva
            dataIns[i][COL_INS.ESTADO]
          ]]);
          
          logDetalle += ` - Edición integral - Nueva Sede: ${datos.institucion}`;
        }

        // Caso 2: Acciones rápidas (asistencia o reset de examen) desde el panel
        if (datos.resetearExamen) {
          sheetIns.getRange(fila, COL_INS.NOTA + 1).setValue("HABILITADO");
          logDetalle += " - Reset Examen (HABILITADO)";
        }
        
        if (datos.ponerPresente) {
          sheetIns.getRange(fila, COL_INS.ASIST + 1).setValue(100);
          logDetalle += " - Asistencia forzada al 100%";
        }

        registrarAccion(dniOperador, "GESTIÓN ADM", logDetalle);
        return { success: true, message: "Actualización exitosa para " + datos.apellido };
      }
    }
    return { success: false, message: "Alumno no encontrado en la base de datos." };
  } catch (e) {
    console.error("Error en actualizarDatosAlumno: " + e.toString());
    return { success: false, message: "Error crítico: " + e.toString() };
  }
}