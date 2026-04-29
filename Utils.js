//Utils.gs
//Funciones de apoyo reusables
function FORZAR_PERMISO_DRIVE() {
  // Esta línea es distinta y obliga a Google a pedir permiso de escritura total
  const nombreTemporal = "TEST_" + new Date().getTime();
  const carpeta = DriveApp.createFolder(nombreTemporal);
  Utilities.sleep(1000);
  carpeta.setTrashed(true); // Lo borramos enseguida
  console.log("✅ PERMISO DE ESCRITURA CONFIRMADO EN DRIVE");
}

function normalizeDNI(dni) {
  if (!dni) return "";
  return String(dni).replace(/[^0-9]/g, '');
}

function findColumnByName(sheet, columnName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[COL_INS.ID];
  const index = headers.indexOf(columnName);
  return index !== -1 ? index + 1 : 1; // Retorna la columna 1 si no la encuentra para evitar error
}

function logError(contexto, error) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("LogsErrores") 
    || SpreadsheetApp.getActive().insertSheet("LogsErrores");

  sheet.appendRow([
    new Date(),
    contexto,
    error.message || error,
    JSON.stringify(error.stack || "")
  ]);
}

function generarPDFAlumnosFiltrados(listaAlumnos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const conf = ss.getSheetByName("Configuracion");
  
  // Datos dinámicos de la Dirección de Tránsito
  const firmaTexto = conf.getRange("B3").getValue(); // "Dirección de Tránsito - Municipalidad de Mercedes"
  const logoUrl = conf.getRange("B4").getValue();    // "https://nw.mercedes.gob.ar/img/logo.svg"

  let html = `
    <div style="font-family: sans-serif; padding: 20px;">
      <table style="width: 100%; border-bottom: 2px solid #1a237e;">
        <tr>
          <td><img src="${logoUrl}" style="width: 100px;"></td>
          <td style="text-align: right;">
            <h2 style="color: #1a237e; margin:0;">Reporte de Examen</h2>
            <p style="margin:0;">Mercedes, Buenos Aires</p>
          </td>
        </tr>
      </table>
      
      <table style="width:100%; border-collapse: collapse; margin-top: 20px;">
        <tr style="background-color: #1a237e; color: white;">
          <th style="padding: 10px; border: 1px solid #ddd;">Alumno</th>
          <th style="padding: 10px; border: 1px solid #ddd;">DNI</th>
          <th style="padding: 10px; border: 1px solid #ddd;">Asistencia</th>
          <th style="padding: 10px; border: 1px solid #ddd;">Nota</th>
        </tr>`;

  listaAlumnos.forEach(al => {
    html += `
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd;">${al.Apellido}, ${al.Nombre}</td>
        <td style="padding: 8px; border: 1px solid #ddd; text-align: center;">${al.dni}</td>
        <td style="padding: 8px; border: 1px solid #ddd; text-align: center;">${al['Porcentaje de asistencia']}%</td>
        <td style="padding: 8px; border: 1px solid #ddd; text-align: center;">${al.Calificación || '-'}</td>
      </tr>`;
  });

  html += `
      </table>

      <div style="margin-top: 40px;">
        <p><b>Observaciones:</b> __________________________________________________________________</p>
      </div>

      <div style="margin-top: 80px; text-align: right;">
        <div style="display: inline-block; border-top: 1px solid #000; width: 250px; text-align: center;">
          <p style="margin: 5px 0 0 0;">${firmaTexto}</p>
        </div>
      </div>
    </div>`;

  const blob = HtmlService.createHtmlOutput(html).getAs('application/pdf');
  return Utilities.base64Encode(blob.getBytes());
}

function generarCertificadoAlumno(dni, dniOperador) {
  const res = obtenerDatosMesa(dni);
  if (!res.success) return { success: false, message: "Error al obtener datos: " + res.message };
  
  const al = res.data;

  // 1. VALIDACIÓN DE APROBACIÓN
  const notaValor = al.estadoExamen; 
  const notaNum = parseInt(notaValor);
  
  // Es aprobado si dice "APROBADO" o si es un número >= 75
  const estaAprobado = (!isNaN(notaNum) && notaNum >= 75);

  if (!estaAprobado) {
    return { 
      success: false, 
      message: "El alumno no cumple los requisitos. Estado actual: " + (notaValor || "Sin calificar") 
    };
  }

  // 2. CONFIGURACIÓN E INSTITUCIÓN
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const conf = ss.getSheetByName(SHEETS.CONFIG);
  const logo = conf ? conf.getRange("B4").getValue() : "";
  const institucion = conf ? conf.getRange("B3").getValue() : "Municipalidad de Mercedes";

  // 3. HTML DEL PDF
  const html = `
    <html>
      <body style="font-family: sans-serif; margin:0; padding:0;">
        <div style="border: 15px solid #1565c0; padding: 50px; text-align: center; min-height: 842px; box-sizing: border-box;">
          <img src="${logo}" style="max-height: 80px; margin-bottom: 20px;">
          <h1 style="color: #1565c0; font-size: 32px; margin-bottom: 5px;">CERTIFICADO DE APROBACIÓN</h1>
          <h3 style="color: #666; margin-bottom: 30px; letter-spacing: 2px;">PROGRAMA LICENCIA EN TU BARRIO</h3>

          <p style="font-size: 1.4em; margin-bottom: 10px;">Se deja constancia que:</p>
          <h2 style="text-transform: uppercase; font-size: 36px; color: #1a237e; margin: 10px 0;">${al.apellido}, ${al.nombre}</h2>
          <p style="font-size: 1.3em;">DNI: <b>${al.dni}</b></p>

          <div style="width: 70%; border-top: 2px solid #1565c0; margin: 30px auto;"></div>

          <p style="font-size: 1.2em; line-height: 1.6; padding: 0 40px;">
            Ha cumplido satisfactoriamente con la capacitación teórica y evaluación correspondiente, obteniendo una calificación de <b>${notaNum}%</b>, para la categoría:
          </p>
          <h2 style="color: #2e7d32; font-size: 28px; margin-top: 10px;">${al.categoria || 'Particular'}</h2>
          
          <p style="color: #555; margin-top: 20px;">
            <b>Sede:</b> ${al.institucion} <br> 
            <b>Fecha de Emisión:</b> ${Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy")}
          </p>

          <div style="margin-top: 60px;">
            <p style="margin-bottom: 0;">__________________________</p>
            <p style="font-weight: bold; margin-top: 5px; color: #333;">${institucion}</p>
            <p style="font-size: 0.75em; color: #999;">Documento validado por DNI Operador: ${dniOperador || 'Sistema Automático'}</p>
          </div>
        </div>
      </body>
    </html>
  `;

  // 4. CREACIÓN DEL ARCHIVO
  const blob = Utilities.newBlob(html, "text/html", "temp.html").getAs("application/pdf").setName(`Certificado_${al.dni}.pdf`);
  const archivo = DriveApp.createFile(blob);
  
  // Dar permisos para que el link sea accesible
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // 5. TRAZABILIDAD Y CIERRE
  if (typeof registrarAccion === "function") {
    registrarAccion(dniOperador, "EMITIÓ CERTIFICADO", `DNI Alumno: ${dni}`);
  }
  
  if (typeof marcarTramiteFinalizado === "function") {
    marcarTramiteFinalizado(dni, dniOperador);
  }

  return { 
    success: true, 
    url: archivo.getDownloadUrl(),
    base64: Utilities.base64Encode(blob.getBytes()) 
  };
}


function normalizarFecha(fecha) {
  if (!fecha) return "";

  const f = new Date(fecha);

  if (isNaN(f.getTime())) return "";

  return Utilities.formatDate(
    f,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );
}

function limpiarHojaInscripciones() {
  const sheet = getSheet(SHEETS.INSCRIPCIONES);
  const data = sheet.getDataRange().getValues();

  const limpia = data.map((r, i) => {
    if (i === 0) return r;

    return r.map(c =>
      typeof c === "string"
        ? c.trim()
        : c
    );
  });

  sheet.getRange(1,1,limpia.length,limpia[0].length).setValues(limpia);

  Logger.log("✅ Hoja INSCRIPCIONES limpiada");
}

// TESTS DE FUNCIONES CLAVE PARA ASEGURAR INTEGRIDAD Y PERMISOS

function testCertificadoPDF() {
  const dniParaProbar = "12345678"; // <--- CAMBIA POR UN DNI QUE HAYA APROBADO
  
  console.log("Generando certificado para: " + dniParaProbar);
  
  try {
    const resultado = generarCertificadoAlumno(dniParaProbar);
    
    if (resultado.success) {
      console.log("✅ ÉXITO: Certificado generado.");
      console.log("🔗 URL del archivo en Drive: " + resultado.url);
      console.log("Si entras al link, deberías ver el PDF.");
    } else {
      console.warn("❌ FALLO: " + resultado.message);
    }
  } catch (e) {
    console.error("🚨 ERROR DE CÓDIGO: " + e.message);
  }
  
}

function testFull() {
  console.log("COL_INS:", COL_INS.DNI);
  console.log("COL_PER:", COL_PER.DNI);
}

function testEnvioMail() {
  const miEmail = "pollolopeza@gmail.com"; 
  
  const datosTest = {
    nombre: "Pedro Perez",
    cat: "Clase B1",   // Usamos 'cat' como en tu procesarNuevaInscripcion
    inst: "Sede Centro", // Usamos 'inst' como en tu procesarNuevaInscripcion
    dni: "12345678"
  };
  
  const infoTest = {
    fecha1: new Date(),
    fecha2: new Date(),
    fechaExamen: new Date()
  };
  
  enviarMailConfirmacion(miEmail, datosTest, infoTest);
}

function testGlobals() {
  console.log(COL_INS.DNI);
  console.log(COL_BAR.INST);
  console.log(COL_PER.DNI);
}

function testMesaReal() {
  Logger.log(obtenerDatosMesa("TU_DNI_REAL"));
}

function testPreguntas() {
  const data = getSheet(SHEETS.PREGUNTAS).getDataRange().getValues();
  Logger.log(data.length);
}

// 20.4.26 - Testeo de funciones clave del sistema para asegurar integridad y permisos.
function testSistemaCompleto() {
  Logger.log("===== TEST SISTEMA =====");

  // 1. PERSONAL
  try {
    const personal = getSheet(SHEETS.PERSONAL).getDataRange().getValues();
    Logger.log("PERSONAL OK: " + personal.length + " filas");
  } catch(e) {
    Logger.log("ERROR PERSONAL: " + e);
  }

  // 2. LOGIN
  try {
    const testLogin = loginPersonal("12345678", "12345"); // ⚠ poné un usuario real
    Logger.log("LOGIN: " + JSON.stringify(testLogin));
  } catch(e) {
    Logger.log("ERROR LOGIN: " + e);
  }

  // 3. INSCRIPCIONES
  try {
    const insc = getSheet(SHEETS.INSCRIPCIONES).getDataRange().getValues();
    Logger.log("INSCRIPCIONES OK: " + insc.length);
  } catch(e) {
    Logger.log("ERROR INSCRIPCIONES: " + e);
  }

  // 4. SEDES
  try {
    const sedes = obtenerOpcionesCursada();
    Logger.log("SEDES: " + JSON.stringify(sedes.slice(0,2)));
  } catch(e) {
    Logger.log("ERROR SEDES: " + e);
  }

  // 5. DATOS MESA
  try {
    const mesa = obtenerDatosMesa("12345678"); // ⚠ DNI real
    Logger.log("MESA: " + JSON.stringify(mesa));
  } catch(e) {
    Logger.log("ERROR MESA: " + e);
  }

  // 6. PREGUNTAS
  try {
    const preg = getSheet(SHEETS.PREGUNTAS).getDataRange().getValues();
    Logger.log("PREGUNTAS: " + preg.length);
  } catch(e) {
    Logger.log("ERROR PREGUNTAS: " + e);
  }

  Logger.log("===== FIN TEST =====");
}

function test() {
  obtenerAlumnosPorFiltro("Escuela N° 22");
}
