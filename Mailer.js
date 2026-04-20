// Mailer.gs 

/** ESTA ES LA FUNCIÓN QUE DEBES VINCULAR AL TRIGGER (Reloj) */
function ejecutarTareaDiariaMailer() {
  procesarRecordatoriosDiarios();
}


function enviarMailConfirmacion(email, datos, info) {
  if (!email || !email.includes('@')) return;

  // 1. ASIGNACIÓN DE VARIABLES (Mapeo de tu objeto 'datos')
  const nombre = datos.nombre || datos[COL_INS.NOMBRE] || "Aspirante";
  const categoria = datos.cat || datos.categoria || datos[COL_INS.CATEGORIA] || "No especificada"; 
  const institucion = datos.inst || datos.institucion || datos[COL_BAR.INST] || "Sede Municipal"; 
  const dniAlu = datos.dni || datos[COL_INS.DNI] || "0"; // AQUÍ SE DEFINE PARA EVITAR EL ERROR

  // 2. GENERACIÓN DE URL
  const urlApp = ScriptApp.getService().getUrl();
  const linkCancelacion = `${urlApp}?action=cancelar&dni=${dniAlu}&step=confirmar`;
  
  const manual1 = "https://drive.google.com/file/d/1LbX6zDe1o9XOBpkAF8dMG3rgrU31n79V/view?usp=drive_link";
  const manual2 = "https://drive.google.com/file/d/1sclr-s1lcVy9aKfPTAEAlyTfq4GNT0DE/view?usp=drive_link";
  
  const f = (val) => val instanceof Date ? Utilities.formatDate(val, "GMT-3", "dd/MM/yyyy") : "A confirmar";
  
  const asunto = `✅ ¡Hola ${nombre}! Tu lugar en ${institucion} está reservado`;

  const cuerpo = `
    <div style="font-family: sans-serif; padding: 25px; border: 2px solid #1565c0; border-radius: 15px; max-width: 600px; margin: auto; color: #333;">
      <h2 style="color: #1565c0; text-align: center;">¡Inscripción Confirmada!</h2>
      <p>Hola <strong>${nombre}</strong>,</p>
      <p>Te confirmamos tu lugar para el curso de <strong>${categoria}</strong> en la sede <strong>${institucion}</strong>.</p>
      
      <div style="background-color: #f1f8e9; padding: 20px; border-radius: 10px; border: 1px solid #c5e1a5; margin: 20px 0;">
        <h4 style="margin-top: 0; color: #2e7d32; text-align: center;">📚 Material de Estudio Obligatorio</h4>
        <div style="text-align: center;">
          <a href="${manual1}" style="display: inline-block; background-color: #2e7d32; color: white; padding: 10px 15px; text-decoration: none; border-radius: 5px; font-weight: bold; margin: 5px;">📥 CURSO AUTOS</a>
          <a href="${manual2}" style="display: inline-block; background-color: #2e7d32; color: white; padding: 10px 15px; text-decoration: none; border-radius: 5px; font-weight: bold; margin: 5px;">📥 CURSO MOTOS</a>
        </div>
      </div>

      <p><strong>📅 Cronograma de encuentros:</strong></p>
      <ul style="list-style: none; padding: 0;">
        <li>🔹 <b>Clase 1:</b> ${f(info.fecha1)}</li>
        <li>🔹 <b>Clase 2:</b> ${f(info.fecha2)}</li>
        <li>🚩 <b>EXAMEN FINAL: ${f(info.fechaExamen)}</b></li>
      </ul>

      <hr style="border: 0; border-top: 1px solid #eee; margin: 25px 0;">
      <p style="font-size: 0.8em; color: #888; text-align: center;">
        ¿Deseas cancelar tu inscripción? <br><br>
        <a href="${linkCancelacion}" style="color: #d32f2f; font-weight: bold; text-decoration: underline;">Hacer clic aquí para liberar el cupo</a>
      </p>
    </div>`;

  MailApp.sendEmail({ to: email, subject: asunto, htmlBody: cuerpo });
}

function procesarRecordatoriosDiarios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.INSCRIPCIONES); // Usamos tu constante
  const data = sheet.getDataRange().getValues();
  
  const hoy = new Date();
  const manana = new Date(hoy);
  manana.setDate(hoy.getDate() + 1);
  const mananaStr = Utilities.formatDate(manana, "GMT-3", "dd/MM/yyyy");

  // Columnas K(10), L(11), M(12) en base 0
  const COL_FECHAS = [10, 11, 12]; 

  for (let i = 1; i < data.length; i++) {
    const fila = data[i];
    const email = fila[COL_INS.EMAIL]; // Usa tus constantes para evitar errores
    const nombre = fila[COL_INS.NOMBRE];
    const sede = fila[COL_BAR.INST];
    const barrio = fila[COL_INS.BARRIO] || ""; // Asegúrate de tener esta col definida

    // Validación de email: Si no hay @ o es muy corto, saltar
    if (!email || email.toString().indexOf('@') === -1) continue;

    COL_FECHAS.forEach(col => {
      let fechaEvento = fila[col];
      if (fechaEvento instanceof Date) {
        let fechaEventoStr = Utilities.formatDate(fechaEvento, "GMT-3", "dd/MM/yyyy");
        
        if (fechaEventoStr === mananaStr) {
          let esExamen = (col === 12);
          enviarMailRecordatorio_(email, nombre, esExamen, sede, barrio, fechaEventoStr);
        }
      }
    });
  }
}

function enviarMailRecordatorio_(email, nombre, esExamen, sede, barrio, fecha) {
  const titulo = esExamen ? "¡MAÑANA ES TU EXAMEN!" : "Recordatorio de Cursada";
  const colorPrincipal = esExamen ? "#d32f2f" : "#1565c0";
  
  const cuerpo = `
    <div style="font-family: sans-serif; border: 1px solid #eee; padding: 25px; border-radius: 15px; max-width: 600px; margin: auto; color: #333;">
      <h2 style="color: ${colorPrincipal}; text-align: center;">${titulo}</h2>
      <p>Hola <strong>${nombre}</strong>,</p>
      <p>Te recordamos que tienes una cita programada para el día de mañana:</p>
      <div style="background-color: #f8f9fa; padding: 15px; border-radius: 10px; border-left: 5px solid ${colorPrincipal};">
        <p style="margin: 0;"><strong>📅 Fecha:</strong> ${fecha}</p>
        <p style="margin: 5px 0 0 0;"><strong>📍 Lugar:</strong> ${barrio} - ${sede}</p>
      </div>
      <p style="margin-top: 20px; font-size: 0.9em; color: #666; text-align: center;">
        <i>Por favor, recuerda llevar tu DNI y ser puntual.</i>
      </p>
    </div>
  `;

  try {
    MailApp.sendEmail({
      to: email,
      subject: `⏰ Recordatorio: ${titulo}`,
      htmlBody: cuerpo
    });
  } catch (e) {
    console.error("Error enviando a: " + email + " Error: " + e.toString());
  }
}

function enviarCorreoResultado(email, nombre, nota, aprobado, excluyente) {
  try {
    const urlApp = ScriptApp.getService().getUrl(); // Obtenemos la URL de tu sistema
    const subject = aprobado ? "¡Felicitaciones! Aprobaste el Examen Teórico" : "Resultado de tu Examen Teórico";
    const colorFondo = aprobado ? "#2e7d32" : "#c62828";
    
    // CORRECCIÓN: El linkBoton debe ser solo la URL, no una etiqueta <a> completa
    const linkBoton = aprobado 
      ? "https://licencias.mercedes.gob.ar/" 
      : urlApp + "?p=inscripcion"; // Lo mandamos directo a la página de inscripción
    
    const textoBoton = aprobado ? "SOLICITAR TURNO PRÁCTICO" : "VOLVER A INSCRIBIRSE";

    let detalleMensaje = aprobado 
      ? "Has superado con éxito la instancia teórica. Ya puedes avanzar al examen práctico." 
      : "Lamentablemente no has alcanzado los requisitos mínimos para aprobar.";
    
    if (excluyente && !aprobado) {
      detalleMensaje += "<br><br><strong>Nota:</strong> Se detectó un error en una pregunta de seguridad vial obligatoria (Excluyente).";
    }

    const fechaHoy = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy");

    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 550px; margin: auto; border: 1px solid #ddd; border-radius: 15px; overflow: hidden;">
        <div style="background-color: ${colorFondo}; color: white; padding: 25px; text-align: center;">
          <h1 style="margin: 0; font-size: 24px;">${aprobado ? '¡APROBADO!' : 'REPROBADO'}</h1>
        </div>
        <div style="padding: 30px; line-height: 1.6; color: #333;">
          <p>Hola <strong>${nombre}</strong>,</p>
          <p>Informamos el resultado de tu examen realizado el ${fechaHoy}:</p>
          <div style="background: #f9f9f9; padding: 15px; border-radius: 8px; text-align: center; margin: 20px 0; border: 1px solid #eee;">
            <span style="font-size: 1rem; color: #666;">Calificación Final:</span><br>
            <strong style="font-size: 2.8rem; color: ${colorFondo};">${nota}%</strong>
          </div>
          <p>${detalleMensaje}</p>
          <div style="text-align: center; margin-top: 35px;">
            <a href="${linkBoton}" style="background-color: ${colorFondo}; color: white; padding: 16px 30px; text-decoration: none; border-radius: 8px; font-weight: bold; display: inline-block;">
              ${textoBoton}
            </a>
          </div>
        </div>
        <div style="background: #f1f1f1; padding: 15px; text-align: center; font-size: 12px; color: #777;">
          <strong>Municipalidad de Mercedes</strong><br>
          Dirección de Seguridad Vial y Licencias de Conducir
        </div>
      </div>
    `;

    MailApp.sendEmail({ to: email, subject: subject, htmlBody: htmlBody });
  } catch (e) {
    console.error("Error en enviarCorreoResultado: " + e.message);
  }
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
