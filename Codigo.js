// Código.gs
function doGet(e) {
  try {
    const p = e.parameter.p || 'inicio';
    const dniURL = e.parameter.dni || "0";
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const conf = ss.getSheetByName("Configuracion") || ss.getSheetByName("Config");

    const instData = {
      nombre: conf ? (conf.getRange("B3").getValue() || "Secretaria Seguridad Mercedes") : "Licencia en tu barrio",
      logo: conf ? (conf.getRange("B4").getValue() || "") : "",
      url: ScriptApp.getService().getUrl()
    };

    Logger.log("URL: " + instData.url); // 👈 ahora sí

    if (e.parameter.action === 'cancelar' && e.parameter.dni) {
      const step = e.parameter.step || "confirmar";
      return manejarCancelacion(e.parameter.dni, step, instData);
    }

    const fileMap = {
      'examen': 'Examen',
      'inscripcion': 'Inscripcion',
      'instructores': 'Instructores',
      'mesa': 'MesaEntradas',
      'super': 'SuperUsuario',
      'inicio': 'Inicio'
    };

    let fileName = fileMap[p] || 'Inicio';

    const template = HtmlService.createTemplateFromFile(fileName);
    template.user = { dni: dniURL };
    template.inst = instData;

    return template.evaluate()
      .setTitle(instData.nombre)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    return HtmlService.createHtmlOutput(`<h2>Error de Sistema</h2><p>${error.message}</p>`);
  }
}

function manejarCancelacion(dni, step, instData) {
  const containerStyle = `style="font-family:sans-serif; text-align:center; padding:50px; background:#f5f5f5; min-height:100vh;"`;
  const cardStyle = `style="background:white; padding:30px; border-radius:15px; display:inline-block; box-shadow:0 4px 15px rgba(0,0,0,0.1); max-width:400px;"`;

  if (step === "confirmar") {
    const urlSi = instData.url + "?action=cancelar&dni=" + dni + "&step=ejecutar";
    return HtmlService.createHtmlOutput(`
      <div ${containerStyle}>
        <div ${cardStyle}>
          <h2 style="color:#d32f2f;">¿Cancelar inscripción?</h2>
          <p>DNI: <b>${dni}</b></p>
          <a href="${urlSi}" target="_top" style="background:#d32f2f; color:white; padding:15px 25px; text-decoration:none; border-radius:8px; font-weight:bold; display:block;">SÍ, CANCELAR MI CUPO</a>
        </div>
      </div>`).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (step === "ejecutar") {
    const resultado = cancelarInscripcion(dni);
    return HtmlService.createHtmlOutput(`
      <div ${containerStyle}>
        <div ${cardStyle}>
          <h2 style="color:#2e7d32;">Listo</h2>
          <p>${resultado}</p>
          <a href="${instData.url}" target="_top">Volver al Inicio</a>
        </div>
      </div>`).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getScriptUrl() { return ScriptApp.getService().getUrl(); }