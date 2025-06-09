function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Certificados')
    .addSubMenu(
      ui.createMenu('Generar Certificados')
        .addItem('Todos', 'mostrarModalCertificados')
        .addItem('Por Filas', 'mostrarModalCertificadosPorFilas')
    )
    .addSubMenu(
      ui.createMenu('Enviar Certificados')
        .addItem('Todos', 'mostrarModalEnviarTodos')
        .addItem('Por Filas', 'mostrarModalEnviarPorFilas')
        .addItem('Por Rango de Filas', 'mostrarModalEnviarPorRango')
    )
    .addToUi();
}

mostrarModalEnviarTodos

function mostrarModalEnviarTodos() {
  const html = HtmlService.createHtmlOutputFromFile('modal_send_all')
    .setWidth(450)
    .setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'Enviar Certificados');
}

function mostrarModalCertificados() {
  const html = HtmlService.createHtmlOutputFromFile('modal_cert_all')
    .setWidth(450)
    .setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generar Certificados');
}

function mostrarModalCertificadosPorFilas() {
  const html = HtmlService.createHtmlOutputFromFile('modal_cert_rows')
    .setWidth(450)
    .setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generar Certificados');
}

function procesarYEnviarCertificados(sheetUrl, folderUrl, batchSize, mensajeEmail) {
  const sheetId = obtenerIdHojaCalculo(sheetUrl);
  const folderId = obtenerIdCarpeta(folderUrl);

  Logger.log("URL de hoja de cálculo: " + sheetUrl);
  Logger.log("ID de hoja de cálculo: " + sheetId);
  Logger.log("////////////////");
  Logger.log("URL de carpeta destino: " + folderUrl);
  Logger.log("ID de carpeta destino: " + folderId);

  if (!sheetId || !folderId) {
    throw new Error("Alguno de los IDs no pudo extraerse correctamente.");
  }

  enviarCertificadosEmail(sheetId, folderId, batchSize, mensajeEmail);
}

function procesarYGenerarCertificados(sheetUrl, templateUrl, folderUrl, batchSize) {
  const sheetId = obtenerIdHojaCalculo(sheetUrl);
  const templateId = obtenerIdPlantillaSlide(templateUrl);
  const folderId = obtenerIdCarpeta(folderUrl);

  Logger.log("URL de hoja de cálculo: " + sheetUrl);
  Logger.log("ID de hoja de cálculo: " + sheetId);
  Logger.log("////////////////");
  Logger.log("URL de la plantilla del certificado: " + templateUrl);
  Logger.log("ID de la plantilla del certificado: " + templateId);
  Logger.log("////////////////");
  Logger.log("URL de carpeta destino: " + folderUrl);
  Logger.log("ID de carpeta destino: " + folderId);

  if (!sheetId || !templateId || !folderId) {
    throw new Error("Alguno de los IDs no pudo extraerse correctamente.");
  }

  generarCertificados(sheetId, templateId, folderId, batchSize);
}


function procesarYGenerarCertificadosPorFilas(filasCSV, sheetUrl, templateUrl, folderUrl) {
  const sheetId = obtenerIdHojaCalculo(sheetUrl);
  const templateId = obtenerIdPlantillaSlide(templateUrl);
  const folderId = obtenerIdCarpeta(folderUrl);

  Logger.log("URL de hoja de cálculo: " + sheetUrl);
  Logger.log("ID de hoja de cálculo: " + sheetId);
  Logger.log("////////////////");
  Logger.log("URL de la plantilla del certificado: " + templateUrl);
  Logger.log("ID de la plantilla del certificado: " + templateId);
  Logger.log("////////////////");
  Logger.log("URL de carpeta destino: " + folderUrl);
  Logger.log("ID de carpeta destino: " + folderId);

  if (!sheetId || !templateId || !folderId) {
    throw new Error("Alguno de los IDs no pudo extraerse correctamente.");
  }

  generarCertificadosPorFilas(filasCSV, sheetId, templateId, folderId);
}

function detenerGeneracionCertificados() {
  eliminarTrigger();
  PropertiesService.getScriptProperties().deleteAllProperties();
  PropertiesService.getScriptProperties().deleteProperty("lastProcessedIndex");
  PropertiesService.getScriptProperties().deleteProperty("totalCertificados");
}