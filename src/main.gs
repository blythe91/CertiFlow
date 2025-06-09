function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Certificados')
    .addSubMenu(
      ui.createMenu('Generar Certificados')
        .addItem('Todos', 'mostrarModalCertificados')
        .addItem('Por Filas', 'mostrarModalCertificadosPorFilas')
    )
    .addToUi();
}


function mostrarModalCertificados() {
  const html = HtmlService.createHtmlOutputFromFile('modal_certificados')
    .setWidth(450)
    .setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generar Certificados');
}

function mostrarModalCertificadosPorFilas() {
  const html = HtmlService.createHtmlOutputFromFile('modal_cert_filas')
    .setWidth(450)
    .setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generar Certificados');
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


function detenerGeneracionCertificados() {
  eliminarTrigger();

  PropertiesService.getScriptProperties().deleteAllProperties();
  PropertiesService.getScriptProperties().deleteProperty("lastProcessedIndex");
  PropertiesService.getScriptProperties().deleteProperty("totalCertificados");

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