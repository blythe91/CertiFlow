function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Certificados')
    .addItem('Generar Todos Los Certificados', 'mostrarModalCertificados')
    .addToUi();
}

function mostrarModalCertificados() {
  const html = HtmlService.createHtmlOutputFromFile('modal_certificados')
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
  Logger.log("URL de hoja de cálculo: " + templateUrl);
  Logger.log("ID de hoja de cálculo: " + templateId);
  Logger.log("////////////////");
  Logger.log("URL de hoja de cálculo: " + folderUrl);
  Logger.log("ID de hoja de cálculo: " + folderId);

  if (!sheetId || !templateId || !folderId) {
    throw new Error("Alguno de los IDs no pudo extraerse correctamente.");
  }

  //generarCertificados(sheetId, templateId, folderId, batchSize);
}


function detenerGeneracionCertificados() {
  eliminarTrigger();
  PropertiesService.getScriptProperties().deleteAllProperties();
}
