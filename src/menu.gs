/**
 * CertiFlow - Librería para generación y envío de certificados digitales
 * © 2025 Oscar Giovanni Castro Contreras
 * 
 * Licencia dual:
 * - MIT License (LICENSE-MIT)
 * - GNU GPLv3 (LICENSE-GPL)
 * 
 * El usuario puede optar por cualquiera de estas licencias.
 */

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

function mostrarModalEnviarPorFilas() {
  const html = HtmlService.createHtmlOutputFromFile('modal_send_rows')
    .setWidth(450)
    .setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'Enviar Certificados');
}

function mostrarModalEnviarPorRango() {
  const html = HtmlService.createHtmlOutputFromFile('modal_send_range')
    .setWidth(450)
    .setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'Enviar Certificados');
}