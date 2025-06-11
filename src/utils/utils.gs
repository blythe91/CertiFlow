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

const DEFAULT_BATCH_SIZE = 30;

function obtenerBatchSizePorDefecto() {
  return DEFAULT_BATCH_SIZE;
}

function obtenerIdHojaCalculo(url) {
  // Extrae el ID de una URL de Google Sheets
  var regex = /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/;
  var match = url.match(regex);

  Logger.log("////////////////");
  Logger.log("URL de hoja de cálculo: " + url);
  Logger.log("ID de hoja de cálculo: " + (match ? match[1] : "NO DETECTADO"));
  Logger.log("////////////////");

  return match ? match[1] : null;

  
}
function obtenerIdPlantillaSlide(url) {
  // Extrae el ID de una URL de Google Slides
  var regex = /\/presentation\/d\/([a-zA-Z0-9-_]+)/;
  var match = url.match(regex);

  Logger.log("////////////////");
  Logger.log("URL de plantillas slide: " + url);
  Logger.log("ID de plantillas slide: " + (match ? match[1] : "NO DETECTADO"));
  Logger.log("////////////////");

  return match ? match[1] : null;
}
function obtenerIdCarpeta(url) {
  // Extrae el ID de una URL de Google Drive Folder
  var regex = /\/folders\/([a-zA-Z0-9-_]+)/;
  var match = url.match(regex);

  Logger.log("////////////////");
  Logger.log("URL de carpeta: " + url);
  Logger.log("ID de carpeta: " + (match ? match[1] : "NO DETECTADO"));
  Logger.log("////////////////");

  return match ? match[1] : null;
}
