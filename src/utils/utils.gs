const DEFAULT_BATCH_SIZE = 30;

function obtenerBatchSizePorDefecto() {
  return DEFAULT_BATCH_SIZE;
}

function obtenerIdHojaCalculo(url) {
  // Extrae el ID de una URL de Google Sheets
  var regex = /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/;
  var match = url.match(regex);
  return match ? match[1] : null;
}
function obtenerIdPlantillaSlide(url) {
  // Extrae el ID de una URL de Google Slides
  var regex = /\/presentation\/d\/([a-zA-Z0-9-_]+)/;
  var match = url.match(regex);
  return match ? match[1] : null;
}
function obtenerIdCarpeta(url) {
  // Extrae el ID de una URL de Google Drive Folder
  var regex = /\/folders\/([a-zA-Z0-9-_]+)/;
  var match = url.match(regex);
  return match ? match[1] : null;
}
