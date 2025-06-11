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

/**
 * Genera certificados PDF personalizados a partir de una plantilla de Google Slides.
 * Trabaja en lotes para no exceder límites de tiempo y guarda progreso para continuar después.
 * 
 * @param {string} sheet_Id    ID de la hoja de cálculo donde están los datos.
 * @param {string} template_Id ID de la plantilla de Google Slides.
 * @param {string} folder_Id   ID de la carpeta de Google Drive para guardar los certificados generados.
 * @param {number} batch_size  Número de certificados a generar por ejecución.
 */
function generarCertificados(sheet_Id, template_Id, folder_Id, batch_size) {
  try {
    const sheet = SpreadsheetApp.openById(sheet_Id).getSheetByName("data");
    const data = sheet.getDataRange().getValues();
    const folder = DriveApp.getFolderById(folder_Id);

    PropertiesService.getScriptProperties().setProperty("totalCertificados", data.length - 1);

    let lastProcessedIndex = PropertiesService.getScriptProperties().getProperty("lastProcessedIndex");
    let startIndex = lastProcessedIndex ? parseInt(lastProcessedIndex) + 1 : 1;
    let endIndex = Math.min(startIndex + batch_size, data.length);

    for (let i = startIndex; i < endIndex; i++) {
      const row = data[i];
      const nombreCompleto = `${row[1]} ${row[2] || ""} ${row[3]} ${row[4] || ""}`.trim();
      const documentoIdentidad = (row[5] ? row[5] + " " : "") + row[6];
      const codigoCertificado = row[12];
      const nombreCertificado = `Certificado_${codigoCertificado}_${row[1]}_${row[3]}.pdf`;

      if (folder.getFilesByName(nombreCertificado).hasNext()) {
        Logger.log("Certificado ya existe: " + nombreCertificado);
        continue;
      }

      const slideCopy = DriveApp.getFileById(template_Id).makeCopy("Temp_" + nombreCertificado, folder);
      const presentation = SlidesApp.openById(slideCopy.getId());
      const slide = presentation.getSlides()[0];

      slide.replaceAllText("{{nombre-participante}}", nombreCompleto);
      slide.replaceAllText("{{di-participante}}", documentoIdentidad);
      slide.replaceAllText("{{tipo-participante}}", row[9]);
      slide.replaceAllText("{{nombre-evento}}", row[10]);
      slide.replaceAllText("{{modalidad}}", row[13]);
      slide.replaceAllText("{{horas}}", row[14]);
      slide.replaceAllText("{{texto-fecha}}", row[17]);
      slide.replaceAllText("{{ubicacion}}", row[18]);
      slide.replaceAllText("{{cod-certificado}}", codigoCertificado);

      presentation.saveAndClose();

      const pdfUrl = "https://docs.google.com/presentation/d/" + slideCopy.getId() + "/export/pdf";
      const response = UrlFetchApp.fetch(pdfUrl, {
        headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
      });

      const pdfBlob = response.getBlob().setName(nombreCertificado);
      folder.createFile(pdfBlob);
      sheet.getRange(i + 1, 20).setValue(pdfUrl);

      DriveApp.getFileById(slideCopy.getId()).setTrashed(true);

      PropertiesService.getScriptProperties().setProperty("lastProcessedIndex", i);
    }

    // Aquí se decide si continuar o finalizar
    if (endIndex < data.length) {
      triggerGenerarCertificados(sheet_Id, template_Id, folder_Id, batch_size);
      return; // Salimos de la función para dejar que el trigger continúe
    } else {
      eliminarTrigger();
      PropertiesService.getScriptProperties().deleteProperty("lastProcessedIndex");
      PropertiesService.getScriptProperties().deleteProperty("totalCertificados");
    }
  } catch (e) {
    Logger.log("❌ Error en generarCertificados: " + e.toString());
  }
}


/**
 * Crea un trigger temporizado para continuar la generación de certificados en un lote posterior.
 * 
 * @param {string} sheet_Id
 * @param {string} template_Id
 * @param {string} folder_Id
 * @param {number} batch_size
 */
function triggerGenerarCertificados(sheet_Id, template_Id, folder_Id, batch_size) {
  // Elimina cualquier trigger anterior de esta función
  eliminarTrigger();  // Esto ya lo haces bien

  // Evita crear un nuevo trigger si ya hay uno agendado (doble verificación opcional)
  var triggers = ScriptApp.getProjectTriggers();
  var existe = triggers.some(t => t.getHandlerFunction() === "continuarGeneracionCertificados");
  if (existe) return;  // Ya hay uno activo, no creamos otro

  // Guarda parámetros
  PropertiesService.getScriptProperties().setProperty("sheet_Id", sheet_Id);
  PropertiesService.getScriptProperties().setProperty("template_Id", template_Id);
  PropertiesService.getScriptProperties().setProperty("folder_Id", folder_Id);
  PropertiesService.getScriptProperties().setProperty("batch_size", batch_size.toString());

  // Crea nuevo trigger para continuar después de 50 segundos
  ScriptApp.newTrigger("continuarGeneracionCertificados")
    .timeBased()
    .after(50000)
    .create();
}


/**
 * Función disparada por el trigger para continuar la generación de certificados.
 * Recupera los parámetros guardados y llama a la función principal.
 */
function continuarGeneracionCertificados() {
  var sheet_Id = PropertiesService.getScriptProperties().getProperty("sheet_Id");
  var template_Id = PropertiesService.getScriptProperties().getProperty("template_Id");
  var folder_Id = PropertiesService.getScriptProperties().getProperty("folder_Id");
  var batch_size = parseInt(PropertiesService.getScriptProperties().getProperty("batch_size"));

  if (sheet_Id && template_Id && folder_Id && batch_size) {
    generarCertificados(sheet_Id, template_Id, folder_Id, batch_size);
  }
}

/**
 * Elimina todos los triggers asociados a la función continuarGeneracionCertificados
 * para evitar múltiples ejecuciones paralelas o duplicadas.
 */
function eliminarTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "continuarGeneracionCertificados") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function obtenerProgresoCertificados() {
  var props = PropertiesService.getScriptProperties();
  var last = parseInt(props.getProperty("lastProcessedIndex")) || 0;
  var total = parseInt(props.getProperty("totalCertificados")) || 0;

  var porcentaje = total ? Math.floor((last / total) * 100) : 0;

  return {
    porcentaje: porcentaje,
    mensaje: "Generando certificados...",
    generados: last,
    total: total
  };
}
