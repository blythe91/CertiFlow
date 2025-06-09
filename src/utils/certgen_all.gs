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
    // Abre la hoja de cálculo y obtiene todos los datos de la hoja "data"
    var sheet = SpreadsheetApp.openById(sheet_Id).getSheetByName("data");
    var data = sheet.getDataRange().getValues();

    PropertiesService.getScriptProperties().setProperty("totalCertificados", data.length - 1); // excluye encabezado

    
    // Obtiene la carpeta de Drive donde se guardarán los certificados PDF
    var folder = DriveApp.getFolderById(folder_Id);

    // Recupera el índice del último certificado procesado para continuar donde quedó
    var lastProcessedIndex = PropertiesService.getScriptProperties().getProperty("lastProcessedIndex");
    var startIndex = lastProcessedIndex ? parseInt(lastProcessedIndex) + 1 : 1; // Si no hay registro, empieza en 1
    var endIndex = Math.min(startIndex + batch_size, data.length); // Límite máximo para batch

    // Itera sobre el rango de participantes a procesar en este lote
    for (var i = startIndex; i < endIndex; i++) {
      // Extrae los datos relevantes de cada fila (participante)
      var primerNombre = data[i][1];
      var segundoNombre = data[i][2];
      var primerApellido = data[i][3];
      var segundoApellido = data[i][4];
      var prefijoDoc = data[i][5];
      var docIdentidad = data[i][6];
      var tipoParticipante = data[i][9];
      var nombreEvento = data[i][10];
      var modalidad = data[i][13];
      var horas = data[i][14];
      var textoFecha = data[i][17];
      var ubicacion = data[i][18];
      var codigoCertificado = data[i][12];

      // Construye el nombre completo y el documento con prefijo si existe
      var nombreCompleto = primerNombre + " " + (segundoNombre || "") + " " + primerApellido + " " + (segundoApellido || "");
      nombreCompleto = nombreCompleto.trim();
      var documentoIdentidad = (prefijoDoc ? prefijoDoc + " " : "") + docIdentidad;

      // Nombre que tendrá el archivo PDF del certificado
      var nombreCertificado = "Certificado_" + codigoCertificado + "_" + primerNombre + "_" + primerApellido + ".pdf";

      // Verifica si el certificado ya existe para evitar duplicados
      var archivos = folder.getFilesByName(nombreCertificado);
      if (archivos.hasNext()) {
        Logger.log("El certificado " + nombreCertificado + " ya existe. Se omite.");
        continue; // Salta al siguiente participante
      }

      // Crea una copia temporal de la plantilla para personalizar
      var slideCopy = DriveApp.getFileById(template_Id).makeCopy("Temp_" + nombreCertificado, folder);
      var presentation = SlidesApp.openById(slideCopy.getId());
      var slide = presentation.getSlides()[0];

      // Reemplaza los marcadores de posición en la diapositiva con los datos personalizados
      slide.replaceAllText("{{nombre-participante}}", nombreCompleto);
      slide.replaceAllText("{{di-participante}}", documentoIdentidad);
      slide.replaceAllText("{{tipo-participante}}", tipoParticipante);
      slide.replaceAllText("{{nombre-evento}}", nombreEvento);
      slide.replaceAllText("{{modalidad}}", modalidad);
      slide.replaceAllText("{{horas}}", horas);
      slide.replaceAllText("{{texto-fecha}}", textoFecha);
      slide.replaceAllText("{{ubicacion}}", ubicacion);
      slide.replaceAllText("{{cod-certificado}}", codigoCertificado);

      // Guarda y cierra la presentación modificada
      presentation.saveAndClose();

      // Construye la URL para exportar la presentación como PDF
      var pdfUrl = "https://docs.google.com/presentation/d/" + slideCopy.getId() + "/export/pdf";

      // Realiza la petición para obtener el PDF con autorización OAuth del script
      var response = UrlFetchApp.fetch(pdfUrl, {
        headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
      });

      // Obtiene el PDF como Blob y lo guarda en la carpeta
      var pdfBlob = response.getBlob().setName(nombreCertificado);
      folder.createFile(pdfBlob);

      // Guarda la URL del PDF en la columna 20 (T) de la hoja para referencia futura
      sheet.getRange(i + 1, 20).setValue(pdfUrl);

      // Mueve a la papelera el archivo temporal de Slides para no saturar el Drive
      DriveApp.getFileById(slideCopy.getId()).setTrashed(true);

      // Actualiza la propiedad con el índice del último certificado procesado
      PropertiesService.getScriptProperties().setProperty("lastProcessedIndex", i);
    }

    // Si quedan participantes pendientes, programa un trigger para continuar el proceso
    if (endIndex < data.length) {
      triggerGenerarCertificados(sheet_Id, template_Id, folder_Id, batch_size);
    } else {
      // Si ya terminó, elimina triggers pendientes y limpia propiedad de progreso
      eliminarTrigger();
      PropertiesService.getScriptProperties().deleteProperty("lastProcessedIndex");
      PropertiesService.getScriptProperties().deleteProperty("totalCertificados");
    }
  } catch (e) {
    // Loguea cualquier error para facilitar depuración
    Logger.log("Error en generarCertificados: " + e.toString());
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
