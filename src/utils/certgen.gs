/**
    var sheet = SpreadsheetApp.openById("1s8jKetA0Raoob90lynpTyYdaEgYfmo2ioM4CLYSVbMc").getSheetByName("data");  
    var data = sheet.getDataRange().getValues(); // Obtener todos los datos de la hoja
    var templateId = "1avAc0K8N1dAjl2361-3MHkBbDbGzlKHsbJfOpR9RZNc"; // ID de la plantilla del certificado en Google Slides
    var folderId = "161P6RsKp-AyYN6Gh2853HvYtNJNQ4dZw"; // ID de la carpeta donde se guardarán los certificados
    var folder = DriveApp.getFolderById(folderId); // Obtener la carpeta en Google Drive
  */

/**
 * Función principal para generar certificados en lotes automáticamente.
 * @param {string} sheet_Id - ID de la hoja de cálculo donde están los datos de los participantes.
 * @param {string} template_Id - ID de la plantilla de certificado en Google Slides.
 * @param {string} folder_Id - ID de la carpeta donde se guardarán los certificados generados.
 * @param {int} batch_size - Tamaño del lote (por defecto, 30).
 */
function generarCertificados(sheet_Id, template_Id, folder_Id, batch_size) {
  try {
    var sheet = SpreadsheetApp.openById(sheet_Id).getSheetByName("data");
    var data = sheet.getDataRange().getValues();
    var folder = DriveApp.getFolderById(folder_Id);

    var lastProcessedIndex = PropertiesService.getScriptProperties().getProperty("lastProcessedIndex");
    var startIndex = lastProcessedIndex ? parseInt(lastProcessedIndex) + 1 : 1;
    var endIndex = Math.min(startIndex + batch_size, data.length);

    for (var i = startIndex; i < endIndex; i++) {
      // Extraer los valores según la estructura de la hoja de cálculo
      var primerNombre = data[i][1];  // pri_nom
      var segundoNombre = data[i][2]; // seg_nom
      var primerApellido = data[i][3]; // pri_ape
      var segundoApellido = data[i][4]; // seg_ape
      var prefijoDoc = data[i][5];  // prefijo_doc_i
      var docIdentidad = data[i][6]; // doc_i
      var tipoParticipante = data[i][9]; // tipo-participante
      var nombreEvento = data[i][10]; // nombre-evento
      var modalidad = data[i][13]; // modalidad
      var horas = data[i][14]; // dur-hr-ac
      var textoFecha = data[i][17]; // texto-fecha
      var ubicacion = data[i][18]; // ubicacion
      var codigoCertificado = data[i][12]; // cod-certificado

      var nombreCompleto = primerNombre + " " + (segundoNombre || "") + " " + primerApellido + " " + (segundoApellido || "");
      nombreCompleto = nombreCompleto.trim(); // Eliminar espacios extra

      var documentoIdentidad = (prefijoDoc ? prefijoDoc + " " : "") + docIdentidad;

     

      var nombreCertificado = "Certificado_" + codigoCertificado + "_" + primerNombre + "_" + primerApellido + ".pdf";

      // Verificar si el certificado ya existe
      var archivos = folder.getFilesByName(nombreCertificado);
      if (archivos.hasNext()) {
        Logger.log("El certificado " + nombreCertificado + " ya existe. Se omite.");
        continue;
      }

      // Verificar si el certificado ya existe
      var archivos = folder.getFilesByName(nombreCertificado);
      if (archivos.hasNext()) {
        Logger.log("El certificado " + nombreCertificado + " ya existe. Se omite.");
        continue;
      }

      var slideCopy = DriveApp.getFileById(template_Id).makeCopy("Temp_" + nombreCertificado, folder);
      var presentation = SlidesApp.openById(slideCopy.getId());
      var slide = presentation.getSlides()[0];

      // Reemplazar los campos en la plantilla
      slide.replaceAllText("{{nombre-participante}}", nombreCompleto);
      slide.replaceAllText("{{di-participante}}", documentoIdentidad);
      slide.replaceAllText("{{tipo-participante}}", tipoParticipante);
      slide.replaceAllText("{{nombre-evento}}", nombreEvento);
      slide.replaceAllText("{{modalidad}}", modalidad);
      slide.replaceAllText("{{horas}}", horas);
      slide.replaceAllText("{{texto-fecha}}", textoFecha);
      slide.replaceAllText("{{ubicacion}}", ubicacion);
      slide.replaceAllText("{{cod-certificado}}", codigoCertificado);

      presentation.saveAndClose();

      var pdfUrl = "https://docs.google.com/presentation/d/" + slideCopy.getId() + "/export/pdf";
      var response = UrlFetchApp.fetch(pdfUrl, {
        headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
      });

      var pdfBlob = response.getBlob().setName(nombreCertificado);
      folder.createFile(pdfBlob);
      sheet.getRange(i + 1, 20).setValue(pdfUrl);

      DriveApp.getFileById(slideCopy.getId()).setTrashed(true);
      PropertiesService.getScriptProperties().setProperty("lastProcessedIndex", i);
    }

    if (endIndex < data.length) {
      triggerGenerarCertificados(sheet_Id, template_Id, folder_Id, batch_size);
    } else {
      eliminarTrigger();
      PropertiesService.getScriptProperties().deleteProperty("lastProcessedIndex");
    }
  } catch (e) {
    Logger.log("Error en generarCertificados: " + e.toString());
  }
}


/**
 * Crea un trigger para ejecutar la función generarCertificados automáticamente después de 1 minuto.
 */
function triggerGenerarCertificados(sheet_Id, template_Id, folder_Id, batch_size) {
  eliminarTrigger();
  
  // Guardar los parámetros en PropertiesService
  PropertiesService.getScriptProperties().setProperty("sheet_Id", sheet_Id);
  PropertiesService.getScriptProperties().setProperty("template_Id", template_Id);
  PropertiesService.getScriptProperties().setProperty("folder_Id", folder_Id);
  PropertiesService.getScriptProperties().setProperty("batch_size", batch_size.toString());

  ScriptApp.newTrigger("continuarGeneracionCertificados")
    .timeBased()
    .after(60000)
    .create();
}

/**
 * Función auxiliar para continuar la generación en el siguiente lote.
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
 * Elimina cualquier trigger existente para evitar ejecuciones innecesarias.
 */
function eliminarTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "continuarGeneracionCertificados") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * Agrega un menú en la hoja de cálculo al abrir el archivo.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Certificados')
    .addItem('Generar Certificados', 'mostrarFormulario')
    .addToUi();
}



/** Envío de certificados via correo electrónico */

/**
 * Envía certificados por correo electrónico en lotes desde una hoja de cálculo de Google Sheets.
 * 
 * @param {string} sheet_Id - ID de la hoja de cálculo de Google Sheets que contiene los datos.
 * @param {string} folder_Id - ID de la carpeta de Google Drive donde están almacenados los certificados.
 * @param {number} batchSize - Cantidad de certificados a enviar por cada ejecución (para evitar el límite de tiempo de ejecución).
 */
function enviarCertificadosEmail(sheet_Id, folder_Id, batchSize) {
  try {
    // Abre la hoja de cálculo por su ID y selecciona la hoja llamada "data"
    var sheet = SpreadsheetApp.openById(sheet_Id).getSheetByName("data");
    
    // Obtiene todos los datos de la hoja en forma de matriz
    var data = sheet.getDataRange().getValues();
    
    // Abre la carpeta en Google Drive donde están almacenados los certificados
    var folder = DriveApp.getFolderById(folder_Id);
    
    // Recupera el índice de la última fila procesada para continuar desde allí
    var properties = PropertiesService.getScriptProperties();
    var startIndex = parseInt(properties.getProperty("lastProcessedIndex")) || 1; // Empieza en 1 para ignorar la fila de encabezados
    
    // Calcula el índice final del lote actual
    var endIndex = Math.min(startIndex + batchSize, data.length); 

    Logger.log("Procesando desde la fila " + startIndex + " hasta la fila " + (endIndex - 1));

    // Recorre los registros dentro del rango determinado por batchSize
    for (var i = startIndex; i < endIndex; i++) {
      // Extrae los datos de cada columna relevante
      var primerNombre = data[i][1];  // Columna B
      var segundoNombre = data[i][2]; // Columna C
      var primerApellido = data[i][3]; // Columna D
      var segundoApellido = data[i][4]; // Columna E
      var correo = data[i][7]; // Columna H
      var tituloEvento = data[i][10]; // Columna K
      var textoFecha = data[i][17]; // Columna R
      var codigoCertificado = data[i][12]; // Columna M

      // Construye el nombre completo del participante
      var nombreCompleto = primerNombre + " " + (segundoNombre || "") + " " + primerApellido + " " + (segundoApellido || "");
      nombreCompleto = nombreCompleto.trim(); // Elimina espacios adicionales en caso de que falte algún nombre o apellido

      // Construye el nombre del archivo del certificado según el formato especificado
      var nombreCertificado = "Certificado_" + codigoCertificado + "_" + primerNombre + "_" + primerApellido + ".pdf";

      // Busca el archivo del certificado en la carpeta de Drive
      var archivos = folder.getFilesByName(nombreCertificado);
      
      if (archivos.hasNext()) {
        // Si el archivo existe, lo obtiene
        var archivo = archivos.next();
        
        // Define el asunto del correo
        var asunto = "Certificado de participación - " + tituloEvento;
        
        // Define el cuerpo del correo con la información del evento y el certificado adjunto
        var cuerpo = "Estimado(a) " + nombreCompleto + ", reciba un cordial saludo en nombre de la Universidad Nacional Experimental del Táchira UNET\n\n" +
                     "Nos permitimos informarle que en el archivo adjunto podrá obtener el certificado digital correspondiente a su participación en el evento \n\"" + tituloEvento + "\", " + textoFecha + ".\n\n" +
                     "Es de destacar que el presente certificado ha sido firmado por las autoridades de nuestra institución universitaria; lo que lo valida como un documento oficial y avala que usted ha recibido formación profesional a través de esta casa de estudios; de haber cancelado el aporte para el certificado físico, a la brevedad se le indicará para que proceda a retirarlo.\n\n" +
                     "Le invitamos a seguir participando en futuros eventos.\n\n" +
                     "Atentamente,\n\nEquipo de Soporte Tecnológico\nDecanato de Investigación\nUNET\n\n" +
                     "Nota: Ante cualquier duda o aclaratoria relacionada con su certificado, escribir al correo electrónico: soporteinv@unet.edu.ve; sugiriendo en tal caso que sea como respuesta a este correo electrónico.\n";

        // Envía el correo con el certificado adjunto
        MailApp.sendEmail({
          to: correo,
          subject: asunto,
          body: cuerpo,
          attachments: [archivo.getAs(MimeType.PDF)]
        });

        Logger.log("Certificado enviado a: " + correo);
      } else {
        Logger.log("No se encontró el certificado para: " + nombreCompleto);
      }
    }

    // Verifica si quedan registros pendientes para continuar en la próxima ejecución
    if (endIndex < data.length) {
      properties.setProperty("lastProcessedIndex", endIndex);
      Logger.log("Proceso pausado. Continuará desde la fila: " + endIndex);
    } else {
      // Si ya se procesaron todos los registros, se elimina la propiedad para empezar desde el inicio en la próxima ejecución
      properties.deleteProperty("lastProcessedIndex");
      Logger.log("Proceso finalizado. Todos los certificados fueron enviados.");
    }

  } catch (e) {
    Logger.log("Error en enviarCertificadosEmail: " + e.toString());
  }
}
/**
 * Ejecuta el proceso de envío de certificados desde el menú de la hoja de cálculo.
 * Solicita al usuario los datos necesarios para la ejecución del proceso:
 * - ID de la hoja de cálculo
 * - ID de la carpeta donde están almacenados los certificados
 * - Cantidad de correos a enviar por lote
 * Luego, llama a la función `enviarCertificadosEmail` con los parámetros proporcionados.
 */
function ejecutarEnvioDesdeMenu() {
  var ui = SpreadsheetApp.getUi(); // Obtiene la interfaz de usuario para mostrar cuadros de diálogo

  // Solicita el ID de la hoja de cálculo
  var sheetResponse = ui.prompt("Ingrese el ID de la hoja de cálculo:");
  if (sheetResponse.getSelectedButton() !== ui.Button.OK) return; // Si el usuario cancela, se detiene la ejecución
  var sheet_Id = sheetResponse.getResponseText().trim(); // Obtiene y limpia el ID de la hoja de cálculo

  // Solicita el ID de la carpeta de certificados en Google Drive
  var folderResponse = ui.prompt("Ingrese el ID de la carpeta de certificados:");
  if (folderResponse.getSelectedButton() !== ui.Button.OK) return; // Si el usuario cancela, se detiene la ejecución
  var folder_Id = folderResponse.getResponseText().trim(); // Obtiene y limpia el ID de la carpeta

  // Solicita la cantidad de certificados a enviar por lote
  var batchResponse = ui.prompt("Ingrese la cantidad de envíos por lote (ejemplo: 50):");
  if (batchResponse.getSelectedButton() !== ui.Button.OK) return; // Si el usuario cancela, se detiene la ejecución
  var batchSize = parseInt(batchResponse.getResponseText().trim()) || 50; // Convierte el valor a número, usa 50 por defecto si hay error

  // Llama a la función de envío con los parámetros proporcionados
  enviarCertificadosEmail(sheet_Id, folder_Id, batchSize);
}
/**
 * Muestra el formulario FormularioParam.html en la barra lateral de Google Sheets.
 */
function mostrarFormulario() {
  var html = HtmlService.createHtmlOutputFromFile("FormularioParam")
    .setTitle("Generar Certificados")
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
  }
/**
 * Agrega un menú personalizado en la hoja de cálculo cuando se abre el archivo.
 * El menú "Certificados" contiene dos opciones:
 * - "Generar Certificados" (llama a `mostrarFormulario`)
 * - "Enviar Certificados" (llama a `ejecutarEnvioDesdeMenu`)
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi(); // Obtiene la interfaz de usuario de la hoja de cálculo

  // Crea un menú personalizado en la barra superior de la hoja de cálculo
  ui.createMenu("Certificados")
    .addItem("Generar Certificados", "mostrarFormulario") // Opción para generar certificados
    .addItem("Enviar Certificados", "ejecutarEnvioDesdeMenu") // Opción para enviar certificados por correo
    .addToUi(); // Agrega el menú a la interfaz
}
