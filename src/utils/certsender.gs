/**
 * Envía certificados por correo electrónico en lotes desde una hoja de cálculo de Google Sheets.
 * 
 * @param {string} sheet_Id - ID de la hoja de cálculo de Google Sheets que contiene los datos.
 * @param {string} folder_Id - ID de la carpeta de Google Drive donde están almacenados los certificados.
 * @param {number} batchSize - Cantidad de certificados a enviar por cada ejecución (para evitar el límite de tiempo de ejecución).
 */
function enviarCertificadosEmail(sheet_Id, folder_Id, batchSize) {
  try {
    var sheet = SpreadsheetApp.openById(sheet_Id).getSheetByName("data");
    var data = sheet.getDataRange().getValues();
    var folder = DriveApp.getFolderById(folder_Id);
    
    var properties = PropertiesService.getScriptProperties();
    var startIndex = parseInt(properties.getProperty("lastProcessedIndex")) || 1;
    var endIndex = Math.min(startIndex + batchSize, data.length); 

    Logger.log("Procesando desde la fila " + startIndex + " hasta la fila " + (endIndex - 1));

    for (var i = startIndex; i < endIndex; i++) {
      var primerNombre = data[i][1];
      var segundoNombre = data[i][2];
      var primerApellido = data[i][3];
      var segundoApellido = data[i][4];
      var correo = data[i][7];
      var tituloEvento = data[i][10];
      var textoFecha = data[i][17];
      var codigoCertificado = data[i][12];

      var nombreCompleto = primerNombre + " " + (segundoNombre || "") + " " + primerApellido + " " + (segundoApellido || "");
      nombreCompleto = nombreCompleto.trim();

      var nombreCertificado = "Certificado_" + codigoCertificado + "_" + primerNombre + "_" + primerApellido + ".pdf";
      var archivos = folder.getFilesByName(nombreCertificado);
      
      if (archivos.hasNext()) {
        var archivo = archivos.next();
        
        var asunto = "Certificado de participación - " + tituloEvento;
        var cuerpo = "Estimado(a) " + nombreCompleto + ", reciba un cordial saludo en nombre de la Universidad Nacional Experimental del Táchira UNET\n\n" +
                     "Nos permitimos informarle que en el archivo adjunto podrá obtener el certificado digital correspondiente a su participación en el evento \n\"" + tituloEvento + "\", " + textoFecha + ".\n\n" +
                     "Es de destacar que el presente certificado ha sido firmado por las autoridades de nuestra institución universitaria; lo que lo valida como un documento oficial y avala que usted ha recibido formación profesional a través de esta casa de estudios; de haber cancelado el aporte para el certificado físico, a la brevedad se le indicará para que proceda a retirarlo.\n\n" +
                     "Le invitamos a seguir participando en futuros eventos.\n\n" +
                     "Atentamente,\n\nEquipo de Soporte Tecnológico\nDecanato de Investigación\nUNET\n\n" +
                     "Nota: Ante cualquier duda o aclaratoria relacionada con su certificado, escribir al correo electrónico: soporteinv@unet.edu.ve; sugiriendo en tal caso que sea como respuesta a este correo electrónico.\n";

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

    if (endIndex < data.length) {
      properties.setProperty("lastProcessedIndex", endIndex);
      Logger.log("Proceso pausado. Continuará desde la fila: " + endIndex);
    } else {
      properties.deleteProperty("lastProcessedIndex");
      Logger.log("Proceso finalizado. Todos los certificados fueron enviados.");
    }

  } catch (e) {
    Logger.log("Error en enviarCertificadosEmail: " + e.toString());
  }
}

/**
 * Ejecuta el proceso de envío de certificados desde el menú de la hoja de cálculo.
 * Solicita al usuario los datos necesarios para la ejecución del proceso.
 */
function ejecutarEnvioDesdeMenu() {
  var ui = SpreadsheetApp.getUi();

  var sheetResponse = ui.prompt("Ingrese el ID de la hoja de cálculo:");
  if (sheetResponse.getSelectedButton() !== ui.Button.OK) return;
  var sheet_Id = sheetResponse.getResponseText().trim();

  var folderResponse = ui.prompt("Ingrese el ID de la carpeta de certificados:");
  if (folderResponse.getSelectedButton() !== ui.Button.OK) return;
  var folder_Id = folderResponse.getResponseText().trim();

  var batchResponse = ui.prompt("Ingrese la cantidad de envíos por lote (ejemplo: 50):");
  if (batchResponse.getSelectedButton() !== ui.Button.OK) return;
  var batchSize = parseInt(batchResponse.getResponseText().trim()) || 50;

  enviarCertificadosEmail(sheet_Id, folder_Id, batchSize);
}

/**
 * Agrega un menú personalizado en la hoja de cálculo al abrir el archivo,
 * específicamente con la opción de enviar certificados.

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Certificados")
    .addItem("Enviar Certificados", "ejecutarEnvioDesdeMenu")
    .addToUi();
}
 */
