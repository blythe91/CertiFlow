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
 * Envía certificados por correo para un rango de filas dado.
 *
 * @param {number} startRow     Fila inicial (1-based).
 * @param {number} endRow       Fila final (1-based).
 * @param {string} sheetId      ID de la hoja de cálculo.
 * @param {string} folderId     ID de la carpeta de certificados PDF.
 * @param {string} mensajeEmail Texto adicional a incluir en el cuerpo del correo.
 */
function enviarCertificadosEmailPorRango(startRow, endRow, sheetId, folderId, mensajeEmail) {
  try {
    Logger.log("📄 Rango: " + startRow + " a " + endRow);
    if (endRow - startRow + 1 < 5 || endRow - startRow + 1 > 30) {
      Logger.log("❌ El rango debe tener entre 5 y 30 filas.");
      return;
    }

    const sheet = SpreadsheetApp.openById(sheetId).getSheetByName("data");
    const data = sheet.getDataRange().getValues();
    const folder = DriveApp.getFolderById(folderId);

    for (let i = startRow; i <= endRow; i++) {
      if (i <= 0 || i >= data.length) {
        Logger.log("⚠️ Índice de fila inválido: " + (i + 1));
        continue;
      }

      const fila = data[i];
      const primerNombre = fila[1];
      const segundoNombre = fila[2];
      const primerApellido = fila[3];
      const segundoApellido = fila[4];
      const correo = fila[7];
      const tituloEvento = fila[10];
      const textoFecha = fila[17];
      const codigoCertificado = fila[12];

      const nombreCompleto = (primerNombre + " " + (segundoNombre || "") + " " + primerApellido + " " + (segundoApellido || "")).trim();
      const nombreCertificado = `Certificado_${codigoCertificado}_${primerNombre}_${primerApellido}.pdf`;

      const archivos = folder.getFilesByName(nombreCertificado);
      if (archivos.hasNext()) {
        const archivo = archivos.next();

        const asunto = `Certificado de participación - ${codigoCertificado} - ${tituloEvento}`;
        const cuerpo = `Estimado(a) ${nombreCompleto}, reciba un cordial saludo en nombre de la Universidad Nacional Experimental del Táchira UNET\n\n` +
          `Nos permitimos informarle que en el archivo adjunto podrá obtener el certificado digital correspondiente a su participación en el evento \n\"${tituloEvento}\", ${textoFecha}.\n\n` +
          `Es de destacar que el presente certificado ha sido firmado por las autoridades de nuestra institución universitaria; lo que lo valida como un documento oficial y avala que usted ha recibido formación profesional a través de esta casa de estudios. ${mensajeEmail}\n\n` +
          `Le invitamos a seguir participando en futuros eventos.\n\n` +
          `Atentamente,\n\nEquipo de Soporte Tecnológico\nDecanato de Investigación\nUNET\n\n` +
          `Nota: Ante cualquier duda o aclaratoria relacionada con su certificado, escribir al correo electrónico: soporteinv@unet.edu.ve; sugiriendo en tal caso que sea como respuesta a este correo electrónico.\n`;

        MailApp.sendEmail({
          to: correo,
          subject: asunto,
          body: cuerpo,
          attachments: [archivo.getAs(MimeType.PDF)]
        });

        Logger.log("📧 " + nombreCertificado + ".pdf - " + tituloEvento + " - enviado a: " + correo);
      } else {
        Logger.log("⚠️ No se encontró el certificado para: " + nombreCompleto);
      }
    }

    Logger.log("✅ Proceso completado para las filas del " + startRow + " al " + endRow);

  } catch (e) {
    Logger.log("❌ Error en enviarCertificadosEmailPorRango: " + e.toString());
  }
}
