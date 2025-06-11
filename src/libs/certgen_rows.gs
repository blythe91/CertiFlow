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

function generarCertificadosPorFilas(filasCSV, sheet_Id, template_Id, folder_Id) {
  try {
    var sheet = SpreadsheetApp.openById(sheet_Id).getSheetByName("data");
    var data = sheet.getDataRange().getValues();
    var folder = DriveApp.getFolderById(folder_Id);

    // Convierte "4,7,25" → [4, 7, 25]
    var filas = filasCSV.split(",").map(f => parseInt(f.trim()));

    
    filas.forEach(fila => {
      var i = fila; // Ajustar índice (fila 2 es índice 1)

      if (i <= 0 || i >= data.length) {
        Logger.log("Fila " + fila + " fuera de rango. Se omite.");
        return;
      }

      var filaData = data[i];

      // Validación: si no tiene nombre o cédula, se omite
      if (!filaData[1] || !filaData[3] || !filaData[6]) {
        Logger.log("Fila " + fila + " con datos incompletos. Se omite.");
        return;
      }

      var primerNombre = filaData[1];
      var segundoNombre = filaData[2];
      var primerApellido = filaData[3];
      var segundoApellido = filaData[4];
      var prefijoDoc = filaData[5];
      var docIdentidad = filaData[6];
      var tipoParticipante = filaData[9];
      var nombreEvento = filaData[10];
      var modalidad = filaData[13];
      var horas = filaData[14];
      var codigoCertificado = filaData[12];
      var textoFecha = filaData[17];
      var ubicacion = filaData[18];

      var nombreCompleto = primerNombre + " " + (segundoNombre || "") + " " + primerApellido + " " + (segundoApellido || "");
      nombreCompleto = nombreCompleto.trim();
      var documentoIdentidad = (prefijoDoc ? prefijoDoc + " " : "") + docIdentidad;
      var nombreCertificado = "Certificado_" + codigoCertificado + "_" + primerNombre + "_" + primerApellido + ".pdf";

      var archivos = folder.getFilesByName(nombreCertificado);
      if (archivos.hasNext()) {
        Logger.log("Certificado ya existe para la fila " + fila + ". Se omite.");
        return;
      }

      var slideCopy = DriveApp.getFileById(template_Id).makeCopy("Temp_" + nombreCertificado, folder);
      var presentation = SlidesApp.openById(slideCopy.getId());
      var slide = presentation.getSlides()[0];

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

      // Guarda la URL del PDF en la columna 20 (columna T)
      sheet.getRange(i + 1, 20).setValue(pdfUrl);

      // Elimina la copia temporal
      DriveApp.getFileById(slideCopy.getId()).setTrashed(true);

      Logger.log("Certificado generado para la fila " + fila);
    });

  } catch (e) {
    Logger.log("Error en generarCertificadosPorFilas: " + e.toString());
  }
}
