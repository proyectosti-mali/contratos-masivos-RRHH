// Versión 1.0 - Para Appscript de Google
// Envío masivo de contratos en PDF desde Google Drive
// Basado en una hoja de cálculo con columnas: DNI | Nombre | TipoContrato | Correo
function envioMasivoContratos() {

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hoja 1");
  const datos = hoja.getDataRange().getValues();

  const carpetaId = "13CrgNA0JIH1iEoEPwreRPlPuydPQDX5U";
  const carpeta = DriveApp.getFolderById(carpetaId);
  const archivos = carpeta.getFiles();

  // Indexa por DNI encontrado dentro del nombre del archivo
  let pdfPorDni = {};

  while (archivos.hasNext()) {
    let archivo = archivos.next();
    let name = archivo.getName().toLowerCase();

    // busca un bloque de 8 dígitos (DNI) dentro del nombre
    let match = name.match(/\b\d{8}\b/);
    if (match) {
      pdfPorDni[match[0]] = archivo;
    }
  }

  for (let i = 1; i < datos.length; i++) {
    let dni = String(datos[i][0]).split(".")[0].trim();
    let nombre = String(datos[i][1]).trim();
    let tipocontrato = String(datos[i][2]).trim();
    let correo = datos[i][3];

    if (!correo || !dni) continue;

    Logger.log("Fila " + (i+1) + " DNI: " + dni);

    let archivo = pdfPorDni[dni];

    if (!archivo) {
      Logger.log("NO ENCONTRADO PDF para DNI: " + dni);
      continue;
    }

    MailApp.sendEmail({
      to: correo,
      subject: `Envío de contrato por ${tipocontrato} - ${dni} - ${nombre}`,
      htmlBody: `
        Hola <b>${nombre}</b>,<br><br>
        Adjuntamos tu contrato bajo la modalidad <b>${tipocontrato}</b>.<br><br>
        Saludos,<br>
        RRHH – Mali
      `,
      attachments: [archivo.getBlob()]
    });

    Logger.log("ENVIADO a: " + correo + " con PDF: " + archivo.getName());
  }
}
