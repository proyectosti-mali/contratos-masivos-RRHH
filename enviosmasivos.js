// Versión 2.0 Para Appscript de Google
// Envío masivo de contratos/prácticas por correo con adjunto PDF desde Drive
//Lee el documento sin perder ceros (ej. '07654321 → 07654321) y sirve para 8–11 dígitos.
//Indexa archivos con regex (\d{8,11}) (DNI/CE).
//Cambia subject + htmlBody según TIPO DE CONTRATO

function envioMasivoContratos() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hoja 1");
  const datos = hoja.getDataRange().getValues();

  const carpetaId = "10ycIn7wyKTFBhPgYbj-VQvjIdwJ2yJ19";
  const carpeta = DriveApp.getFolderById(carpetaId);
  const archivos = carpeta.getFiles();

  // --- Helpers ---
  const normalizarDocumento = (valor) => {
    if (valor === null || valor === undefined) return "";
    // Mantener como string, quitar apostrofe si existiera, quitar espacios y todo lo no numérico
    const s = String(valor).trim().replace(/^'/, "");
    const soloDigitos = s.replace(/\D/g, ""); // conserva ceros a la izquierda
    // A prueba de CE 9-11 y DNI 8 (en general 8-11)
    if (soloDigitos.length < 8 || soloDigitos.length > 11) return "";
    return soloDigitos;
  };

  const limpiarTexto = (valor) => (valor === null || valor === undefined) ? "" : String(valor).trim();

  // Quita tildes/acentos y caracteres raros para nombre de archivo adjunto (opcional pero ayuda)
  const slugNombre = (texto) => {
    return limpiarTexto(texto)
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // quita acentos
      .replace(/[^\w\s.-]/g, "") // deja letras, números, _, espacios, . y -
      .replace(/\s+/g, " ") // colapsa espacios
      .trim();
  };

  const obtenerPlantillaCorreo = (tipoContrato, dni, nombre) => {
    const tipo = limpiarTexto(tipoContrato).toLowerCase();

    if (tipo === "prácticas" || tipo === "practicas") {
      return {
        subject: `Convenio ${tipoContrato} - ${dni} - ${nombre}`,
        htmlBody: `
          Estimado(a) <b>${nombre}</b>,<br><br>
          Cumplimos con enviar su convenio debidamente firmado por ambas partes.<br><br>
          Saludos,<br><br>
          <strong>Talento Humano</strong><br>
          <strong>ASOCIACIÓN MUSEO DE ARTE DE LIMA</strong>
        `
      };
    }

    if (tipo === "laboral") {
      return {
        subject: `Contrato ${tipoContrato} - ${dni} - ${nombre}`,
        htmlBody: `
          Estimado(a) <b>${nombre}</b>,<br><br>
          Cumplimos con enviar su contrato debidamente firmado por ambas partes.<br><br>
          Saludos,<br><br>
          <strong>Talento Humano</strong><br>
          <strong>ASOCIACIÓN MUSEO DE ARTE DE LIMA</strong>
        `
      };
    }

    // Default: Locación (y cualquier otro valor cae aquí)
    return {
      subject: `Contrato ${tipoContrato} - ${dni} - ${nombre}`,
      htmlBody: `
        Estimado(a) <b>${nombre}</b>,<br><br>
        Cumplimos con enviar su contrato bajo la modalidad <b>Locación de Servicio</b>
        debidamente firmado por ambas partes.<br><br>
        Saludos,<br><br>
        <strong>Talento Humano</strong><br>
        <strong>ASOCIACIÓN MUSEO DE ARTE DE LIMA</strong>
      `
    };
  };

  // --- Indexar PDFs por documento (DNI/CE) ---
  // Guarda lista por doc por si hubiera más de 1 archivo con el mismo doc
  const pdfPorDoc = {}; // { "87654321": [file1, file2], ... }

  while (archivos.hasNext()) {
    const archivo = archivos.next();
    const nombreArchivo = archivo.getName();

    // Busca cualquier secuencia de 8 a 11 dígitos en el nombre (DNI/CE)
    const match = nombreArchivo.match(/(\d{8,11})/);
    if (match) {
      const doc = match[1];
      if (!pdfPorDoc[doc]) pdfPorDoc[doc] = [];
      pdfPorDoc[doc].push(archivo);
    }
  }

  // --- Recorre filas ---
  // Columnas (según tu tabla):
  // 0 DNI/CE | 1 NOMBRE | 2 TIPO | 3 correos | 4 ESTADO
  for (let i = 1; i < datos.length; i++) {
    const doc = normalizarDocumento(datos[i][0]); // DNI/CE
    const nombre = limpiarTexto(datos[i][1]);
    const tipoContrato = limpiarTexto(datos[i][2]);
    const correo = limpiarTexto(datos[i][3]);
    const estado = limpiarTexto(datos[i][4]).toUpperCase();

    // No reenviar
    if (estado === "ENVIADO") {
      Logger.log(`Fila ${i + 1} ya enviada, se omite`);
      continue;
    }

    // Validaciones
    if (!doc || !correo || !nombre || !tipoContrato) {
      hoja.getRange(i + 1, 5).setValue("ERROR");
      Logger.log(`ERROR fila ${i + 1}: faltan datos (doc/correo/nombre/tipo)`);
      continue;
    }

    const candidatos = pdfPorDoc[doc];
    if (!candidatos || candidatos.length === 0) {
      hoja.getRange(i + 1, 5).setValue("ERROR");
      Logger.log(`PDF no encontrado para documento ${doc} (fila ${i + 1})`);
      continue;
    }

    // Si hay varios con el mismo doc, toma el primero (o podrías mejorar con más filtros)
    const archivo = candidatos[0];

    try {
      const plantilla = obtenerPlantillaCorreo(tipoContrato, doc, nombre);

      // Nombre del adjunto: conserva espacios en el nombre (como tu ejemplo)
      const nombreAdjunto = `${(plantilla.subject.toLowerCase().startsWith("convenio") ? "Convenio" : "Contrato")} ${tipoContrato}_${doc}_${slugNombre(nombre)}.pdf`;
      const pdfAdjunto = archivo.getBlob().setName(nombreAdjunto);

      MailApp.sendEmail({
        to: correo,
        subject: plantilla.subject,
        htmlBody: plantilla.htmlBody,
        attachments: [pdfAdjunto]
      });

      hoja.getRange(i + 1, 5).setValue("ENVIADO");
      Logger.log(`ENVIADO: ${correo} (fila ${i + 1})`);

    } catch (error) {
      hoja.getRange(i + 1, 5).setValue("ERROR");
      Logger.log(`ERROR fila ${i + 1}: ${error && error.message ? error.message : error}`);
    }
  }
}

