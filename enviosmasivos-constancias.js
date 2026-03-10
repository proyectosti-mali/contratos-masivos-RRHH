function envioMasivoConstancias() {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hoja 1");
    const datos = hoja.getDataRange().getValues();
  
    const carpetaId = "1WwKpt62gNpVmR67oXglM-EujvwxcPaPR";
    const carpeta = DriveApp.getFolderById(carpetaId);
    const archivos = carpeta.getFiles();
  
    // --- Helpers ---
    const limpiarTexto = (valor) => {
      return valor === null || valor === undefined ? "" : String(valor).trim();
    };
  
    const normalizarNumeroOrden = (valor) => {
      if (valor === null || valor === undefined) return "";
      const s = String(valor).trim().replace(/^'/, "");
      const soloDigitos = s.replace(/\D/g, "");
      return soloDigitos;
    };
  
    const slugNombre = (texto) => {
      return limpiarTexto(texto)
        .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
        .replace(/[^\w\s.-]/g, "")
        .replace(/\s+/g, " ")
        .trim();
    };
  
    const obtenerPlantillaCorreo = (tipoContrato, nombre) => {
      const tipo = limpiarTexto(tipoContrato).toLowerCase();
  
      if (tipo === "locación" || tipo === "locacion") {
        return {
          subject: `Constancia de ${tipoContrato} - ${nombre}`,
          htmlBody: `
            Estimado(a) <b>${nombre}</b>,<br><br>
            Cumplimos con enviar su constancia de locación de servicios.<br><br>
            Saludos,<br><br>
            <strong>Talento Humano</strong><br>
            <strong>ASOCIACIÓN MUSEO DE ARTE DE LIMA</strong>
          `
        };
      }
  
      if (tipo === "prácticas" || tipo === "practicas") {
        return {
          subject: `Constancia de ${tipoContrato} - ${nombre}`,
          htmlBody: `
            Estimado(a) <b>${nombre}</b>,<br><br>
            Cumplimos con enviar su constancia de prácticas.<br><br>
            Saludos,<br><br>
            <strong>Talento Humano</strong><br>
            <strong>ASOCIACIÓN MUSEO DE ARTE DE LIMA</strong>
          `
        };
      }
  
      if (tipo === "laboral") {
        return {
          subject: `Constancia ${tipoContrato} - ${nombre}`,
          htmlBody: `
            Estimado(a) <b>${nombre}</b>,<br><br>
            Cumplimos con enviar su constancia laboral.<br><br>
            Saludos,<br><br>
            <strong>Talento Humano</strong><br>
            <strong>ASOCIACIÓN MUSEO DE ARTE DE LIMA</strong>
          `
        };
      }
  
      return {
        subject: `Constancia - ${nombre}`,
        htmlBody: `
          Estimado(a) <b>${nombre}</b>,<br><br>
          Cumplimos con enviar su constancia correspondiente.<br><br>
          Saludos,<br><br>
          <strong>Talento Humano</strong><br>
          <strong>ASOCIACIÓN MUSEO DE ARTE DE LIMA</strong>
        `
      };
    };
  
    // --- Indexar PDFs por número de orden al final del nombre ---
    // Ejemplo: "Constancia Juan Perez-12.pdf" => orden 12
    const pdfPorNumero = {}; // { "1": file, "2": file, ... }
  
    while (archivos.hasNext()) {
      const archivo = archivos.next();
      const nombreArchivo = archivo.getName();
  
      // Busca número antes de ".pdf" y después de un guion al final
      const match = nombreArchivo.match(/-(\d+)\.pdf$/i);
      if (match) {
        const numeroOrden = match[1];
        pdfPorNumero[numeroOrden] = archivo;
      }
    }
  
    // --- Recorre filas ---
    // Ajusta estos índices según tu hoja:
    // 0 DNI/CE | 1 NOMBRE | 2 TIPO | 3 CORREO | 4 ESTADO | 5 N°
    for (let i = 1; i < datos.length; i++) {
      const nombre = limpiarTexto(datos[i][1]);
      const tipoContrato = limpiarTexto(datos[i][2]);
      const correo = limpiarTexto(datos[i][3]);
      const estado = limpiarTexto(datos[i][4]).toUpperCase();
      const numeroOrden = normalizarNumeroOrden(datos[i][5]); // columna N°
  
      if (estado === "ENVIADO") {
        Logger.log(`Fila ${i + 1} ya enviada, se omite`);
        continue;
      }
  
      if (!numeroOrden || !correo || !nombre || !tipoContrato) {
        hoja.getRange(i + 1, 5).setValue("ERROR");
        Logger.log(`ERROR fila ${i + 1}: faltan datos (N°/correo/nombre/tipo)`);
        continue;
      }
  
      const archivo = pdfPorNumero[numeroOrden];
  
      if (!archivo) {
        hoja.getRange(i + 1, 5).setValue("ERROR");
        Logger.log(`PDF no encontrado para N° ${numeroOrden} (fila ${i + 1})`);
        continue;
      }
  
      try {
        const plantilla = obtenerPlantillaCorreo(tipoContrato, nombre);
  
        const nombreAdjunto = `Constancia_${numeroOrden}_${slugNombre(nombre)}.pdf`;
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