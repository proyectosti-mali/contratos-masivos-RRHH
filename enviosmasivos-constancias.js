/**
 * ============================================================
 * ENVÍO MASIVO DE CONSTANCIAS
 * Versión: 3.3 (PRODUCCIÓN)
 * ============================================================
 *
 * Busca PDFs por número de orden al final del nombre de archivo:
 * Ejemplo:
 *   "Constancia Juan Perez-1.pdf"   -> N° = 1
 *   "Archivo_final-25.PDF"          -> N° = 25
 *
 * Hace match contra la columna N° del spreadsheet.
 */
function envioMasivoConstancias_v3_3() {

  /* ===================== LOCK ===================== */
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error("Otro proceso está en ejecución");
  }

  try {

    /* ===================== CONFIG ===================== */

    const NOMBRE_HOJA = "Hoja 1";
    const HOJA_LOG = "LOG_ENVÍOS";
    const FOLDER_ID = "1WwKpt62gNpVmR67oXglM-EujvwxcPaPR";

    // Límite defensivo de envíos por ejecución (además de la cuota diaria)
    const MAX_ENVIOS_POR_EJECUCION = 400;

    // Columnas (0-based)
    // Ajusta si tu hoja real cambia
    const COL_NUMERO = 0;   // N°
    const COL_NOMBRE = 1;   // NOMBRE
    const COL_TIPO = 2;     // TIPO
    const COL_CORREO = 3;   // CORREO
    const COL_ESTADO = 4;   // ESTADO

    const ss = SpreadsheetApp.getActive();
    const hoja = ss.getSheetByName(NOMBRE_HOJA);
    if (!hoja) throw new Error("No se encontró la hoja principal");

    const logSheet = ss.getSheetByName(HOJA_LOG) || ss.insertSheet(HOJA_LOG);

    const datos = hoja.getDataRange().getValues();
    if (datos.length < 2) return;

    const nuevosEstados = [];

    /* ===================== HELPERS ===================== */

    const limpiar = v => v === null || v === undefined ? "" : String(v).trim();

    const normalizarNumeroOrden = v => {
      const s = limpiar(v).replace(/^'/, "").replace(/\D/g, "");
      if (!s) return "";
      const n = parseInt(s, 10);
      return Number.isNaN(n) ? "" : String(n);
    };

    const emailValido = e => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(limpiar(e));

    const slugNombre = texto => {
      return limpiar(texto)
        .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
        .replace(/[^\w\s.-]/g, "")
        .replace(/\s+/g, " ")
        .trim();
    };

    const tipoValido = t =>
      ["laboral", "prácticas", "practicas", "locación", "locacion"]
        .includes(limpiar(t).toLowerCase());

    const plantillaCorreo = (tipoConstancia, nombre) => {
      const tipo = limpiar(tipoConstancia).toLowerCase();

      if (tipo === "prácticas" || tipo === "practicas") {
        return {
          subject: `Constancia de ${tipoConstancia} - ${nombre}`,
          body: `
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
          subject: `Constancia laboral - ${nombre}`,
          body: `
            Estimado(a) <b>${nombre}</b>,<br><br>
            Cumplimos con enviar su constancia laboral.<br><br>
            Saludos,<br><br>
            <strong>Talento Humano</strong><br>
            <strong>ASOCIACIÓN MUSEO DE ARTE DE LIMA</strong>
          `
        };
      }

      // Default: Locación
      return {
        subject: `Constancia de ${tipoConstancia} - ${nombre}`,
        body: `
          Estimado(a) <b>${nombre}</b>,<br><br>
          Cumplimos con enviar su constancia de locación de servicios.<br><br>
          Saludos,<br><br>
          <strong>Talento Humano</strong><br>
          <strong>ASOCIACIÓN MUSEO DE ARTE DE LIMA</strong>
        `
      };
    };

    const log = (numero, accion, detalle) => {
      logSheet.appendRow([new Date(), numero, accion, detalle]);
    };

    /* ===================== CUOTA ===================== */

    let cuotaDisponible = MailApp.getRemainingDailyQuota();
    if (cuotaDisponible <= 0) {
      throw new Error("Cuota de correos agotada");
    }

    let enviadosEnEjecucion = 0;

    /* ===================== INDEXAR PDFs ===================== */

    // Clave: número de orden
    // Si hay más de uno, se guarda como array para detectar duplicados
    const index = {};

    // Toma el número final antes de .pdf
    // Ejemplos válidos:
    // "archivo-1.pdf"
    // "archivo final - 25.pdf"
    // "constancia_x-003.pdf"
    const regexPDF = /-\s*(\d+)\.pdf$/i;

    const files = DriveApp.getFolderById(FOLDER_ID).getFiles();
    while (files.hasNext()) {
      const f = files.next();
      const nombreArchivo = f.getName();
      const m = nombreArchivo.match(regexPDF);
      if (!m) continue;

      const clave = normalizarNumeroOrden(m[1]);
      if (!clave) continue;

      index[clave] = index[clave] || [];
      index[clave].push(f);
    }

    /* ===================== PROCESO ===================== */

    for (let i = 1; i < datos.length; i++) {
      const fila = i + 1;
      const estadoActual = limpiar(datos[i][COL_ESTADO]).toUpperCase();

      if (estadoActual === "ENVIADO") {
        nuevosEstados.push([estadoActual]);
        continue;
      }

      if (cuotaDisponible <= 0) {
        nuevosEstados.push(["PENDIENTE_CUOTA"]);
        continue;
      }

      if (enviadosEnEjecucion >= MAX_ENVIOS_POR_EJECUCION) {
        nuevosEstados.push(["PENDIENTE_BATCH"]);
        continue;
      }

      const numeroOrden = normalizarNumeroOrden(datos[i][COL_NUMERO]);
      const nombre = limpiar(datos[i][COL_NOMBRE]);
      const tipo = limpiar(datos[i][COL_TIPO]);
      const correo = limpiar(datos[i][COL_CORREO]);

      if (!numeroOrden || !nombre || !tipoValido(tipo) || !emailValido(correo)) {
        nuevosEstados.push(["ERROR_DATOS"]);
        log(numeroOrden || `FILA_${fila}`, "ERROR_DATOS", `Fila ${fila}`);
        continue;
      }

      const pdfs = index[numeroOrden];

      if (!pdfs) {
        nuevosEstados.push(["PDF_NO_ENCONTRADO"]);
        log(numeroOrden, "PDF_NO_ENCONTRADO", `N° ${numeroOrden}`);
        continue;
      }

      if (pdfs.length > 1) {
        nuevosEstados.push(["DUPLICADO"]);
        log(numeroOrden, "DUPLICADO", `N° ${numeroOrden}`);
        continue;
      }

      try {
        const mail = plantillaCorreo(tipo, nombre);

        const nombreAdjunto = `Constancia_${numeroOrden}_${slugNombre(nombre)}.pdf`;
        const pdfAdjunto = pdfs[0].getBlob().setName(nombreAdjunto);

        MailApp.sendEmail({
          to: correo,
          subject: mail.subject,
          htmlBody: mail.body,
          attachments: [pdfAdjunto]
        });

        cuotaDisponible--;
        enviadosEnEjecucion++;
        nuevosEstados.push(["ENVIADO"]);
        log(numeroOrden, "ENVIADO", correo);

      } catch (e) {
        nuevosEstados.push(["ERROR_TEMP"]);
        log(numeroOrden, "ERROR_ENVIO", e && e.message ? e.message : String(e));
      }
    }

    /* ===================== ESCRITURA MASIVA ===================== */

    hoja.getRange(2, COL_ESTADO + 1, nuevosEstados.length, 1)
      .setValues(nuevosEstados);

  } finally {
    lock.releaseLock();
  }
}

// Wrapper para compatibilidad con menús / versiones anteriores
function envioMasivoConstancias() {
  envioMasivoConstancias_v3_3();
}
