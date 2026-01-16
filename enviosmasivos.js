/**
 * ============================================================
 * ENVÍO MASIVO DE CONTRATOS / CONVENIOS / ADENDAS
 * Versión: 3.2 (PRODUCCIÓN)
 * ============================================================
 */

function envioMasivoContratos_v3_2() {

  /* ===================== LOCK ===================== */
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error("Otro proceso está en ejecución");
  }

  try {

    /* ===================== CONFIG ===================== */

    const NOMBRE_HOJA = "Hoja 1";
    const HOJA_LOG = "LOG_ENVÍOS";
    const FOLDER_ID = "10ycIn7wyKTFBhPgYbj-VQvjIdwJ2yJ19";

    // Columnas (0-based)
    const COL_DNI = 0;
    const COL_NOMBRE = 1;
    const COL_TIPO = 2;
    const COL_PERIODO = 3;
    const COL_CORREO = 4;
    const COL_ESTADO = 5;

    const hoja = SpreadsheetApp.getActive().getSheetByName(NOMBRE_HOJA);
    if (!hoja) throw new Error("No se encontró la hoja principal");

    const ss = SpreadsheetApp.getActive();
    const logSheet = ss.getSheetByName(HOJA_LOG) || ss.insertSheet(HOJA_LOG);

    const datos = hoja.getDataRange().getValues();
    const nuevosEstados = [];

    /* ===================== HELPERS ===================== */

    const limpiar = v => v === null || v === undefined ? "" : String(v).trim();

    const normalizarDNI = v => {
      const s = limpiar(v).replace(/^'/, "").replace(/\D/g, "");
      return (s.length >= 8 && s.length <= 11) ? s : "";
    };

    // ✅ PERIODO CON _
    const validarPeriodo = p => /^\d{6}_\d{6}$/.test(p) ? p : "";

    const emailValido = e => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e);

    const tipoValido = t =>
      ["laboral", "prácticas", "practicas", "locación", "locacion", "adenda"]
        .includes(t.toLowerCase());

    const tipoArchivo = t => {
      t = t.toLowerCase();
      if (t.includes("práct")) return "convenio";
      if (t === "adenda") return "adenda";
      return "contrato";
    };

    const plantillaCorreo = (tipoContrato, dni, nombre) => {
      const tipo = tipoContrato.toLowerCase();

      if (tipo === "prácticas" || tipo === "practicas") {
        return {
          subject: `Convenio ${tipoContrato} - ${dni} - ${nombre}`,
          body: `
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
          body: `
        Estimado(a) <b>${nombre}</b>,<br><br>
        Cumplimos con enviar su contrato debidamente firmado por ambas partes.<br><br>
        Saludos,<br><br>
        <strong>Talento Humano</strong><br>
        <strong>ASOCIACIÓN MUSEO DE ARTE DE LIMA</strong>
      `
        };
      }

      if (tipo === "adenda") {
        return {
          subject: `Adenda ${tipoContrato} - ${dni} - ${nombre}`,
          body: `
        Estimado(a) <b>${nombre}</b>,<br><br>
        Cumplimos con enviar su adenda debidamente firmada por ambas partes.<br><br>
        Saludos,<br><br>
        <strong>Talento Humano</strong><br>
        <strong>ASOCIACIÓN MUSEO DE ARTE DE LIMA</strong>
      `
        };
      }

      // Default: Locación
      return {
        subject: `Contrato ${tipoContrato} - ${dni} - ${nombre}`,
        body: `
      Estimado(a) <b>${nombre}</b>,<br><br>
      Cumplimos con enviar su contrato bajo la modalidad <b>Locación de Servicio</b>
      debidamente firmado por ambas partes.<br><br>
      Saludos,<br><br>
      <strong>Talento Humano</strong><br>
      <strong>ASOCIACIÓN MUSEO DE ARTE DE LIMA</strong>
    `
      };
    };

    const log = (dni, accion, detalle) => {
      logSheet.appendRow([new Date(), dni, accion, detalle]);
    };

    /* ===================== CUOTA ===================== */

    let cuotaDisponible = MailApp.getRemainingDailyQuota();
    if (cuotaDisponible <= 0) {
      throw new Error("Cuota de correos agotada");
    }

    /* ===================== INDEXAR PDFs ===================== */

    const index = {};

    // ✅ NUEVO FORMATO PDF
    const regexPDF = /^(Contrato Laboral|Contrato|Convenio|Adenda)_(\d{8,11})_([^_]+)_(\d{6})_(\d{6})\.pdf$/i;

    const files = DriveApp.getFolderById(FOLDER_ID).getFiles();
    while (files.hasNext()) {
      const f = files.next();
      const m = f.getName().match(regexPDF);
      if (!m) continue;

      const periodoPDF = `${m[4]}_${m[5]}`;
      const clave = `${m[1].toLowerCase()}_${m[2]}_${periodoPDF}`;

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

      const dni = normalizarDNI(datos[i][COL_DNI]);
      const nombre = limpiar(datos[i][COL_NOMBRE]);
      const tipo = limpiar(datos[i][COL_TIPO]);
      const periodo = validarPeriodo(limpiar(datos[i][COL_PERIODO]));
      const correo = limpiar(datos[i][COL_CORREO]);

      if (!dni || !nombre || !periodo || !tipoValido(tipo) || !emailValido(correo)) {
        nuevosEstados.push(["ERROR_DATOS"]);
        log(dni, "ERROR_DATOS", `Fila ${fila}`);
        continue;
      }

      const clave = `${tipoArchivo(tipo)}_${dni}_${periodo}`;
      const pdfs = index[clave];

      if (!pdfs) {
        nuevosEstados.push(["ERROR_DATOS"]);
        log(dni, "PDF_NO_ENCONTRADO", clave);
        continue;
      }

      if (pdfs.length > 1) {
        nuevosEstados.push(["DUPLICADO"]);
        log(dni, "DUPLICADO", clave);
        continue;
      }

      try {
        const mail = plantillaCorreo(tipo, dni, nombre);
        MailApp.sendEmail({
          to: correo,
          subject: mail.subject,
          htmlBody: mail.body,
          attachments: [pdfs[0].getBlob()]
        });

        cuotaDisponible--;
        nuevosEstados.push(["ENVIADO"]);
        log(dni, "ENVIADO", correo);

      } catch (e) {
        nuevosEstados.push(["ERROR_TEMP"]);
        log(dni, "ERROR_ENVIO", e.message);
      }
    }

    /* ===================== ESCRITURA MASIVA ===================== */

    hoja.getRange(2, COL_ESTADO + 1, nuevosEstados.length, 1)
      .setValues(nuevosEstados);

  } finally {
    lock.releaseLock();
  }
}
