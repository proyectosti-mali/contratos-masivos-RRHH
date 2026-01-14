/**
 * ============================================================
 * OCR + RENOMBRADO AUTOMÁTICO DE PDFs (PRODUCCIÓN)
 * ------------------------------------------------------------
 * - Extrae DNI del TRABAJADOR (no gerente)
 * - Extrae rango de fechas desde texto (ignora firmas)
 * - Detecta tipo: Contrato / Convenio / Adenda
 * - Renombra solo si hay certeza total
 * - Marca archivos ambiguos como REVISAR_*
 * - Registra todo en hoja OCR_LOG
 * ============================================================
 */

function ocrRenombrarPDFs_PROD() {

  /* ===================== CONFIG ===================== */

  const FOLDER_ID = "10ycIn7wyKTFBhPgYbj-VQvjIdwJ2yJ19"; // carpeta PDFs
  const OCR_LANG = "es";
  const LOG_SHEET = "OCR_LOG";

  const ss = SpreadsheetApp.getActive();
  const logSheet = ss.getSheetByName(LOG_SHEET) || ss.insertSheet(LOG_SHEET);

  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow([
      "Fecha",
      "Archivo Original",
      "Resultado",
      "Nuevo Nombre / Motivo"
    ]);
  }

  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();

  Logger.log("===== INICIO OCR =====");

  while (files.hasNext()) {
    const file = files.next();
    const nombreOriginal = file.getName();

    // Solo PDFs
    if (!nombreOriginal.toLowerCase().endsWith(".pdf")) continue;

    // Ya procesados correctamente
    if (/^(Contrato|Convenio|Adenda)_\d{8,11}_\d{6}-\d{6}\.pdf$/i.test(nombreOriginal)) {
      continue;
    }

    try {
      Logger.log(`Procesando: ${nombreOriginal}`);

      const texto = extraerTextoOCR(file.getId(), OCR_LANG);

      const tipo = detectarTipo(texto);
      const dni = extraerDniTrabajador(texto);
      const periodo = extraerPeriodo(texto);

      // Validaciones estrictas
      if (!dni && !periodo) {
        marcarRevision(file, nombreOriginal, "SIN_DNI_Y_FECHAS", logSheet);
        continue;
      }

      if (!dni) {
        marcarRevision(file, nombreOriginal, "SIN_DNI", logSheet);
        continue;
      }

      if (!periodo) {
        marcarRevision(file, nombreOriginal, "SIN_FECHAS", logSheet);
        continue;
      }

      if (!tipo) {
        marcarRevision(file, nombreOriginal, "SIN_TIPO", logSheet);
        continue;
      }

      const nuevoNombre = `${tipo}_${dni}_${periodo}.pdf`;
      file.setName(nuevoNombre);

      logSheet.appendRow([
        new Date(),
        nombreOriginal,
        "RENOMBRADO",
        nuevoNombre
      ]);

      Logger.log(`✅ Renombrado → ${nuevoNombre}`);

    } catch (e) {
      marcarRevision(file, nombreOriginal, "ERROR_OCR", logSheet, e.message);
    }
  }

  Logger.log("===== FIN OCR =====");
}

/* ===================== OCR ===================== */

function extraerTextoOCR(fileId, idioma) {

  const blob = DriveApp.getFileById(fileId).getBlob();

  const recurso = {
    title: "OCR_TEMP",
    mimeType: MimeType.GOOGLE_DOCS
  };

  const doc = Drive.Files.insert(recurso, blob, {
    ocr: true,
    ocrLanguage: idioma
  });

  const texto = DocumentApp
    .openById(doc.id)
    .getBody()
    .getText();

  DriveApp.getFileById(doc.id).setTrashed(true);

  return texto.replace(/\s+/g, " ");
}

/* ===================== EXTRACCIÓN ===================== */

function detectarTipo(texto) {
  if (/adenda/i.test(texto)) return "Adenda";
  if (/práctic|practica/i.test(texto)) return "Convenio";
  if (/contrato de trabajo/i.test(texto)) return "Contrato";
  return null;
}

function extraerDniTrabajador(texto) {
  // Busca DNI asociado explícitamente a EL TRABAJADOR
  const regex =
    /DNI\s*(?:N°|:)?\s*(\d{8,11})[\s\S]{0,200}?EL\s+TRABAJADOR/i;

  const match = texto.match(regex);
  if (match) return match[1];

  // Fallback: solo si hay UN ÚNICO DNI en todo el documento
  const candidatos = [...texto.matchAll(/\b\d{8,11}\b/g)].map(m => m[0]);
  return candidatos.length === 1 ? candidatos[0] : null;
}

function extraerPeriodo(texto) {
  const meses = {
    enero: "01",
    febrero: "02",
    marzo: "03",
    abril: "04",
    mayo: "05",
    junio: "06",
    julio: "07",
    agosto: "08",
    setiembre: "09",
    septiembre: "09",
    octubre: "10",
    noviembre: "11",
    diciembre: "12"
  };

  const regexFecha =
    /(\d{1,2})\s+de\s+(enero|febrero|marzo|abril|mayo|junio|julio|agosto|setiembre|septiembre|octubre|noviembre|diciembre)\s+de\s+(20\d{2})/gi;

  const fechas = [...texto.matchAll(regexFecha)];
  if (fechas.length < 2) return null;

  const inicio = fechas[0];
  const fin = fechas[1];

  return `${inicio[3]}${meses[inicio[2]]}-${fin[3]}${meses[fin[2]]}`;
}

/* ===================== REVISIÓN ===================== */

function marcarRevision(file, nombreOriginal, motivo, logSheet, detalle) {

  if (!nombreOriginal.startsWith("REVISAR_")) {
    file.setName(`REVISAR_${motivo}_${nombreOriginal}`);
  }

  logSheet.appendRow([
    new Date(),
    nombreOriginal,
    "REVISAR",
    motivo + (detalle ? ` | ${detalle}` : "")
  ]);

  Logger.log(`⚠️ REVISAR (${motivo}): ${nombreOriginal}`);
}
/* pendiente añadir Drive API y Google Cloud OCR en servicios avanzados */