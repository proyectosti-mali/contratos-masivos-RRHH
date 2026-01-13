# Envío masivo de contratos, convenios y adendas

Script de Google Apps Script (versión 3.2) que automatiza el envío de contratos, convenios de prácticas y adendas en PDF desde Google Drive a una lista de destinatarios en Google Sheets.

## Requisitos previos

- Hoja de cálculo con la pestaña **Hoja 1** y columnas en este orden:
  1. DNI/CE
  2. Nombre completo
  3. Tipo de documento (Laboral, Prácticas, Locación, Adenda)
  4. Período en formato `AAAAMM-AAAAMM` (ej. `202401-202406`)
  5. Correo electrónico
  6. Estado (vacío al inicio)
- Carpeta en Google Drive (ID configurada en el script) que contenga los PDFs firmados.
- Cada PDF debe seguir la convención de nombre `Contrato|Convenio|Adenda_DNI_PERIODO.pdf`, por ejemplo: `Contrato_12345678_202401-202406.pdf`.
- Cuenta con permisos para ejecutar Apps Script y enviar correos mediante `MailApp`.

## Instalación y configuración

1. Abre la hoja de cálculo y ve a **Extensions → Apps Script**.
2. Crea/elimina archivos hasta quedarte con un único archivo `enviosmasivos.js` (o pega el contenido en uno nuevo).
3. Ajusta las constantes si lo necesitas:
   - `NOMBRE_HOJA`: nombre de la pestaña principal.
   - `HOJA_LOG`: nombre de la hoja donde quedará el historial.
   - `FOLDER_ID`: ID de la carpeta de Drive que aloja los PDFs.
4. Guarda el proyecto y autoriza la ejecución la primera vez que lo corras.

## Flujo de trabajo

1. Completa los registros en **Hoja 1** asegurando que el período y el tipo sean válidos.
2. Coloca los PDFs en la carpeta de Drive con el nombre correcto.
3. Desde Apps Script ejecuta la función `envioMasivoContratos_v3_2`.
4. El script:
   - Obtiene un candado (`LockService`) para evitar ejecuciones simultáneas.
   - Indexa todos los PDFs coincidentes.
   - Valida fila por fila y envía correos según plantilla por tipo.
   - Marca el resultado en la columna Estado y registra cada evento en `LOG_ENVÍOS`.

## Estados posibles

- `ENVIADO`: correo enviado correctamente.
- `ERROR_DATOS`: faltan campos, correo inválido o PDF no encontrado.
- `DUPLICADO`: más de un PDF coincide con la misma clave `Tipo_DNI_Periodo`.
- `PENDIENTE_CUOTA`: se detuvo por falta de cuota diaria de correo.
- `ERROR_TEMP`: error inesperado al enviar (reintenta luego).

## Registro (LOG)

Cada ejecución agrega filas en la hoja `LOG_ENVÍOS` con: fecha/hora, DNI, acción y detalle (correo enviado, error, duplicado, etc.). Úsalo para auditoría o depuración.

## Personalización

- **Plantillas de correo**: Modifica la función `plantillaCorreo` para ajustar el asunto o el cuerpo HTML por tipo.
- **Tipos adicionales**: Amplía `tipoValido` y `tipoArchivo` si necesitas nuevas categorías.
- **Regex de PDFs**: Cambia `regexPDF` si adoptas otro esquema de nombres.

## Buenas prácticas

- Ejecuta pruebas con pocas filas antes de un envío masivo.
- No borres manualmente `LOG_ENVÍOS`; puedes crear pivotes o filtros para consultarlo.
- Si editas la hoja mientras corre el script, espera a que finalice para evitar inconsistencias.
- Vigila la cuota diaria de `MailApp`: el script la consulta y se detiene si se agota.
