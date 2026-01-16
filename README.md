# Envío masivo de contratos, convenios y adendas

Script de **Google Apps Script** (versión **3.2 – Producción**) que automatiza el envío de **contratos, convenios de prácticas y adendas en PDF** desde Google Drive a una lista de destinatarios gestionada en Google Sheets.

---

## Requisitos previos

- Hoja de cálculo con la pestaña **Hoja 1** y **las columnas en este orden exacto**:
  1. **DNI / CE**
  2. **Nombre completo**
  3. **Tipo de documento**  
     Valores admitidos: `Laboral`, `Prácticas`, `Locación`, `Adenda`
  4. **Período** en formato `AAAAMM_AAAAMM`  
     Ejemplo: `202601_202603`
  5. **Correo electrónico**
  6. **Estado** (vacío al inicio)

- Carpeta en **Google Drive** (ID configurado en el script) que contenga los PDFs firmados.
- Los archivos PDF deben seguir **estrictamente** esta convención de nombre:
Contrato Laboral_87654321_NOMBRECOMPLETO_202601_202603.pdf
Contrato_87654321_NOMBRECOMPLETO_202601_202603.pdf
Convenio_87654321_NOMBRECOMPLETO_202601_202603.pdf
Adenda_87654321_NOMBRECOMPLETO_202601_202603.pdf


> 📌 El nombre completo se ignora para la validación, pero **debe existir como bloque** en el nombre del archivo.

- Cuenta con permisos para:
  - Ejecutar Apps Script
  - Enviar correos mediante `MailApp`
  - Acceder a Google Drive

---

## Instalación y configuración

1. Abre la hoja de cálculo y ve a  
   **Extensiones → Apps Script**
2. Crea o deja un único archivo (por ejemplo `enviosmasivos.js`) y pega el script.
3. Ajusta las constantes si es necesario:
   - `NOMBRE_HOJA`: nombre de la hoja principal.
   - `HOJA_LOG`: nombre de la hoja de auditoría.
   - `FOLDER_ID`: ID de la carpeta de Drive que contiene los PDFs.
4. Guarda el proyecto.
5. Ejecuta la función `envioMasivoContratos_v3_2` por primera vez y **autoriza los permisos**.

---

## Flujo de trabajo

1. Completa los registros en **Hoja 1** asegurando:
   - DNI válido
   - Tipo correcto
   - Período en formato `AAAAMM_AAAAMM`
   - Correo válido
2. Coloca los PDFs firmados en la carpeta de Drive respetando el formato de nombre.
3. Ejecuta la función `envioMasivoContratos_v3_2` desde Apps Script.
4. El script:
   - Obtiene un candado (`LockService`) para evitar ejecuciones simultáneas.
   - Indexa todos los PDFs válidos de la carpeta.
   - Procesa fila por fila:
     - Valida datos
     - Busca el PDF correspondiente
     - Envía el correo según la plantilla
   - Actualiza la columna **Estado**.
   - Registra cada evento en la hoja **LOG_ENVÍOS**.

---

## Estados posibles

| Estado | Significado |
|------|------------|
| `ENVIADO` | Correo enviado correctamente |
| `ERROR_DATOS` | Datos inválidos o PDF no encontrado |
| `DUPLICADO` | Más de un PDF coincide con la misma clave |
| `PENDIENTE_CUOTA` | Se detuvo por agotamiento de cuota diaria |
| `ERROR_TEMP` | Error inesperado al enviar (reintentar luego) |

---

## Registro (LOG_ENVÍOS)

Cada acción genera una fila en la hoja **LOG_ENVÍOS** con:

- Fecha y hora
- DNI
- Acción (`ENVIADO`, `ERROR_DATOS`, `DUPLICADO`, etc.)
- Detalle (correo, clave del PDF, mensaje de error)

Útil para **auditoría**, **seguimiento** y **depuración**.

---

## Personalización

- **Plantillas de correo**  
  Modifica la función `plantillaCorreo` para ajustar asunto o cuerpo HTML según el tipo.

- **Tipos adicionales**  
  Amplía `tipoValido` y `tipoArchivo` si necesitas nuevas categorías.

- **Formato de PDFs**  
  Ajusta la expresión regular `regexPDF` si cambia el esquema de nombres.

---

## Buenas prácticas

- Realiza pruebas con pocas filas antes de envíos masivos.
- No edites la hoja mientras el script está en ejecución.
- No borres manualmente `LOG_ENVÍOS`; usa filtros o tablas dinámicas.
- Controla la **cuota diaria de `MailApp`**: el script la valida y se detiene automáticamente si se agota.
