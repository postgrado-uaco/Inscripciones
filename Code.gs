/**
 * Google Apps Script – Backend con validación de entradas y salidas
 * Secretaría de Investigación y Postgrado – UACO/UNPA
 *
 * INSTRUCCIONES DE DEPLOY:
 *  1. Abrí script.google.com y creá un nuevo proyecto.
 *  2. Pegá este archivo como "Code.gs".
 *  3. En "Implementar > Nueva implementación":
 *       - Tipo: Aplicación web
 *       - Ejecutar como: Yo (tu cuenta)
 *       - Quién tiene acceso: "Cualquier usuario de Google" (o "Sólo yo" si es uso interno)
 *  4. Copiá la URL generada y reemplazá SCRIPT_URL en index.html.
 */

// ─── CONSTANTES DE CONFIGURACIÓN ────────────────────────────────────────────

/** ID de la carpeta de Drive donde están las planillas de inscripción. */
const CARPETA_INSCRIPCIONES_ID = '1AQnG88tlwQGkGYeTz3qJ_cbEEa2sV2NV';

/** ID de la planilla de implementación de cursos. */
const PLANILLA_IMPLEMENTACION_ID = '1-pHtPMxfcnLLQq-0NAI8AxuZtBoyIbKInvXRIh5jvHU';

/** Nombre de la hoja dentro de la planilla de implementación. */
const HOJA_IMPLEMENTACION = 'Implementación';

/** Acciones permitidas (whitelist). */
const ACCIONES_PERMITIDAS = ['listSheets', 'getSheet', 'getImplementacion'];

/** Regex para validar un ID de Google Drive (alfanumérico + guiones bajos, 25-50 chars). */
const REGEX_DRIVE_ID = /^[a-zA-Z0-9_\-]{25,50}$/;

// ─── PUNTO DE ENTRADA ────────────────────────────────────────────────────────

/**
 * Maneja todas las solicitudes GET entrantes.
 * Valida entradas, ejecuta la acción y sanitiza la salida.
 */
function doGet(e) {
  try {
    const params = e && e.parameter ? e.parameter : {};

    // 1. Validar que se recibió una acción
    const action = sanitizeString(params.action);
    if (!action) {
      return errorResponse('Parámetro "action" requerido.', 400);
    }

    // 2. Validar que la acción esté en la whitelist
    if (!ACCIONES_PERMITIDAS.includes(action)) {
      return errorResponse('Acción no permitida: ' + action, 403);
    }

    // 3. Dispatchar según la acción
    switch (action) {
      case 'listSheets':
        return handleListSheets();

      case 'getSheet': {
        const sheetId = sanitizeString(params.sheetId);
        if (!sheetId) return errorResponse('Parámetro "sheetId" requerido.', 400);
        if (!REGEX_DRIVE_ID.test(sheetId)) return errorResponse('sheetId inválido.', 400);
        return handleGetSheet(sheetId);
      }

      case 'getImplementacion':
        return handleGetImplementacion();

      default:
        return errorResponse('Acción desconocida.', 400);
    }

  } catch (err) {
    Logger.log('Error inesperado en doGet: ' + err.message);
    return errorResponse('Error interno del servidor.', 500);
  }
}

// ─── HANDLERS ────────────────────────────────────────────────────────────────

/**
 * Lista todas las planillas (archivos Google Sheets) dentro de la carpeta
 * de inscripciones. Devuelve id y name de cada una.
 */
function handleListSheets() {
  const folder = getDriveFolder(CARPETA_INSCRIPCIONES_ID);
  if (!folder) return errorResponse('Carpeta de inscripciones no encontrada.', 404);

  const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  const sheets = [];

  while (files.hasNext()) {
    const file = files.next();
    sheets.push({
      id:   sanitizeString(file.getId()),
      name: sanitizeString(file.getName())
    });
  }

  return jsonResponse({ sheets: sheets });
}

/**
 * Devuelve las filas de la primera hoja de una planilla dado su ID.
 * Incluye el estado del formulario de Google si está disponible.
 */
function handleGetSheet(sheetId) {
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(sheetId);
  } catch (err) {
    Logger.log('No se pudo abrir planilla ' + sheetId + ': ' + err.message);
    return errorResponse('No se pudo acceder a la planilla.', 404);
  }

  const sheet = spreadsheet.getSheets()[0];
  if (!sheet) return jsonResponse({ values: [], formAbierto: true });

  const rawValues = sheet.getDataRange().getValues();

  // Sanitizar cada celda de la salida
  const values = rawValues.map(row =>
    row.map(cell => sanitizeCellValue(cell))
  );

  // Intentar detectar si el formulario vinculado está abierto
  let formAbierto = true;
  try {
    const formUrl = sheet.getParent().getFormUrl();
    if (formUrl) {
      const form = FormApp.openByUrl(formUrl);
      formAbierto = form.isAcceptingResponses();
    }
  } catch (e) {
    // Sin formulario vinculado: se asume abierto
  }

  return jsonResponse({ values: values, formAbierto: formAbierto });
}

/**
 * Devuelve los datos de la planilla de implementación de cursos.
 */
function handleGetImplementacion() {
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(PLANILLA_IMPLEMENTACION_ID);
  } catch (err) {
    Logger.log('No se pudo abrir planilla de implementación: ' + err.message);
    return errorResponse('No se pudo acceder a la planilla de implementación.', 404);
  }

  const sheet = spreadsheet.getSheetByName(HOJA_IMPLEMENTACION)
    || spreadsheet.getSheets()[0];

  if (!sheet) return jsonResponse({ values: [] });

  const rawValues = sheet.getDataRange().getValues();

  const values = rawValues.map(row =>
    row.map(cell => sanitizeCellValue(cell))
  );

  return jsonResponse({ values: values });
}

// ─── HELPERS DE VALIDACIÓN Y SANITIZACIÓN ───────────────────────────────────

/**
 * Convierte cualquier valor a string, elimina caracteres de control
 * y trunca a 2000 caracteres para evitar payloads excesivos.
 */
function sanitizeString(value) {
  if (value === null || value === undefined) return '';
  // Eliminar caracteres de control (exceptuando tabuladores y saltos de línea normales)
  return String(value)
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '')
    .trim()
    .slice(0, 2000);
}

/**
 * Sanitiza el valor de una celda de Sheets:
 * - Fechas → ISO string
 * - Números → número
 * - Strings → sanitizeString
 * - Booleanos → booleano
 */
function sanitizeCellValue(cell) {
  if (cell instanceof Date) return cell.toISOString();
  if (typeof cell === 'number') return cell;
  if (typeof cell === 'boolean') return cell;
  return sanitizeString(cell);
}

/**
 * Obtiene una carpeta de Drive de forma segura.
 * @param {string} folderId
 * @returns {GoogleAppsScript.Drive.Folder|null}
 */
function getDriveFolder(folderId) {
  if (!folderId || !REGEX_DRIVE_ID.test(folderId)) return null;
  try {
    return DriveApp.getFolderById(folderId);
  } catch (e) {
    Logger.log('Carpeta no encontrada: ' + folderId);
    return null;
  }
}

// ─── HELPERS DE RESPUESTA ────────────────────────────────────────────────────

/**
 * Devuelve una respuesta JSON exitosa.
 */
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Devuelve una respuesta de error estructurada.
 * @param {string} message  Mensaje legible.
 * @param {number} code     Código HTTP semántico (para logging; Apps Script siempre responde 200).
 */
function errorResponse(message, code) {
  Logger.log('Error ' + code + ': ' + message);
  return ContentService
    .createTextOutput(JSON.stringify({ error: sanitizeString(message), code: code }))
    .setMimeType(ContentService.MimeType.JSON);
}
