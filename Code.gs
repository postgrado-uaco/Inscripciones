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
 *       - Quién tiene acceso: "Cualquier persona" ← OBLIGATORIO para que fetch() funcione desde el navegador
 *  4. Copiá la URL generada y reemplazá SCRIPT_URL en index.html.
 */

// ─── CONSTANTES DE CONFIGURACIÓN ────────────────────────────────────────────

const CARPETA_INSCRIPCIONES_PADRE_ID = '1qcn7akKHbKlXZgNpAAe9XM6B03eO7txP'; // Carpeta que contiene las subcarpetas por año (2025, 2026, 2027...)
const PLANILLA_IMPLEMENTACION_ID = '1-pHtPMxfcnLLQq-0NAI8AxuZtBoyIbKInvXRIh5jvHU';
const HOJA_IMPLEMENTACION       = 'Implementación';
const ACCIONES_PERMITIDAS       = ['listSheets','getSheet','getImplementacion','getAllCursos','getFile','processFile','saveIngreso','getIngresos','saveMontoFila','saveColumna'];
const REGEX_DRIVE_ID            = /^[a-zA-Z0-9_\-]{25,50}$/;

// ─── PUNTO DE ENTRADA ────────────────────────────────────────────────────────

function doGet(e) {
  try {
    const params = e && e.parameter ? e.parameter : {};
    const action = sanitizeString(params.action);
    if (!action) return errorResponse('Parámetro "action" requerido.', 400);
    if (!ACCIONES_PERMITIDAS.includes(action)) return errorResponse('Acción no permitida: ' + action, 403);

    switch (action) {

      case 'listSheets':
        return handleListSheets();

      case 'getSheet': {
        const sheetId = sanitizeString(params.sheetId);
        if (!sheetId) return errorResponse('sheetId requerido.', 400);
        if (!REGEX_DRIVE_ID.test(sheetId)) return errorResponse('sheetId inválido.', 400);
        return handleGetSheet(sheetId);
      }

      case 'getImplementacion':
        return handleGetImplementacion();

      case 'getAllCursos':
        return handleGetAllCursos();

      case 'getFile': {
        const fileId = sanitizeString(params.fileId);
        if (!fileId) return errorResponse('fileId requerido.', 400);
        if (!REGEX_DRIVE_ID.test(fileId)) return errorResponse('fileId inválido.', 400);
        return handleGetFile(fileId);
      }

      case 'processFile': {
        const fileId = sanitizeString(params.fileId);
        const tipo   = sanitizeString(params.tipo);
        if (!fileId) return errorResponse('fileId requerido.', 400);
        if (!REGEX_DRIVE_ID.test(fileId)) return errorResponse('fileId inválido.', 400);
        if (!tipo) return errorResponse('tipo requerido.', 400);
        return handleProcessFile(fileId, tipo);
      }

      case 'saveIngreso': {
        const sheetId = sanitizeString(params.sheetId);
        const monto   = sanitizeString(params.monto);
        if (!sheetId) return errorResponse('sheetId requerido.', 400);
        if (!REGEX_DRIVE_ID.test(sheetId)) return errorResponse('sheetId inválido.', 400);
        return handleSaveIngreso(sheetId, monto);
      }

      case 'getIngresos':
        return handleGetIngresos();

      case 'saveMontoFila': {
        const sheetId = sanitizeString(params.sheetId);
        const rowIdx  = parseInt(params.rowIdx, 10);
        const monto   = sanitizeString(params.monto);
        if (!sheetId) return errorResponse('sheetId requerido.', 400);
        if (!REGEX_DRIVE_ID.test(sheetId)) return errorResponse('sheetId inválido.', 400);
        if (isNaN(rowIdx) || rowIdx < 1) return errorResponse('rowIdx inválido.', 400);
        return handleSaveMontoFila(sheetId, rowIdx, monto);
      }

      case 'saveColumna': {
        const sheetId    = sanitizeString(params.sheetId);
        const rowIdx     = parseInt(params.rowIdx, 10);
        const colName    = sanitizeString(params.colName);
        const value      = sanitizeString(params.value);
        const sheetIndex = parseInt(params.sheetIndex || '0', 10) || 0;
        if (!sheetId) return errorResponse('sheetId requerido.', 400);
        if (!REGEX_DRIVE_ID.test(sheetId)) return errorResponse('sheetId inválido.', 400);
        if (isNaN(rowIdx) || rowIdx < 1) return errorResponse('rowIdx inválido.', 400);
        if (!colName || colName.length > 100) return errorResponse('colName inválido.', 400);
        return handleSaveColumna(sheetId, rowIdx, colName, value, sheetIndex);
      }

      default:
        return errorResponse('Acción desconocida.', 400);
    }

  } catch (err) {
    Logger.log('Error inesperado en doGet: ' + err.message);
    return errorResponse('Error interno del servidor.', 500);
  }
}

// ─── HANDLERS ────────────────────────────────────────────────────────────────

function handleListSheets() {
  const folder = getCarpetaInscripcionesDelAnio();
  if (!folder) return errorResponse('Carpeta de inscripciones no encontrada.', 404);
  const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  const sheets = [];
  while (files.hasNext()) {
    const file = files.next();
    sheets.push({ id: sanitizeString(file.getId()), name: sanitizeString(file.getName()) });
  }
  return jsonResponse({ sheets: sheets, carpeta: sanitizeString(folder.getName()) });
}

/**
 * Devuelve los datos de la hoja con más filas de la planilla.
 * Incluye sheetIndex para que saveColumna escriba en la misma hoja.
 */
function handleGetSheet(sheetId) {
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(sheetId);
  } catch (err) {
    Logger.log('No se pudo abrir planilla ' + sheetId + ': ' + err.message);
    return errorResponse('No se pudo acceder a la planilla.', 404);
  }

  // Buscar la hoja con más filas (donde están los datos del formulario)
  const sheets = spreadsheet.getSheets();
  let sheet = sheets[0];
  let sheetIndex = 0;
  let maxRows = 0;
  sheets.forEach(function(s, i) {
    const rows = s.getLastRow();
    if (rows > maxRows) { maxRows = rows; sheet = s; sheetIndex = i; }
  });

  if (!sheet) return jsonResponse({ values: [], formAbierto: true, sheetIndex: 0 });

  const rawValues = sheet.getDataRange().getValues();
  const values = rawValues.map(row => row.map(cell => sanitizeCellValue(cell)));

  let formAbierto = true;
  try {
    const formUrl = spreadsheet.getFormUrl();
    if (formUrl) formAbierto = FormApp.openByUrl(formUrl).isAcceptingResponses();
  } catch (e) {}

  return jsonResponse({ values: values, formAbierto: formAbierto, sheetIndex: sheetIndex });
}

function handleGetImplementacion() {
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(PLANILLA_IMPLEMENTACION_ID);
  } catch (err) {
    return errorResponse('No se pudo acceder a la planilla de implementación.', 404);
  }

  const sheets = spreadsheet.getSheets();
  const anio = new Date().getFullYear().toString();

  // 1. Buscar hoja que contenga el año actual en su nombre
  let sheet = sheets.find(s => s.getName().includes(anio)) || null;

  // 2. Si no encuentra el año actual, informar error con los nombres disponibles
  if (!sheet) {
    const nombres = sheets.map(s => s.getName()).join(', ');
    return errorResponse('No se encontró una hoja para el año ' + anio + '. Hojas disponibles: ' + nombres, 404);
  }

  Logger.log('getImplementacion: usando hoja "' + sheet.getName() + '"');
  const rawValues = sheet.getDataRange().getValues();
  const values = rawValues.map(row => row.map(cell => sanitizeCellValue(cell)));
  return jsonResponse({ values: values, sheetName: sheet.getName() });
}

function handleGetAllCursos() {
  const implData = {};
  try {
    const ss = SpreadsheetApp.openById(PLANILLA_IMPLEMENTACION_ID);
    const sheet = ss.getSheetByName(HOJA_IMPLEMENTACION) || ss.getSheets()[0];
    if (sheet) {
      const [hdr, ...rows] = sheet.getDataRange().getValues();
      const idx = name => hdr.findIndex(h => sanitizeString(h) === name);
      const cNombre   = idx('Denominacion del curso');
      const cDocente  = idx('Docente/s Responsable/s');
      const cCarga    = idx('Carga Horaria');
      const cFecha    = idx('Fechas de realización');
      const cInscriptos = idx('Cantidad de inscriptos');
      rows.forEach(r => {
        const nombre = sanitizeString(r[cNombre]);
        if (!nombre) return;
        const entry = {
          docente:        cDocente   >= 0 ? sanitizeString(r[cDocente])   : '',
          cargaHoraria:   cCarga     >= 0 ? sanitizeString(r[cCarga])     : '',
          fecha:          cFecha     >= 0 ? sanitizeString(r[cFecha])     : '',
          cantInscriptos: cInscriptos >= 0 ? sanitizeCellValue(r[cInscriptos]) : ''
        };
        implData[nombre.toLowerCase()] = entry;
        const corta = nombre.toLowerCase()
          .replace(/^\d{4}\s*[-–]\s*/i, '')
          .replace(/^(taller|curso|seminario)(\s+de\s+postgrado)?\s*/i, '')
          .trim();
        if (corta && corta !== nombre.toLowerCase()) implData[corta] = entry;
      });
    }
  } catch (e) {
    Logger.log('Error leyendo implementación: ' + e.message);
  }

  const folder = getCarpetaInscripcionesDelAnio();
  if (!folder) return errorResponse('Carpeta de inscripciones no encontrada.', 404);

  const archivos = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  const cursos = [];

  while (archivos.hasNext()) {
    const file = archivos.next();
    const sheetId   = file.getId();
    const sheetName = sanitizeString(file.getName());
    const nombre    = sheetName
      .replace(/^\d{4}\s*[-–]\s*/i, '')
      .replace(/\s*\(File responses\)\s*$/i, '')
      .replace(/\s*\(respuestas de formulario\s*\d*\)\s*$/i, '')
      .replace(/\s*\(respuestas\)\s*$/i, '')
      .replace(/["'`«»\u201c\u201d\u2018\u2019]/g, '')
      .trim();

    const normClave = s => s.toLowerCase()
      .replace(/^\d{4}\s*[-–]\s*/i, '')
      .replace(/^(taller|curso|seminario)(\s+de\s+postgrado)?\s*/i, '')
      .trim();
    const impl = implData[nombre.toLowerCase()]
      || implData[normClave(nombre)]
      || Object.entries(implData).find(([k]) => {
           const kn = normClave(k), nn = normClave(nombre);
           return kn === nn || kn.includes(nn) || nn.includes(kn);
         })?.[1]
      || {};

    let inscriptos = 0;
    let formAbierto = true;
    try {
      const ss   = SpreadsheetApp.openById(sheetId);
      const hoja = ss.getSheets()[0];
      if (hoja) {
        inscriptos = Math.max(0, hoja.getLastRow() - 1);
        try {
          const formUrl = ss.getFormUrl();
          if (formUrl) formAbierto = FormApp.openByUrl(formUrl).isAcceptingResponses();
        } catch (fe) {}
      }
    } catch (e) {
      Logger.log('Error leyendo planilla ' + sheetId + ': ' + e.message);
    }

    cursos.push({
      id:           sanitizeString(sheetId),
      name:         sheetName,
      nombre:       sanitizeString(nombre),
      inscriptos:   inscriptos,
      abierto:      formAbierto,
      docente:      impl.docente      || '',
      cargaHoraria: impl.cargaHoraria || '',
      fecha:        impl.fecha        || ''
    });
  }

  return jsonResponse({ cursos: cursos });
}

function handleGetFile(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    return jsonResponse({
      base64:   Utilities.base64Encode(blob.getBytes()),
      mimeType: blob.getContentType()
    });
  } catch(e) {
    Logger.log('Error getFile ' + fileId + ': ' + e.message);
    return errorResponse('No se pudo acceder al archivo.', 404);
  }
}

const PROMPTS = {
  comprobante:        'Comprobante de transferencia bancaria argentina. Responde SOLO JSON sin texto extra:\n{"monto":"...","fecha":"...","cbu_destinatario":"...","id_transaccion":"..."}\nmonto: solo números y punto decimal (ej: 1500.50). fecha: DD-MM-AAAA. cbu_destinatario: CBU o CVU del destinatario/receptor de la transferencia, solo dígitos sin guiones ni espacios. id_transaccion: número o código de operación. null si no podés leer.',
  comprobante_simple: 'Comprobante de transferencia bancaria argentina. Responde SOLO JSON sin texto extra: {"monto":"...","fecha":"...","id_transaccion":"..."} monto: solo numeros y punto decimal (ej: 1500.50). fecha: DD-MM-AAAA. id_transaccion: numero o codigo de operacion/comprobante. null si no podes leer algun campo.',
  titulo:             'Este es un título universitario argentino. Responde SOLO JSON sin texto extra:\n{"nombre_apellido":"...","carrera":"...","universidad":"..."}\nnombre_apellido: nombre completo del graduado tal como figura en el título. carrera: nombre completo de la carrera o título otorgado. universidad: nombre completo de la institución. null si no podés leer.',
  dni:                'Documento Nacional de Identidad argentino. Extraé los datos exactamente como figuran en el documento. Responde SOLO JSON sin texto extra ni markdown:\n{"numero_dni":"...","apellidos":"...","nombres":"..."}\n- numero_dni: los 8 dígitos del número de documento, sin puntos ni espacios.\n- apellidos: TODOS los apellidos completos tal como figuran impresos en el DNI, en mayúsculas. NO abreviar.\n- nombres: TODOS los nombres completos tal como figuran impresos en el DNI, en mayúsculas. NO abreviar ni truncar.',
  dni_simple:         'Documento de identidad argentino (DNI). Responde SOLO JSON sin texto extra: {"apellido":"...","nombre":"...","dni":"..."} dni: solo digitos sin puntos ni espacios. null si no podes leer algun campo.'
};

function handleProcessFile(fileId, tipo) {
  const prompt = PROMPTS[tipo];
  if (!prompt) return errorResponse('Tipo inválido: ' + tipo, 400);

  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) return errorResponse('GEMINI_API_KEY no configurada.', 500);

  const MIMETYPES_SOPORTADOS = ['image/jpeg','image/jpg','image/png','image/gif','image/webp','application/pdf'];

  let base64, mimeType;
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    base64   = Utilities.base64Encode(blob.getBytes());
    mimeType = blob.getContentType();
  } catch(e) {
    return errorResponse('No se pudo acceder al archivo.', 404);
  }

  if (!MIMETYPES_SOPORTADOS.includes(mimeType)) {
    return errorResponse('Formato no compatible (' + mimeType + ').', 400);
  }

  const gemUrl = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;
  const payload = JSON.stringify({
    contents: [{ parts: [
      { inline_data: { mime_type: mimeType, data: base64 } },
      { text: prompt }
    ]}],
    generationConfig: { temperature: 0 }
  });

  try {
    const res = UrlFetchApp.fetch(gemUrl, { method:'post', contentType:'application/json', payload:payload, muteHttpExceptions:true });
    const result = JSON.parse(res.getContentText());
    if (result.error) return errorResponse('Gemini: ' + sanitizeString(result.error.message), 500);
    const text = result.candidates?.[0]?.content?.parts?.[0]?.text || '';
    let parsed;
    try   { parsed = JSON.parse(text.trim()); }
    catch { parsed = JSON.parse(text.replace(/```json|```/g, '').trim()); }
    return jsonResponse({ result: parsed });
  } catch(e) {
    return errorResponse('Error llamando Gemini: ' + e.message, 500);
  }
}

function handleSaveIngreso(sheetId, monto) {
  const montoNum = parseFloat(monto);
  if (isNaN(montoNum) || montoNum < 0) return errorResponse('monto inválido.', 400);
  try {
    const props = PropertiesService.getScriptProperties();
    const raw   = props.getProperty('pagos_totales') || '{}';
    const pagos = JSON.parse(raw);
    pagos[sheetId] = montoNum;
    props.setProperty('pagos_totales', JSON.stringify(pagos));
    return jsonResponse({ ok: true });
  } catch(e) {
    return errorResponse('Error guardando ingreso.', 500);
  }
}

function handleGetIngresos() {
  try {
    const props = PropertiesService.getScriptProperties();
    const raw   = props.getProperty('pagos_totales') || '{}';
    const pagos = JSON.parse(raw);
    const total = Object.values(pagos).reduce((s, v) => s + (Number(v) || 0), 0);
    return jsonResponse({ pagos: pagos, total: total });
  } catch(e) {
    return errorResponse('Error leyendo ingresos.', 500);
  }
}

function handleSaveMontoFila(sheetId, rowIdx, monto) {
  const montoNum = parseFloat(monto);
  if (isNaN(montoNum) || montoNum < 0) return errorResponse('monto inválido.', 400);
  try {
    const ss      = SpreadsheetApp.openById(sheetId);
    const sheet   = ss.getSheets()[0];
    const lastCol = sheet.getLastColumn();
    const hdr     = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const COL_NAME = 'Monto verificado';
    let colPos = hdr.findIndex(h => String(h).trim() === COL_NAME) + 1;
    if (colPos === 0) { colPos = lastCol + 1; sheet.getRange(1, colPos).setValue(COL_NAME); }
    sheet.getRange(rowIdx + 1, colPos).setValue(montoNum);
    return jsonResponse({ ok: true });
  } catch(e) {
    return errorResponse('Error guardando monto.', 500);
  }
}

/**
 * Guarda un valor en una columna de la planilla.
 * sheetIndex: índice de la hoja (0-based) donde están los datos.
 * Si la columna no existe, la crea al final del encabezado.
 */
function handleSaveColumna(sheetId, rowIdx, colName, value, sheetIndex) {
  try {
    const ss     = SpreadsheetApp.openById(sheetId);
    const sheets = ss.getSheets();
    // Usar la hoja indicada por sheetIndex, con fallback a la de más filas
    let sheet = sheets[Math.min(sheetIndex || 0, sheets.length - 1)];
    if (!sheet) sheet = sheets[0];

    const lastCol = sheet.getLastColumn();
    const hdr     = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    let colPos = hdr.findIndex(h => String(h).trim() === colName) + 1;
    if (colPos === 0) {
      colPos = lastCol + 1;
      sheet.getRange(1, colPos).setValue(colName);
    }
    // rowIdx es 1-based (fila 1 = primera fila de datos, fila 2 en la hoja)
    sheet.getRange(rowIdx + 1, colPos).setValue(value);
    Logger.log('saveColumna OK: hoja=' + sheet.getName() + ' fila=' + (rowIdx+1) + ' col=' + colPos + ' colName=' + colName);
    return jsonResponse({ ok: true, sheet: sheet.getName(), row: rowIdx + 1, col: colPos });
  } catch(e) {
    Logger.log('Error saveColumna: ' + e.message);
    return errorResponse('Error guardando columna: ' + e.message, 500);
  }
}

// ─── HELPERS ────────────────────────────────────────────────────────────────

function sanitizeString(value) {
  if (value === null || value === undefined) return '';
  return String(value).replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '').trim().slice(0, 2000);
}

function sanitizeCellValue(cell) {
  if (cell instanceof Date) return cell.toISOString();
  if (typeof cell === 'number') return cell;
  if (typeof cell === 'boolean') return cell;
  return sanitizeString(cell);
}

function getDriveFolder(folderId) {
  if (!folderId || !REGEX_DRIVE_ID.test(folderId)) return null;
  try { return DriveApp.getFolderById(folderId); }
  catch (e) { Logger.log('Carpeta no encontrada: ' + folderId); return null; }
}

/**
 * Devuelve la carpeta de inscripciones correspondiente al año actual.
 * En Drive hay una carpeta por año (ej: "2026", "2027", ...) todas
 * dentro de CARPETA_INSCRIPCIONES_PADRE_ID. Si la carpeta del año
 * actual todavía no existe (ej: el 1° de enero, antes de crearla),
 * cae de respaldo a la carpeta de año más reciente que encuentre.
 */
function getCarpetaInscripcionesDelAnio() {
  const padre = getDriveFolder(CARPETA_INSCRIPCIONES_PADRE_ID);
  if (!padre) return null;

  const anio = new Date().getFullYear().toString();
  const exacta = padre.getFoldersByName(anio);
  if (exacta.hasNext()) return exacta.next();

  // No existe todavía la carpeta del año actual: usar la más reciente como respaldo
  let mejor = null, mejorAnio = -1;
  const todas = padre.getFolders();
  while (todas.hasNext()) {
    const f = todas.next();
    const n = parseInt(f.getName().trim(), 10);
    if (!isNaN(n) && n > mejorAnio) { mejorAnio = n; mejor = f; }
  }
  if (mejor) {
    Logger.log('No existe todavía la carpeta "' + anio + '"; se usa "' + mejor.getName() + '" como respaldo.');
    return mejor;
  }

  Logger.log('No se encontró ninguna carpeta de año dentro de la carpeta padre.');
  return null;
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function errorResponse(message, code) {
  Logger.log('Error ' + code + ': ' + message);
  return ContentService.createTextOutput(JSON.stringify({ error: sanitizeString(message), code: code })).setMimeType(ContentService.MimeType.JSON);
}
