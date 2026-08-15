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
// La plantilla de constancia y la carpeta de destino ya NO están hardcodeadas:
// se configuran desde la página "Configuración" de la app y quedan guardadas
// en PropertiesService (ver getConfigConstancias / handleGuardarConfigConstancias).
const ACCIONES_PERMITIDAS       = ['listSheets','getSheet','getImplementacion','getAllCursos','getFile','processFile','saveIngreso','getIngresos','saveMontoFila','saveColumna','saveColumnas','extraerActa','generarConstancias','getConfigConstancias','subirPlantillaConstancia','guardarCarpetaConstancias','mergeDocs','getTotalIngresos','recalcularTotalIngresos','getConfigComprobantes','guardarCarpetaComprobantes','guardarComprobantesLote','login','logout'];
const REGEX_DRIVE_ID            = /^[a-zA-Z0-9_\-]{25,50}$/;

// ─── PUNTO DE ENTRADA (POST) ─────────────────────────────────────────────────
// Se usa para acciones que necesitan mandar payloads grandes (ej: el PDF/DOC
// del acta en base64) que no entran cómodos en una URL de GET.

function doPost(e) {
  try {
    let body = {};
    try {
      body = JSON.parse(e && e.postData && e.postData.contents ? e.postData.contents : '{}');
    } catch (err) {
      return errorResponse('JSON inválido en la solicitud.', 400);
    }
    const action = sanitizeString(body.action);
    if (!action) return errorResponse('Parámetro "action" requerido.', 400);
    if (!ACCIONES_PERMITIDAS.includes(action)) return errorResponse('Acción no permitida: ' + action, 403);

    // Verificar token de sesión (excepto para login y logout)
    if (action !== 'login' && action !== 'logout') {
      const token = sanitizeString(body.token);
      if (!verificarToken(token)) return errorResponse('Sesión no válida. Iniciá sesión nuevamente.', 401);
    }

    switch (action) {
      case 'login': {
        const email = sanitizeString(body.email).toLowerCase();
        const pwd   = sanitizeString(body.pwd);
        if (!email || !pwd) return errorResponse('Credenciales requeridas.', 400);
        return handleLogin(email, pwd);
      }

      case 'logout': {
        const token = sanitizeString(body.token);
        if (token) CacheService.getScriptCache().remove('sess_' + token);
        return jsonResponse({ ok: true });
      }

      case 'extraerActa': {
        const sheetId  = sanitizeString(body.sheetId); // opcional: si no viene, se autodetecta por el nombre del curso en el acta
        const base64   = body.base64;
        const mimeType = sanitizeString(body.mimeType);
        if (sheetId && !REGEX_DRIVE_ID.test(sheetId)) return errorResponse('sheetId inválido.', 400);
        if (!base64 || !mimeType) return errorResponse('Archivo del acta requerido.', 400);
        return handleExtraerActa(sheetId, base64, mimeType);
      }

      case 'generarConstancias': {
        const sheetId = sanitizeString(body.sheetId);
        const curso   = sanitizeString(body.curso);
        const filas   = Array.isArray(body.filas) ? body.filas : [];
        if (!sheetId) return errorResponse('sheetId requerido.', 400);
        if (!REGEX_DRIVE_ID.test(sheetId)) return errorResponse('sheetId inválido.', 400);
        if (!curso) return errorResponse('curso requerido.', 400);
        if (!filas.length) return errorResponse('No se indicaron filas a generar.', 400);
        return handleGenerarConstancias(sheetId, curso, filas);
      }

      case 'subirPlantillaConstancia': {
        const base64   = body.base64;
        const mimeType = sanitizeString(body.mimeType);
        const nombre   = sanitizeString(body.nombre) || 'Plantilla Constancia';
        if (!base64 || !mimeType) return errorResponse('Archivo de plantilla requerido.', 400);
        return handleSubirPlantillaConstancia(base64, mimeType, nombre);
      }

      case 'guardarCarpetaConstancias': {
        const carpeta = sanitizeString(body.carpeta);
        if (!carpeta) return errorResponse('Carpeta requerida.', 400);
        return handleGuardarCarpetaConstancias(carpeta);
      }

      case 'saveColumnas': {
        const sheetId    = sanitizeString(body.sheetId);
        const rowIdx     = parseInt(body.rowIdx, 10);
        const cols       = body.cols;
        const sheetIndex = parseInt(body.sheetIndex || 0, 10) || 0;
        if (!sheetId) return errorResponse('sheetId requerido.', 400);
        if (!REGEX_DRIVE_ID.test(sheetId)) return errorResponse('sheetId inválido.', 400);
        if (isNaN(rowIdx) || rowIdx < 1) return errorResponse('rowIdx inválido.', 400);
        if (!cols || typeof cols !== 'object') return errorResponse('cols requerido.', 400);
        return handleSaveColumnas(sheetId, rowIdx, cols, sheetIndex);
      }

      case 'guardarCarpetaComprobantes': {
        const carpeta = sanitizeString(body.carpeta);
        if (!carpeta) return errorResponse('Carpeta requerida.', 400);
        return handleGuardarCarpetaComprobantes(carpeta);
      }

      case 'guardarComprobantesLote': {
        const curso       = sanitizeString(body.curso);
        const base64      = body.base64;
        const nombreArchivo = sanitizeString(body.nombreArchivo);
        if (!curso) return errorResponse('curso requerido.', 400);
        if (!base64) return errorResponse('base64 requerido.', 400);
        if (!nombreArchivo) return errorResponse('nombreArchivo requerido.', 400);
        return handleGuardarComprobantesLote(curso, base64, nombreArchivo);
      }

      default:
        return errorResponse('Acción desconocida.', 400);
    }
  } catch (err) {
    Logger.log('Error inesperado en doPost: ' + err.message);
    return errorResponse('Error interno del servidor.', 500);
  }
}

// ─── PUNTO DE ENTRADA ────────────────────────────────────────────────────────

function doGet(e) {
  try {
    const params = e && e.parameter ? e.parameter : {};
    const action = sanitizeString(params.action);
    if (!action) return errorResponse('Parámetro "action" requerido.', 400);
    if (!ACCIONES_PERMITIDAS.includes(action)) return errorResponse('Acción no permitida: ' + action, 403);

    // Verificar token de sesión
    const token = sanitizeString(params.token);
    if (!verificarToken(token)) return errorResponse('Sesión no válida. Iniciá sesión nuevamente.', 401);

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

      case 'mergeDocs': {
        const fileIdsRaw = sanitizeString(params.fileIds);
        if (!fileIdsRaw) return errorResponse('fileIds requerido.', 400);
        const fileIds = fileIdsRaw.split(',').map(s => s.trim()).filter(Boolean);
        if (!fileIds.length) return errorResponse('fileIds vacío.', 400);
        if (fileIds.some(id => !REGEX_DRIVE_ID.test(id))) return errorResponse('Algún fileId es inválido.', 400);
        return handleMergeDocs(fileIds);
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

      case 'getConfigConstancias':
        return handleGetConfigConstancias();

      case 'getConfigComprobantes':
        return handleGetConfigComprobantes();

      case 'getTotalIngresos':
        return handleGetTotalIngresos();

      case 'recalcularTotalIngresos':
        return handleRecalcularTotalIngresos();

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

const MIMETYPES_SOPORTADOS_IA = ['image/jpeg','image/jpg','image/png','image/gif','image/webp','application/pdf'];

/**
 * Si el archivo ya es un formato que Gemini puede leer (imagen o PDF), lo
 * devuelve tal cual. Si no (Word, Excel, PowerPoint, Google Docs/Sheets/
 * Slides, etc.), intenta convertirlo a PDF antes de devolverlo, para que
 * "Controlar pagos"/"Controlar títulos" puedan seguir procesándolo aunque
 * alguien haya subido, por ejemplo, un .docx en vez de una foto o PDF.
 */
function convertirAFormatoSoportado(file) {
  const mime = file.getMimeType();
  if (MIMETYPES_SOPORTADOS_IA.indexOf(mime) >= 0) {
    return { blob: file.getBlob(), mimeType: mime, convertido: false };
  }

  try {
    // Google Docs/Sheets/Slides nativos: exportar directo a PDF
    if (mime === MimeType.GOOGLE_DOCS || mime === MimeType.GOOGLE_SHEETS || mime === MimeType.GOOGLE_SLIDES) {
      return { blob: file.getAs(MimeType.PDF), mimeType: 'application/pdf', convertido: true };
    }
    // Word/Excel/PowerPoint u otros formatos de oficina: convertir a Google
    // Docs temporalmente (con el servicio avanzado de Drive) y exportar
    // ese temporal a PDF; el temporal se borra apenas se usa.
    if (typeof Drive !== 'undefined' && Drive.Files) {
      const resource = { name: 'tmp_convert_' + new Date().getTime(), mimeType: MimeType.GOOGLE_DOCS };
      const inserted = Drive.Files.create(resource, file.getBlob(), { convert: true });
      const tempId = inserted.id;
      try {
        const pdfBlob = DriveApp.getFileById(tempId).getAs(MimeType.PDF);
        return { blob: pdfBlob, mimeType: 'application/pdf', convertido: true };
      } finally {
        try { Drive.Files.remove(tempId); }
        catch(e2) { try { DriveApp.getFileById(tempId).setTrashed(true); } catch(e3) {} }
      }
    }
  } catch(e) {
    Logger.log('No se pudo convertir ' + file.getName() + ' (' + mime + ') a PDF: ' + e.message);
  }

  // No se pudo convertir: se devuelve tal cual (el llamador decidirá si lo rechaza)
  return { blob: file.getBlob(), mimeType: mime, convertido: false };
}

function handleGetFile(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const { blob, mimeType, convertido } = convertirAFormatoSoportado(file);
    if (convertido) Logger.log('Archivo ' + fileId + ' convertido a PDF para poder leerlo.');
    return jsonResponse({
      base64:   Utilities.base64Encode(blob.getBytes()),
      mimeType: mimeType
    });
  } catch(e) {
    Logger.log('Error getFile ' + fileId + ': ' + e.message);
    return errorResponse('No se pudo acceder al archivo: ' + e.message, 404);
  }
}

/**
 * Convierte una imagen suelta (JPG/PNG/etc) en un PDF de una página, usando
 * un Google Doc temporal como intermediario (se borra apenas se usa).
 */
function imagenAPdf(blob) {
  const doc = DocumentApp.create('tmp_merge_' + new Date().getTime());
  const docId = doc.getId();
  try {
    const body = doc.getBody();
    body.clear();
    const img = body.appendImage(blob);
    const maxW = 500; // A4 usable width aprox, en puntos
    if (img.getWidth() > maxW) {
      const ratio = maxW / img.getWidth();
      img.setWidth(maxW);
      img.setHeight(Math.round(img.getHeight() * ratio));
    }
    doc.saveAndClose();
    return DriveApp.getFileById(docId).getAs(MimeType.PDF);
  } finally {
    try {
      if (typeof Drive !== 'undefined' && Drive.Files) { Drive.Files.remove(docId); }
      else { DriveApp.getFileById(docId).setTrashed(true); }
    } catch(e2) { /* no bloqueamos el resultado si esto falla */ }
  }
}

/**
 * Recibe una lista de fileIds (DNI, título, comprobante, etc. de un
 * inscripto) y devuelve cada uno como PDF en base64, convirtiendo las
 * imágenes sueltas a PDF de una página. El front-end los une con pdf-lib.
 */
function handleMergeDocs(fileIds) {
  const pdfs = [];
  const errores = [];

  fileIds.forEach(id => {
    try {
      const file = DriveApp.getFileById(id);
      const conv = convertirAFormatoSoportado(file);
      let pdfBlob;
      if (conv.mimeType === 'application/pdf') {
        pdfBlob = conv.blob;
      } else if (['image/jpeg','image/jpg','image/png','image/gif','image/webp'].indexOf(conv.mimeType) >= 0) {
        pdfBlob = imagenAPdf(conv.blob);
      } else {
        // No se pudo convertir a PDF (formato raro): último intento directo
        pdfBlob = file.getAs(MimeType.PDF);
      }
      pdfs.push(Utilities.base64Encode(pdfBlob.getBytes()));
    } catch(e) {
      errores.push(id + ': ' + e.message);
    }
  });

  return jsonResponse({ pdfs: pdfs, errores: errores });
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

  let base64, mimeType;
  try {
    const file = DriveApp.getFileById(fileId);
    const conv = convertirAFormatoSoportado(file);
    base64   = Utilities.base64Encode(conv.blob.getBytes());
    mimeType = conv.mimeType;
  } catch(e) {
    return errorResponse('No se pudo acceder al archivo.', 404);
  }

  if (MIMETYPES_SOPORTADOS_IA.indexOf(mimeType) < 0) {
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

/** Normaliza el nombre de un curso para poder compararlo (saca tildes, "(respuestas)", prefijos tipo "Taller de Postgrado", etc). */
function normCursoTexto(s) {
  return (s || '').toString().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/^\d{4}\s*[-–]\s*/,'')
    .replace(/^(taller|curso|seminario)(\s+de\s+postgrado)?\s*/,'')
    .replace(/\(respuestas[^)]*\)\s*$/,'')
    .replace(/\(file responses\)\s*$/,'')
    .replace(/["'«»""'']/g,'')
    .replace(/\s+/g,' ').trim();
}

/** Busca, entre las planillas de inscripciones del año, la que mejor matchea con el nombre de curso detectado en el acta. */
function buscarCursoPorNombre(nombreActa) {
  const carpeta = getCarpetaInscripcionesDelAnio();
  if (!carpeta || !nombreActa) return { candidatos: [] };

  const files = carpeta.getFilesByType(MimeType.GOOGLE_SHEETS);
  const candidatos = [];
  while (files.hasNext()) { const f = files.next(); candidatos.push({ id: f.getId(), name: f.getName() }); }

  const qn = normCursoTexto(nombreActa);
  let match = candidatos.find(c => normCursoTexto(c.name) === qn);
  if (!match) {
    match = candidatos.find(c => {
      const cn = normCursoTexto(c.name);
      return cn && qn && (cn.includes(qn) || qn.includes(cn));
    });
  }
  return { match: match || null, candidatos: candidatos };
}

/**
 * Lee el acta escaneada/completada (PDF o DOC/imagen) con Gemini: extrae el
 * nombre del curso y, por cada alumno, apellido/nombre/DNI/condición/
 * calificación. Si no viene un sheetId puntual, detecta el curso solo
 * (matcheando el nombre leído del acta contra las planillas de inscripciones
 * del año). Si no puede detectarlo con confianza, devuelve la lista de
 * cursos disponibles para que el usuario lo elija a mano. Una vez con el
 * curso resuelto, guarda Condición/Calificación en la planilla del curso.
 */
function handleExtraerActa(sheetId, base64, mimeType) {
  const MIMETYPES_SOPORTADOS = ['image/jpeg','image/jpg','image/png','image/gif','image/webp','application/pdf'];
  if (!MIMETYPES_SOPORTADOS.includes(mimeType)) {
    return errorResponse('Formato no compatible (' + mimeType + '). Subí el acta en PDF o imagen.', 400);
  }

  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) return errorResponse('GEMINI_API_KEY no configurada.', 500);

  const prompt = 'Esta es el "FORMULARIO PARA LA EVALUACIÓN DE ALUMNOS" de un curso de posgrado universitario argentino. ' +
    'Arriba tiene un campo "Nombre Identificatorio" con el nombre del curso. Abajo tiene una tabla de alumnos con ' +
    'columnas: Apellidos, Nombres, N° de Documento Único, Título de Grado, "Cumplió con: Asistencia (Sí/No)" y ' +
    '"Evaluación (Sí/No)", y "Calificación de la Evaluación (si corresponde)" ' +
    '(texto libre, ej: "Aprobación, nota 10", "Aprobación, nota7", "Asistencia", o vacío). ' +
    'Responde SOLO JSON sin texto extra ni markdown:\n' +
    '{"curso":"...","alumnos":[{"apellido":"...","nombre":"...","dni":"...","condicion":"...","calificacion":"..."}]}\n' +
    '- curso: el texto completo del campo "Nombre Identificatorio".\n' +
    '- apellido/nombre: tal como figuran en la tabla, en mayúsculas.\n' +
    '- dni: solo dígitos, sin puntos ni espacios. null si no figura.\n' +
    '- condicion (exactamente uno de estos 3 valores):\n' +
    '  · "Aprobado": si la columna de Calificación menciona "aprob" junto con un número de nota, o si Evaluación=Sí.\n' +
    '  · "Asistente": si dice "asistencia" sin nota, o si Asistencia=Sí y Evaluación=No/vacío.\n' +
    '  · "Ausente": si la fila del alumno no tiene nada cargado en Asistencia/Evaluación/Calificación (columnas vacías).\n' +
    '- calificacion: el número de nota que figura en el texto de la columna Calificación (ej: de "Aprobación, nota 10" extraé "10"; de "Aprobación, nota7" extraé "7"). null si es "Asistente" o "Ausente".\n' +
    'Incluí una entrada por cada alumno de la tabla (todas las filas numeradas), en el mismo orden, aunque estén vacías.';

  const gemUrl = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;
  const payload = JSON.stringify({
    contents: [{ parts: [ { inline_data: { mime_type: mimeType, data: base64 } }, { text: prompt } ] }],
    generationConfig: { temperature: 0 }
  });

  let alumnosExtraidos, cursoDetectado;
  try {
    const res = UrlFetchApp.fetch(gemUrl, { method:'post', contentType:'application/json', payload:payload, muteHttpExceptions:true });
    const result = JSON.parse(res.getContentText());
    if (result.error) return errorResponse('Gemini: ' + sanitizeString(result.error.message), 500);
    const text = result.candidates?.[0]?.content?.parts?.[0]?.text || '';
    let parsed;
    try   { parsed = JSON.parse(text.trim()); }
    catch { parsed = JSON.parse(text.replace(/```json|```/g, '').trim()); }
    alumnosExtraidos = parsed.alumnos || [];
    cursoDetectado    = parsed.curso || '';
  } catch(e) {
    return errorResponse('Error llamando Gemini: ' + e.message, 500);
  }

  // Resolver el curso: si no vino sheetId, autodetectar por el nombre leído del acta
  let cursoNombre = '';
  if (!sheetId) {
    const { match, candidatos } = buscarCursoPorNombre(cursoDetectado);
    if (!match) {
      return jsonResponse({
        necesitaCursoManual: true,
        cursoDetectado: cursoDetectado,
        cursosDisponibles: candidatos,
        alumnos: []
      });
    }
    sheetId = match.id;
    cursoNombre = match.name;
  } else {
    try { cursoNombre = sanitizeString(DriveApp.getFileById(sheetId).getName()); } catch(e) { cursoNombre = cursoDetectado; }
  }

  // Matchear contra la planilla de inscriptos y guardar Condición/Calificación
  let sheet;
  try {
    sheet = SpreadsheetApp.openById(sheetId).getSheets()[0];
  } catch(e) {
    return errorResponse('No se pudo abrir la planilla del curso.', 404);
  }

  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const hdr     = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const norm    = h => (h || '').toString().trim().toLowerCase();

  const cApellido = hdr.findIndex(h => norm(h).includes('apellido') && norm(h).includes('postulante'));
  const cNombre   = hdr.findIndex(h => norm(h).includes('nombre') && norm(h).includes('postulante'));
  const cDni      = hdr.findIndex(h => norm(h).includes('documento') && norm(h).includes('identidad') && !norm(h).includes('copia'));

  const COL_CONDICION = 'Condición acta';
  const COL_NOTA       = 'Calificación acta';
  let colCondicion = hdr.findIndex(h => norm(h) === norm(COL_CONDICION)) + 1;
  if (!colCondicion) { colCondicion = lastCol + 1; sheet.getRange(1, colCondicion).setValue(COL_CONDICION); }
  let colNota = hdr.findIndex(h => norm(h) === norm(COL_NOTA)) + 1;
  if (!colNota) { colNota = Math.max(colCondicion, sheet.getLastColumn()) + 1; sheet.getRange(1, colNota).setValue(COL_NOTA); }

  const filasData = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  const normDni    = s => (s || '').toString().replace(/\D/g, '');
  const normTexto  = s => (s || '').toString().trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');

  const resultado = [];
  alumnosExtraidos.forEach(al => {
    const dniAl = normDni(al.dni);
    let rowIdx = -1;

    if (dniAl && cDni >= 0) {
      rowIdx = filasData.findIndex(r => normDni(r[cDni]) === dniAl);
    }
    if (rowIdx < 0 && cApellido >= 0 && cNombre >= 0) {
      const apAl = normTexto(al.apellido), noAl = normTexto(al.nombre);
      rowIdx = filasData.findIndex(r => normTexto(r[cApellido]) === apAl && normTexto(r[cNombre]) === noAl);
    }

    const item = {
      apellido: al.apellido || '', nombre: al.nombre || '', dni: al.dni || '',
      condicion: al.condicion || '', calificacion: al.calificacion || '',
      matched: rowIdx >= 0, rowIdx: rowIdx >= 0 ? rowIdx + 2 : null
    };

    if (rowIdx >= 0) {
      sheet.getRange(rowIdx + 2, colCondicion).setValue(item.condicion);
      sheet.getRange(rowIdx + 2, colNota).setValue(item.calificacion);
    }
    resultado.push(item);
  });

  return jsonResponse({ sheetId: sheetId, curso: cursoNombre, alumnos: resultado });
}

// ─── FUNCIÓN TEMPORAL: solo para forzar la pantalla de autorización de ────────
// Drive y Docs. Se puede borrar después de usarla una vez.
function forzarAutorizacionDrive() {
  var blob = Utilities.newBlob('test', 'text/plain', 'test.txt');
  var resource = { name: 'test-autorizacion', mimeType: MimeType.GOOGLE_DOCS };
  var inserted = Drive.Files.create(resource, blob);
  Logger.log('Creado con Drive API: ' + inserted.id);

  // Esto toca específicamente el permiso que está fallando (auth/documents)
  var doc = DocumentApp.openById(inserted.id);
  doc.getBody().setText('probando permiso de Docs');
  Logger.log('DocumentApp.openById OK');

  DriveApp.getFileById(inserted.id).setTrashed(true); // lo borra apenas termina
  Logger.log('Listo, archivo de prueba borrado.');
}

function extraerDriveId(texto) {
  const s = (texto || '').trim();
  const m = s.match(/[-\w]{25,}/);
  return m ? m[0] : s;
}

function handleGetConfigConstancias() {
  const props = PropertiesService.getScriptProperties();
  const plantillaId = props.getProperty('plantilla_constancia_id') || '';
  const carpetaId    = props.getProperty('carpeta_constancias_id') || '';

  const out = { plantillaId: '', plantillaNombre: '', plantillaUrl: '', carpetaId: '', carpetaNombre: '', carpetaUrl: '' };
  if (plantillaId) {
    try {
      const f = DriveApp.getFileById(plantillaId);
      out.plantillaId = plantillaId; out.plantillaNombre = sanitizeString(f.getName()); out.plantillaUrl = f.getUrl();
    } catch(e) { /* la plantilla guardada ya no existe/es accesible */ }
  }
  if (carpetaId) {
    try {
      const c = DriveApp.getFolderById(carpetaId);
      out.carpetaId = carpetaId; out.carpetaNombre = sanitizeString(c.getName()); out.carpetaUrl = c.getUrl();
    } catch(e) { /* la carpeta guardada ya no existe/es accesible */ }
  }
  return jsonResponse(out);
}

/**
 * Recibe el .docx de la plantilla (base64), lo convierte a Google Doc
 * (requiere el servicio avanzado "Drive" habilitado en el proyecto de
 * Apps Script: Editor → Servicios → + → Google Drive API) y guarda su ID
 * en la configuración. Si el servicio avanzado no está habilitado, avisa
 * cómo activarlo en vez de fallar en silencio.
 */
function handleSubirPlantillaConstancia(base64, mimeType, nombre) {
  const MIMETYPES_DOC = [
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document', // .docx
    'application/msword', // .doc
    'application/vnd.google-apps.document'
  ];
  if (!MIMETYPES_DOC.includes(mimeType)) {
    return errorResponse('El archivo debe ser un Word (.docx) o Google Doc.', 400);
  }

  let blob;
  try {
    blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, nombre);
  } catch(e) {
    return errorResponse('No se pudo leer el archivo subido.', 400);
  }

  let docFile;
  try {
    if (mimeType === 'application/vnd.google-apps.document') {
      docFile = DriveApp.createFile(blob);
    } else if (typeof Drive !== 'undefined' && Drive.Files) {
      // Servicio avanzado de Drive (API v3): convierte el .docx a Google Doc nativo
      const resource = { name: nombre, mimeType: MimeType.GOOGLE_DOCS };
      const inserted = Drive.Files.create(resource, blob, { convert: true });
      docFile = DriveApp.getFileById(inserted.id);
    } else {
      return errorResponse('Falta habilitar el servicio avanzado "Google Drive API" en el proyecto de Apps Script (Editor → Servicios → +) para poder convertir el .docx a Google Doc.', 500);
    }
  } catch(e) {
    return errorResponse('Error convirtiendo la plantilla: ' + e.message, 500);
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperty('plantilla_constancia_id', docFile.getId());

  // Si ya hay una carpeta de Constancias configurada, mover la plantilla ahí (en vez de dejarla en la raíz del Drive)
  const carpetaId = props.getProperty('carpeta_constancias_id') || '';
  if (carpetaId) {
    try {
      const carpetaDestino = DriveApp.getFolderById(carpetaId);
      docFile.moveTo(carpetaDestino);
    } catch(e) { /* si falla, la dejamos donde se creó */ }
  }

  return jsonResponse({ ok: true, plantillaId: docFile.getId(), plantillaNombre: sanitizeString(docFile.getName()), plantillaUrl: docFile.getUrl() });
}

function handleGuardarCarpetaConstancias(carpetaTexto) {
  const carpetaId = extraerDriveId(carpetaTexto);
  let carpeta;
  try {
    carpeta = DriveApp.getFolderById(carpetaId);
  } catch(e) {
    return errorResponse('No se encontró esa carpeta en Drive. Revisá el link o el ID.', 404);
  }
  PropertiesService.getScriptProperties().setProperty('carpeta_constancias_id', carpetaId);

  // Si la plantilla ya estaba subida (en otro lado), la movemos ahora a esta carpeta
  const plantillaId = PropertiesService.getScriptProperties().getProperty('plantilla_constancia_id') || '';
  if (plantillaId) {
    try { DriveApp.getFileById(plantillaId).moveTo(carpeta); } catch(e) { /* si falla, la dejamos donde estaba */ }
  }

  return jsonResponse({ ok: true, carpetaId: carpetaId, carpetaNombre: sanitizeString(carpeta.getName()), carpetaUrl: carpeta.getUrl() });
}

function handleGetConfigComprobantes() {
  const props = PropertiesService.getScriptProperties();
  const carpetaId = props.getProperty('carpeta_comprobantes_id') || '';
  const out = { carpetaId: '', carpetaNombre: '', carpetaUrl: '' };
  if (carpetaId) {
    try {
      const c = DriveApp.getFolderById(carpetaId);
      out.carpetaId = carpetaId; out.carpetaNombre = sanitizeString(c.getName()); out.carpetaUrl = c.getUrl();
    } catch(e) { /* la carpeta guardada ya no existe/es accesible */ }
  }
  return jsonResponse(out);
}

function handleGuardarCarpetaComprobantes(carpetaTexto) {
  const carpetaId = extraerDriveId(carpetaTexto);
  let carpeta;
  try {
    carpeta = DriveApp.getFolderById(carpetaId);
  } catch(e) {
    return errorResponse('No se encontró esa carpeta en Drive. Revisá el link o el ID.', 404);
  }
  PropertiesService.getScriptProperties().setProperty('carpeta_comprobantes_id', carpetaId);
  return jsonResponse({ ok: true, carpetaId: carpetaId, carpetaNombre: sanitizeString(carpeta.getName()), carpetaUrl: carpeta.getUrl() });
}

/**
 * Guarda el PDF combinado de un lote de comprobantes (ya armado en el
 * navegador con pdf-lib) dentro de Comprobantes/<curso>/, con el nombre
 * de archivo que ya trae la fecha/hora para distinguir lotes.
 */
function handleGuardarComprobantesLote(curso, base64, nombreArchivo) {
  const carpetaId = PropertiesService.getScriptProperties().getProperty('carpeta_comprobantes_id') || '';
  if (!carpetaId) {
    return errorResponse('Todavía no configuraste la carpeta de Comprobantes. Andá a Configuración → Comprobantes.', 400);
  }

  let subcarpeta;
  try {
    const carpetaRaiz = DriveApp.getFolderById(carpetaId);
    subcarpeta = getOrCreateSubcarpeta(carpetaRaiz, curso);
  } catch(e) {
    return errorResponse('No se pudo acceder a la carpeta de Comprobantes.', 404);
  }

  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(base64), 'application/pdf', nombreArchivo);
    const file = subcarpeta.createFile(blob).setName(nombreArchivo);
    return jsonResponse({ ok: true, fileId: file.getId(), url: file.getUrl(), nombre: file.getName() });
  } catch(e) {
    return errorResponse('Error guardando el PDF: ' + e.message, 500);
  }
}

/**
 * Convierte una calificación numérica (0-10, con decimales opcionales) a su
 * forma "10 (diez)" como en el modelo de constancia.
 */
function calificacionATexto(num) {
  const UNIDADES = ['cero','uno','dos','tres','cuatro','cinco','seis','siete','ocho','nueve','diez'];
  const n = parseFloat(num);
  if (isNaN(n)) return String(num || '');
  if (Number.isInteger(n) && n >= 0 && n <= 10) return n + ' (' + UNIDADES[n] + ')';
  return String(num); // decimales: se deja el número tal cual
}

const MESES = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];

/**
 * A partir del texto libre de "Fechas de realización" (formatos muy variados:
 * "24 y 26 de febrero y 3,5,10 y 12 de marzo, de 18 a 20hrs.", "Del 20/4 al
 * 18/5. De 18 a 21hs.", "Lunes 09/03 - Martes 10/03 de 09 a 12...", etc),
 * arma un texto simple con solo el primer y el último día, tipo "9 de marzo
 * al 13 de marzo de 2026", sin horarios. Si el regex no encuentra ninguna
 * fecha (texto tipo "cronograma a confirmar"), se lo pedimos a Gemini.
 */
function resumirFechas(fechaStr) {
  if (!fechaStr) return '';
  const anioActual = new Date().getFullYear();

  // 1) Sacar franjas horarias para que "18" o "20hs" no se confundan con días (ej: "de 18 a 20hs")
  const sinHoras = fechaStr.replace(/\bde\s*\d{1,2}(:\d{2})?\s*a\s*\d{1,2}(:\d{2})?\s*(hs\.?|hrs\.?|horas)?/gi, ' ');

  // 2) Buscar patrones "D[, D, D y D] de <mes>" (con nombre de mes en texto)
  const pares = [];
  const reMes = new RegExp('((?:\\d{1,2}[\\s,y]*)+)\\s*de\\s*(' + MESES.join('|') + ')', 'gi');
  let m;
  while ((m = reMes.exec(sinHoras)) !== null) {
    const mesIdx = MESES.findIndex(mm => mm === m[2].toLowerCase());
    const dias = m[1].match(/\d{1,2}/g) || [];
    dias.forEach(d => pares.push({ dia: parseInt(d, 10), mes: mesIdx, anio: anioActual }));
  }

  // 3) Si no había nombres de mes, probar formato D/M (con o sin año)
  if (!pares.length) {
    const matches = [...fechaStr.matchAll(/(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?/g)];
    matches.forEach(mm => {
      const dd = parseInt(mm[1], 10), mesIdx = parseInt(mm[2], 10) - 1;
      let yyyy = mm[3] ? parseInt(mm[3], 10) : anioActual;
      if (yyyy < 100) yyyy += 2000;
      pares.push({ dia: dd, mes: mesIdx, anio: yyyy });
    });
  }

  if (pares.length) {
    const f1 = pares[0], f2 = pares[pares.length - 1];
    const nombreMes = i => MESES[i] || '';
    if (f1.dia === f2.dia && f1.mes === f2.mes && f1.anio === f2.anio) {
      return f1.dia + ' de ' + nombreMes(f1.mes) + ' de ' + f1.anio;
    }
    return f1.dia + ' de ' + nombreMes(f1.mes) + ' al ' + f2.dia + ' de ' + nombreMes(f2.mes) + ' de ' + f2.anio;
  }

  // 4) Último respaldo: texto sin ninguna fecha reconocible (ej: "cronograma a confirmar")
  return interpretarFechasConGemini(fechaStr) || fechaStr;
}

function interpretarFechasConGemini(fechaStr) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) return null;
  const anio = new Date().getFullYear();
  const prompt = 'Este texto describe cuándo se dicta un curso de posgrado: "' + fechaStr + '". ' +
    'Respondé SOLO con el primer día y el último día de cursada (sin mencionar horarios ni días de la semana), ' +
    'en el formato "D de MES al D de MES de ' + anio + '" (si hay un solo día de cursada: "D de MES de ' + anio + '"). ' +
    'Si el texto no menciona año, asumí ' + anio + '. ' +
    'Si el texto NO tiene fechas concretas (ej: "cronograma solicitado", "a definir", "mediados de octubre"), respondé exactamente: FECHA A CONFIRMAR. ' +
    'No agregues explicaciones, comillas ni texto adicional, solo esa frase.';
  try {
    const gemUrl = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;
    const payload = JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0 } });
    const res = UrlFetchApp.fetch(gemUrl, { method:'post', contentType:'application/json', payload:payload, muteHttpExceptions:true });
    const result = JSON.parse(res.getContentText());
    const text = (result.candidates?.[0]?.content?.parts?.[0]?.text || '').trim().replace(/^["']|["']$/g, '');
    return text || null;
  } catch(e) {
    return null;
  }
}

/** Devuelve (creándola si no existe) la subcarpeta de un curso dentro de la carpeta de constancias. */
function getOrCreateSubcarpeta(carpetaPadre, nombreCurso) {
  const existentes = carpetaPadre.getFoldersByName(nombreCurso);
  if (existentes.hasNext()) return existentes.next();
  return carpetaPadre.createFolder(nombreCurso);
}

/**
 * Genera un PDF de constancia por cada fila indicada, a partir de la
 * plantilla de Google Docs configurada en "Configuración". Arma una frase
 * distinta según la condición del alumno (Aprobado/Asistente) y reemplaza
 * los {{PLACEHOLDERS}}. Guarda cada PDF en una subcarpeta por curso, dentro
 * de la carpeta de constancias configurada, y devuelve los links.
 * Los alumnos "Ausente" nunca generan constancia (se filtran acá también
 * como resguardo, aunque el front ya no debería dejar tildarlos).
 */
function handleGenerarConstancias(sheetId, curso, filas) {
  const props = PropertiesService.getScriptProperties();
  const plantillaId = props.getProperty('plantilla_constancia_id') || '';
  const carpetaId    = props.getProperty('carpeta_constancias_id') || '';

  if (!plantillaId) {
    return errorResponse('Todavía no configuraste la plantilla de constancia. Andá a Configuración → Constancias.', 400);
  }
  if (!carpetaId) {
    return errorResponse('Todavía no configuraste la carpeta de constancias. Andá a Configuración → Constancias.', 400);
  }

  let sheet, carpetaRaiz, carpetaAnio, subcarpeta;
  try {
    sheet = SpreadsheetApp.openById(sheetId).getSheets()[0];
    carpetaRaiz = DriveApp.getFolderById(carpetaId);
    carpetaAnio = getOrCreateSubcarpeta(carpetaRaiz, new Date().getFullYear().toString());
    subcarpeta = getOrCreateSubcarpeta(carpetaAnio, curso);
  } catch(e) {
    return errorResponse('No se pudo acceder a la planilla o a la carpeta de constancias.', 404);
  }

  const lastCol = sheet.getLastColumn();
  const hdr     = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const norm    = h => (h || '').toString().trim().toLowerCase();
  const cApellido = hdr.findIndex(h => norm(h).includes('apellido') && norm(h).includes('postulante'));
  const cNombre   = hdr.findIndex(h => norm(h).includes('nombre') && norm(h).includes('postulante'));
  const cDni      = hdr.findIndex(h => norm(h).includes('documento') && norm(h).includes('identidad') && !norm(h).includes('copia'));
  const cCondicion = hdr.findIndex(h => norm(h) === norm('Condición acta'));
  const cNota       = hdr.findIndex(h => norm(h) === norm('Calificación acta'));

  // Datos del curso desde la planilla de implementación (docente, fechas, duración, tipo+denominación)
  let docente = '', fechasCrudo = '', duracion = '', cursoTexto = curso; // fallback: nombre de la planilla de Drive
  try {
    const impl = SpreadsheetApp.openById(PLANILLA_IMPLEMENTACION_ID);
    const anio = new Date().getFullYear().toString();
    const hojaImpl = impl.getSheets().find(s => s.getName().includes(anio)) || impl.getSheetByName(HOJA_IMPLEMENTACION) || impl.getSheets()[0];
    const values = hojaImpl.getDataRange().getValues();
    const hdrImpl = values[0];
    const nrm = h => (h || '').toString().trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
    const cCursoImpl = hdrImpl.findIndex(h => nrm(h).includes('denominacion'));
    const cTipoImpl  = hdrImpl.findIndex(h => nrm(h).includes('tipo') && nrm(h).includes('propuesta'));
    const cDocImpl   = hdrImpl.findIndex(h => nrm(h).includes('docente') && nrm(h).includes('responsable'));
    const cFecImpl   = hdrImpl.findIndex(h => nrm(h).includes('fecha'));
    const cCargaImpl = hdrImpl.findIndex(h => nrm(h).includes('carga') && nrm(h).includes('horaria'));
    const normCurso = s => nrm(s).replace(/^\d{4}\s*-\s*/,'').replace(/^(taller|curso|seminario)(\s+de\s+postgrado)?\s*/,'').trim();
    const qn = normCurso(curso);
    const fila = values.slice(1).find(r => normCurso(r[cCursoImpl]) === qn) ||
                 values.slice(1).find(r => { const rn = normCurso(r[cCursoImpl]); return rn && qn && (rn.includes(qn) || qn.includes(rn)); });
    if (fila) {
      docente     = cDocImpl >= 0 ? (fila[cDocImpl] || '').toString().trim() : '';
      fechasCrudo = cFecImpl >= 0 ? (fila[cFecImpl] || '').toString().trim() : '';
      duracion    = cCargaImpl >= 0 ? (fila[cCargaImpl] || '').toString().trim() : '';

      const tipo         = cTipoImpl  >= 0 ? (fila[cTipoImpl]  || '').toString().trim() : '';
      const denominacion = cCursoImpl >= 0 ? (fila[cCursoImpl] || '').toString().trim() : '';
      if (denominacion) cursoTexto = tipo ? (tipo + ' "' + denominacion + '"') : denominacion;
    }
  } catch(e) { /* seguimos con el fallback si falla */ }

  const fechas = resumirFechas(fechasCrudo);
  const hoy = new Date();
  const fechaEmision = 'los ' + hoy.getDate() + ' días del mes de ' + MESES[hoy.getMonth()] + ' del año ' + hoy.getFullYear();

  const resultado = [];
  filas.forEach(rowIdx => {
    let docId = null;
    try {
      const row = sheet.getRange(rowIdx, 1, 1, lastCol).getValues()[0];
      const apellido  = cApellido  >= 0 ? (row[cApellido]||'').toString().trim()  : '';
      const nombre    = cNombre    >= 0 ? (row[cNombre]||'').toString().trim()    : '';
      const dni       = cDni       >= 0 ? (row[cDni]||'').toString().trim()       : '';
      const nota      = cNota      >= 0 ? (row[cNota]||'').toString().trim()      : '';
      const condicion = cCondicion >= 0 ? (row[cCondicion]||'').toString().trim() : '';
      const nombreCompleto = (apellido + ', ' + nombre).replace(/^,\s*/,'').trim();

      // Resguardo: nunca generar constancia para un "Ausente", aunque haya llegado en la lista
      if (condicion === 'Ausente') {
        resultado.push({ rowIdx: rowIdx, nombre: nombreCompleto, ok: false, error: 'Alumno ausente: no corresponde constancia.' });
        return;
      }

      const fraseResultado = condicion === 'Asistente'
        ? 'ASISTIÓ al ' + cursoTexto + ' a cargo de ' + docente + ', realizado desde el ' + fechas + ', con una duración de ' + duracion + '.'
        : 'APROBÓ con calificación ' + calificacionATexto(nota) + ' el ' + cursoTexto + ' a cargo de ' + docente + ', realizado desde el ' + fechas + ', con una duración de ' + duracion + '.';

      const copia = DriveApp.getFileById(plantillaId)
        .makeCopy('Constancia - ' + nombreCompleto + ' (temp)', carpetaRaiz);
      docId = copia.getId();
      const doc = DocumentApp.openById(docId);
      const body = doc.getBody();
      body.replaceText('\\{\\{NOMBRE_COMPLETO\\}\\}', nombreCompleto);
      body.replaceText('\\{\\{DNI\\}\\}', dni);
      body.replaceText('\\{\\{FRASE_RESULTADO\\}\\}', fraseResultado);
      body.replaceText('\\{\\{FECHA_EMISION\\}\\}', fechaEmision);
      doc.saveAndClose();

      const pdfBlob = DriveApp.getFileById(docId).getAs(MimeType.PDF);
      const pdfFile = subcarpeta.createFile(pdfBlob).setName('Constancia - ' + nombreCompleto + '.pdf');

      resultado.push({ rowIdx: rowIdx, nombre: nombreCompleto, ok: true, url: pdfFile.getUrl(), fileId: pdfFile.getId() });
    } catch(e) {
      resultado.push({ rowIdx: rowIdx, ok: false, error: e.message });
    } finally {
      // Pase lo que pase (éxito o error a mitad de camino), nunca dejar el doc temporal
      if (docId) {
        try {
          if (typeof Drive !== 'undefined' && Drive.Files) { Drive.Files.remove(docId); }
          else { DriveApp.getFileById(docId).setTrashed(true); }
        } catch(e2) { /* si ni siquiera esto funciona, no bloqueamos el resultado */ }
      }
    }
  });

  return jsonResponse({ constancias: resultado });
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

    const COL_MONTO_NOMBRE  = 'Monto comprobante';
    const COL_ESTADO_NOMBRE = 'Estado comprobante';
    let deltaIngresos = 0;

    // Si se está por pisar el Monto de una fila que YA estaba marcada OK
    // (por ejemplo, al corregir un pago cargado antes), primero se resta el
    // monto viejo -antes de perderlo- para no arrastrar un valor duplicado.
    // Se vuelve a sumar el monto correcto más abajo, cuando llega el guardado
    // de "Estado comprobante".
    if (colName === COL_MONTO_NOMBRE) {
      const cEstado = hdr.findIndex(h => String(h).trim() === COL_ESTADO_NOMBRE);
      if (cEstado >= 0) {
        const estadoActual = (sheet.getRange(rowIdx + 1, cEstado + 1).getValue() || '').toString().trim();
        if (estadoActual === 'OK') {
          const montoViejo = parseFloat(sheet.getRange(rowIdx + 1, colPos).getValue());
          if (!isNaN(montoViejo) && montoViejo !== 0) deltaIngresos -= montoViejo;
        }
      }
    }

    // rowIdx es 1-based (fila 1 = primera fila de datos, fila 2 en la hoja)
    sheet.getRange(rowIdx + 1, colPos).setValue(value);

    // Si se está marcando el comprobante como verificado, sumar el monto
    // (ya escrito en la fila) al acumulador de ingresos.
    if (colName === COL_ESTADO_NOMBRE && String(value).trim() === 'OK') {
      const cMonto = hdr.findIndex(h => String(h).trim() === COL_MONTO_NOMBRE);
      if (cMonto >= 0) {
        const montoActual = parseFloat(sheet.getRange(rowIdx + 1, cMonto + 1).getValue());
        if (!isNaN(montoActual) && montoActual !== 0) deltaIngresos += montoActual;
      }
    }

    if (deltaIngresos !== 0) {
      const props = PropertiesService.getScriptProperties();
      const total = parseFloat(props.getProperty('total_ingresos_verificados') || '0') || 0;
      props.setProperty('total_ingresos_verificados', String(total + deltaIngresos));
    }

    Logger.log('saveColumna OK: hoja=' + sheet.getName() + ' fila=' + (rowIdx+1) + ' col=' + colPos + ' colName=' + colName);
    return jsonResponse({ ok: true, sheet: sheet.getName(), row: rowIdx + 1, col: colPos });
  } catch(e) {
    Logger.log('Error saveColumna: ' + e.message);
    return errorResponse('Error guardando columna: ' + e.message, 500);
  }
}

/**
 * Igual que handleSaveColumna pero guarda varias columnas de una fila de
 * una sola vez (ej: Monto/Fecha/CBU/ID Transacción/Estado comprobante,
 * usado por "Controlar pagos"). Si entre las columnas está "Estado
 * comprobante", además actualiza el total acumulado de ingresos
 * verificados guardado en PropertiesService (ver handleGetTotalIngresos),
 * calculando la diferencia contra el valor anterior de esa fila para no
 * duplicar ni perder montos si se reprocesa la misma fila.
 */
function handleSaveColumnas(sheetId, rowIdx, cols, sheetIndex) {
  const COL_ESTADO = 'Estado comprobante';
  const COL_MONTO  = 'Monto comprobante';
  try {
    const ss     = SpreadsheetApp.openById(sheetId);
    const sheets = ss.getSheets();
    let sheet = sheets[Math.min(sheetIndex || 0, sheets.length - 1)];
    if (!sheet) sheet = sheets[0];

    // Si se está tocando el estado del comprobante, medir el monto ANTES de sobrescribir
    let deltaIngresos = 0;
    if (Object.prototype.hasOwnProperty.call(cols, COL_ESTADO)) {
      const lastColPrev = sheet.getLastColumn();
      const hdrPrev = sheet.getRange(1, 1, 1, lastColPrev).getValues()[0];
      const norm = h => String(h || '').trim();
      const cEstadoPrev = hdrPrev.findIndex(h => norm(h) === COL_ESTADO);
      const cMontoPrev  = hdrPrev.findIndex(h => norm(h) === COL_MONTO);
      let prevMonto = 0;
      if (cEstadoPrev >= 0 && cMontoPrev >= 0) {
        const fila = sheet.getRange(rowIdx + 1, 1, 1, lastColPrev).getValues()[0];
        if ((fila[cEstadoPrev] || '').toString().trim() === 'OK') {
          prevMonto = parseFloat(fila[cMontoPrev]) || 0;
        }
      }
      const nuevoEstado = (cols[COL_ESTADO] || '').toString().trim();
      const nuevoMonto  = nuevoEstado === 'OK' ? (parseFloat(cols[COL_MONTO]) || 0) : 0;
      deltaIngresos = nuevoMonto - prevMonto;
    }

    // Escribir cada columna (creándola al final del encabezado si no existe)
    Object.keys(cols).forEach(colName => {
      const lastCol = sheet.getLastColumn();
      const hdr = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      let colPos = hdr.findIndex(h => String(h).trim() === colName) + 1;
      if (colPos === 0) { colPos = lastCol + 1; sheet.getRange(1, colPos).setValue(colName); }
      sheet.getRange(rowIdx + 1, colPos).setValue(cols[colName]);
    });

    if (deltaIngresos !== 0) {
      const props = PropertiesService.getScriptProperties();
      const actual = parseFloat(props.getProperty('total_ingresos_verificados') || '0') || 0;
      props.setProperty('total_ingresos_verificados', String(actual + deltaIngresos));
    }

    Logger.log('saveColumnas OK: hoja=' + sheet.getName() + ' fila=' + (rowIdx + 1) + ' cols=' + Object.keys(cols).join(',') + (deltaIngresos ? ' deltaIngresos=' + deltaIngresos : ''));
    return jsonResponse({ ok: true, sheet: sheet.getName(), row: rowIdx + 1 });
  } catch(e) {
    Logger.log('Error saveColumnas: ' + e.message);
    return errorResponse('Error guardando columnas: ' + e.message, 500);
  }
}

/** Devuelve el total de ingresos verificados acumulado (ver handleSaveColumnas). */
function handleGetTotalIngresos() {
  const props = PropertiesService.getScriptProperties();
  const total = parseFloat(props.getProperty('total_ingresos_verificados') || '0') || 0;
  return jsonResponse({ total: total });
}

/**
 * Recalcula desde cero el total de ingresos verificados, recorriendo TODAS
 * las planillas de cursos del año y sumando "Monto comprobante" (o el
 * esquema viejo "Monto verificado") de las filas verificadas. Guarda el
 * resultado como nuevo punto de partida del acumulador.
 *
 * Usar UNA SOLA VEZ luego de desplegar el acumulador (para no arrancar en
 * $0 perdiendo lo ya procesado antes de este cambio), o cuando se
 * sospeche que el acumulador quedó desincronizado. Se puede ejecutar desde
 * el editor de Apps Script (▶ Ejecutar, eligiendo esta función) o llamando
 * a la acción "recalcularTotalIngresos".
 */
function recalcularTotalIngresos() {
  const carpeta = getCarpetaInscripcionesDelAnio();
  if (!carpeta) { Logger.log('No se encontró la carpeta de inscripciones.'); return 0; }

  const files = carpeta.getFilesByType(MimeType.GOOGLE_SHEETS);
  let total = 0;

  while (files.hasNext()) {
    const file = files.next();
    try {
      const sheet = SpreadsheetApp.openById(file.getId()).getSheets()[0];
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow < 2) continue;

      const hdr = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      const exact = h => (h || '').toString().trim();
      // Nombres exactos: si una planilla tiene las dos columnas de monto a
      // la vez (quedó "a medio migrar" del esquema viejo al nuevo), cada
      // fila se decide por separado para no perder ni duplicar ninguna.
      const montoNuevoIdx = hdr.findIndex(h => exact(h) === 'Monto comprobante');
      const estadoIdx     = hdr.findIndex(h => exact(h) === 'Estado comprobante');
      const montoViejoIdx = hdr.findIndex(h => exact(h) === 'Monto verificado');
      if (montoNuevoIdx < 0 && montoViejoIdx < 0) continue;

      const rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
      rows.forEach(row => {
        let monto = NaN;
        // 1) Esquema nuevo: solo cuenta si ESTA fila puntual está OK
        if (estadoIdx >= 0 && montoNuevoIdx >= 0 && (row[estadoIdx] || '').toString().trim() === 'OK') {
          monto = parseFloat((row[montoNuevoIdx] || '').toString().trim());
        }
        // 2) Si no se contó por el nuevo, esquema viejo como respaldo (solo para esta fila)
        if (isNaN(monto) && montoViejoIdx >= 0) {
          const raw = (row[montoViejoIdx] || '').toString().trim();
          if (raw) monto = parseFloat(raw);
        }
        if (!isNaN(monto)) total += monto;
      });
    } catch(e) {
      Logger.log('Error recalculando ' + file.getName() + ': ' + e.message);
    }
  }

  PropertiesService.getScriptProperties().setProperty('total_ingresos_verificados', String(total));
  Logger.log('Total de ingresos recalculado: $' + total);
  return total;
}

function handleRecalcularTotalIngresos() {
  const total = recalcularTotalIngresos();
  return jsonResponse({ ok: true, total: total });
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

// ─── AUTENTICACIÓN ───────────────────────────────────────────────────────────
// Usuarios almacenados en Script Properties con clave 'USUARIOS_APP'.
// Formato JSON: [{"name":"Nombre Apellido","email":"usuario@unpa.edu.ar","hash":"<sha256_de_la_contrasena>"}]
// Para generar el hash de una contraseña, ejecutar en el editor del script:
//   Logger.log(hashPassword('mi_contrasena'))

function handleLogin(email, pwd) {
  try {
    var props     = PropertiesService.getScriptProperties();
    var usersJson = props.getProperty('USUARIOS_APP');
    if (!usersJson) {
      return ContentService.createTextOutput(
        JSON.stringify({ ok: false, msg: 'Sin usuarios configurados. Contactar al administrador.' })
      ).setMimeType(ContentService.MimeType.JSON);
    }
    var usuarios;
    try {
      usuarios = JSON.parse(usersJson);
    } catch (err) {
      return ContentService.createTextOutput(
        JSON.stringify({ ok: false, msg: 'La configuración de usuarios es inválida. Contactar al administrador.' })
      ).setMimeType(ContentService.MimeType.JSON);
    }
    var hash     = hashPassword(pwd);
    var found    = null;
    for (var i = 0; i < usuarios.length; i++) {
      var storedEmail = String(usuarios[i].email || '').trim().toLowerCase();
      if (storedEmail === String(email || '').trim().toLowerCase() && String(usuarios[i].hash || '') === hash) {
        found = usuarios[i];
        break;
      }
    }
    if (!found) {
      return ContentService.createTextOutput(
        JSON.stringify({ ok: false, msg: 'Email o contraseña incorrectos.' })
      ).setMimeType(ContentService.MimeType.JSON);
    }
    var sessionToken = Utilities.getUuid();
    CacheService.getScriptCache().put('sess_' + sessionToken, '1', 21600); // 6 horas
    return ContentService.createTextOutput(
      JSON.stringify({ ok: true, name: found.name, email: found.email, token: sessionToken })
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log('Error en handleLogin: ' + err.message);
    return ContentService.createTextOutput(
      JSON.stringify({ ok: false, msg: 'Error al validar credenciales.' })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function hashPassword(pwd) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pwd, Utilities.Charset.UTF_8);
  return bytes.map(function(b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}

function verificarToken(token) {
  if (!token || token.length < 10) return false;
  return CacheService.getScriptCache().get('sess_' + token) !== null;
}
