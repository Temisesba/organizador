// ============================================================
//  ORGANIZADOR — Backend de Google Apps Script
//  Cada usuario vincula su propio Google Sheet (URL pedida al abrir
//  la app por primera vez) — el sheetId viaja en cada petición y se
//  abre con SpreadsheetApp.openById(sheetId). Nadie sin ese enlace
//  puede ver o tocar los datos.
//  IA: Gemini (principal) + Claude (respaldo y OCR de imágenes)
// ============================================================

// Ejecuta esta función UNA VEZ desde el editor (▶ Run): de paso autoriza los
// permisos (Sheets/Drive/red) — si nunca se autorizan, la Web App responde
// 403 a cualquiera, incluso con "acceso: cualquiera", porque Google nunca
// le pidió permiso a la cuenta dueña del script — y crea tu Google Sheet
// de Organizador ya con las pestañas listas. El enlace queda en el registro
// de ejecución (Ver → Registros / el panel que aparece abajo al terminar).
function crearMiSheet() {
  const ss = SpreadsheetApp.create('Organizador — datos');
  Object.keys(TABLE_DEFS).forEach(t => getTable(ss, t));
  Logger.log('Tu Google Sheet: ' + ss.getUrl());
  return ss.getUrl();
}

// Llaves guardadas en Configuración del proyecto → Propiedades del script
// (Project Settings → Script Properties), NO aquí — este archivo se sube
// a un repo de GitHub público y no debe llevar llaves reales.
const GEMINI_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
const GEMINI_API = 'https://generativelanguage.googleapis.com/v1/models/gemini-2.5-flash:generateContent?key=' + GEMINI_KEY;
const CLAUDE_KEY = PropertiesService.getScriptProperties().getProperty('CLAUDE_KEY');
const CLAUDE_API = 'https://api.anthropic.com/v1/messages';
const CLAUDE_MODEL = 'claude-haiku-4-5-20251001';
// El Client ID de Google NO es secreto (ver comentario igual en el
// frontend) — el Client Secret sí, por eso vive en Script Properties
// como las demás llaves. Necesario para intercambiar/refrescar tokens
// de Google Calendar sin que el usuario tenga que reconectar cada ~1h:
// el navegador nunca ve el secret, solo manda el "code" o el
// "refresh_token" aquí y este backend hace el intercambio real con
// Google por él.
const GOOGLE_CLIENT_ID = '596591883908-8qo7c1i44e5lv5p9o4nan1fkn20qflq6.apps.googleusercontent.com';
const GOOGLE_CLIENT_SECRET = PropertiesService.getScriptProperties().getProperty('GOOGLE_CLIENT_SECRET');

const TABLE_DEFS = {
  NOTAS:         ['ID', 'Fecha', 'FechaEdicion', 'Titulo', 'Contenido', 'Tags', 'Notebook', 'Cornell', 'CornellCue', 'Color', 'Pinned', 'Orden', 'Archivo', 'Papelera', 'GCalEventId', 'PlantillaPropID', 'CamposGCalJSON'],
  TAGS:          ['ID', 'Etiqueta', 'Color', 'EsComando', 'Descripcion'],
  NOTEBOOKS:     ['ID', 'Nombre', 'TagsHeredados', 'Color', 'TagAuto', 'Password'],
  PLANTILLAS:    ['ID', 'Nombre', 'ItemsJSON', 'Tags', 'Recordatorio', 'ColumnasJSON', 'Papelera'],
  HISTORIAL:     ['ID', 'PlantillaID', 'Fecha', 'ItemsJSON', 'NotaID', 'Nombre'],
  DICCIONARIO:   ['ID', 'Termino', 'Traduccion', 'Notas', 'Fecha', 'Favorito'],
  FEEDBACK:      ['ID', 'Fecha', 'Original', 'Correccion', 'Contexto'],
  BLOQUE_TABLAS: ['ID', 'Nombre', 'TagFiltro', 'CampoFiltro', 'ValorFiltro', 'Columnas', 'Carpeta'],
  PLANTILLAS_PROP: ['ID', 'Nombre', 'CamposJSON', 'UsaPersona', 'Recordatorio', 'Notebook', 'TiposJSON'],
  FLAGS:         ['ID', 'Nombre', 'Emoji', 'Color'],
  BUSQUEDAS:     ['ID', 'Nombre', 'Query', 'Tags', 'Modo'],
  PERSONAS:      ['ID', 'Nombre', 'Color', 'FechaNacimiento', 'GCalEventId'],
  HABITOS:       ['ID', 'Nombre', 'Emoji', 'Color', 'Activo', 'Hora', 'Recordatorio', 'Meta', 'HorasJSON', 'Dosis', 'FechaInicio', 'FechaFin', 'Observaciones'],
  HABITO_LOG:    ['ID', 'HabitoID', 'Fecha'],
  COMANDOS:      ['ID', 'Nombre', 'Tipo', 'Valor'],
  HORARIO:       ['ID', 'Hora', 'Actividad', 'Dias'],
  HORARIO_LOG:   ['ID', 'HorarioID', 'Fecha']
};

// ── Router GET (solo lectura ligera, ej. ping de conexión) ────
function doGet(e) {
  const action = (e.parameter.action || '').toString();
  let result;
  try {
    if (action === 'ping') result = { ok: true, hora: new Date().toISOString() };
    else result = { error: 'Acción no reconocida' };
  } catch (err) { result = { error: err.message }; }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// ── Router POST ─────────────────────────────────────────────
function doPost(e) {
  let body;
  try {
    const raw = e.parameter.payload || (e.postData ? e.postData.contents : '{}');
    body = JSON.parse(raw);
  } catch (err) { body = {}; }

  const action = body.action || '';
  let result;
  try {
    if (action !== 'ping' && !body.sheetId) throw new Error('Falta sheetId — vincula tu Google Sheet primero');

    if      (action === 'ping')                 result = { ok: true };
    else if (action === 'init')                 result = initSheet(body.sheetId);
    else if (action === 'getAll')                result = getAllData(body.sheetId);
    else if (action === 'saveRow')              result = saveRow(body.sheetId, body.table, body.row);
    else if (action === 'deleteRow')            result = deleteRow(body.sheetId, body.table, body.id);
    else if (action === 'fillTemplate')         result = fillTemplate(body.sheetId, body.plantillaId);
    else if (action === 'toggleHistorialItem')  result = toggleHistorialItem(body.sheetId, body.historialId, body.index);
    else if (action === 'updateHistorialItems') result = updateHistorialItems(body.sheetId, body.historialId, body.items);
    else if (action === 'aiAssist')             result = aiAssist(body.mode, body.text, body.instruction);
    else if (action === 'translate')            result = translate(body.text, body.direction);
    else if (action === 'aiEnglishLesson')      result = aiEnglishLesson(body.text);
    else if (action === 'ocrImage')             result = ocrImage(body.imageBase64, body.mimeType);
    else if (action === 'uploadImage')          result = uploadImage(body.sheetId, body.imageBase64, body.mimeType, body.fileName);
    else if (action === 'exchangeGoogleCode')   result = exchangeGoogleCode(body.code);
    else if (action === 'refreshGoogleToken')   result = refreshGoogleToken(body.refresh_token);
    else                                        result = { error: 'Acción no reconocida' };
  } catch (err) { result = { error: err.message }; }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// ── Helpers de hoja genéricos ──────────────────────────────────
function getSS(sheetId) {
  return SpreadsheetApp.openById(sheetId);
}

// Columnas que GUARDAN texto con pinta de fecha/hora ("22:00") pero que
// deben tratarse SIEMPRE como texto literal, nunca como celda de
// Fecha/Hora real de Sheets — de lo contrario Sheets auto-detecta el
// valor al escribirlo y lo convierte a su propio tipo Hora (un número
// serie interno). Sheets.getValues() entonces devuelve un objeto Date
// de Apps Script (anclado a 1899-12-30/31 en la zona horaria DEL
// SPREADSHEET), que al viajar a JSON se serializa como ISO en UTC (ej.
// "1899-12-31T01:36:36.000Z") — si el navegador de quien lee está en
// OTRA zona horaria que la del spreadsheet, la hora reconstruida sale
// corrida (bug real reportado: "puse 22:00 y sale 13:36"). Forzar
// formato de texto ("@") en estas columnas evita que Sheets convierta
// el valor en primer lugar — nunca se guarda como Hora-de-Sheets, así
// que nunca hay nada que reconstruir/adivinar en zona horaria alguna.
const TEXT_FORMAT_COLUMNS = { HABITOS: ['Hora', 'HorasJSON'], HORARIO: ['Hora'] };
function getTable(ss, tableName) {
  const headers = TABLE_DEFS[tableName];
  if (!headers) throw new Error('Tabla desconocida: ' + tableName);
  let sh = ss.getSheetByName(tableName);
  if (!sh) {
    sh = ss.insertSheet(tableName);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  } else {
    const lastCol = Math.max(sh.getLastColumn(), 1);
    const existing = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    headers.forEach((h, i) => { if (existing[i] !== h) sh.getRange(1, i + 1).setValue(h); });
  }
  const textCols = TEXT_FORMAT_COLUMNS[tableName];
  if (textCols) {
    textCols.forEach(colName => {
      const colIdx = headers.indexOf(colName) + 1;
      if (colIdx > 0) sh.getRange(1, colIdx, sh.getMaxRows(), 1).setNumberFormat('@');
    });
  }
  return sh;
}

function sheetToObjects(sh, headers, tableName) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  // Forzar la columna a texto (ver TEXT_FORMAT_COLUMNS en getTable) evita
  // que Sheets vuelva a convertir un valor NUEVO a su tipo Hora interno —
  // pero un valor que YA se guardó ANTES de aplicar ese formato sigue
  // siendo, por dentro, una celda de tipo Hora real hasta que se
  // reescriba; getValues() la sigue devolviendo como Date de Apps
  // Script mientras tanto. Acá se reconstruye el texto correcto para
  // esos casos ya existentes, usando la zona horaria DEL SPREADSHEET
  // (Utilities.formatDate con ese tz es exactamente el inverso de cómo
  // Sheets/Apps Script generó ese Date en primer lugar) — no la del
  // navegador de quien esté leyendo, que es lo que causaba la hora
  // corrida en el bug real reportado.
  const textCols = TEXT_FORMAT_COLUMNS[tableName];
  if (textCols) {
    const tz = sh.getParent().getSpreadsheetTimeZone();
    const colIdxs = textCols.map(c => headers.indexOf(c)).filter(i => i >= 0);
    values.forEach(r => {
      colIdxs.forEach(i => { if (r[i] instanceof Date) r[i] = Utilities.formatDate(r[i], tz, 'HH:mm'); });
    });
  }
  return values
    .filter(r => r.some(c => c !== ''))
    .map(r => {
      const o = {};
      headers.forEach((h, i) => { o[h] = r[i]; });
      return o;
    });
}

function findRowIndexById(sh, id) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function initSheet(sheetId) {
  const ss = getSS(sheetId);
  Object.keys(TABLE_DEFS).forEach(t => getTable(ss, t));
  return getAllData(sheetId);
}

function getAllData(sheetId) {
  const ss = getSS(sheetId);
  const out = {};
  Object.keys(TABLE_DEFS).forEach(t => {
    const sh = getTable(ss, t);
    out[t.toLowerCase()] = sheetToObjects(sh, TABLE_DEFS[t], t);
  });
  return { ok: true, data: out };
}

// "Comprobar si ya existe la fila" y "agregarla/actualizarla" NO son un
// solo paso atómico — si dos peticiones de guardar la MISMA nota nueva se
// procesan casi al mismo tiempo (dos guardados casi simultáneos desde el
// navegador, ej. un blur y un autoguardado que caen juntos), las dos
// pueden ver "todavía no existe" y las dos hacer appendRow, dejando DOS
// filas con el mismo ID — justo el bug reportado ("se veía doble, borré
// una y se borró la otra"). LockService serializa esta sección para que
// solo una ejecución esté adentro a la vez.
function saveRow(sheetId, table, row) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = getSS(sheetId);
    const headers = TABLE_DEFS[table];
    if (!headers) throw new Error('Tabla desconocida: ' + table);
    const sh = getTable(ss, table);
    if (!row.ID) row.ID = Utilities.getUuid();
    const rowIdx = findRowIndexById(sh, row.ID);
    const values = headers.map(h => (row[h] !== undefined && row[h] !== null) ? row[h] : '');
    if (rowIdx > 0) sh.getRange(rowIdx, 1, 1, headers.length).setValues([values]);
    else sh.appendRow(values);
    return { ok: true, row: row };
  } finally {
    lock.releaseLock();
  }
}

function deleteRow(sheetId, table, id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = getSS(sheetId);
    const sh = getTable(ss, table);
    const rowIdx = findRowIndexById(sh, id);
    if (rowIdx > 0) sh.deleteRow(rowIdx);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// ── Checklists: llenar plantilla = nueva fila de historial fechada ──
function fillTemplate(sheetId, plantillaId) {
  const ss = getSS(sheetId);
  const shP = getTable(ss, 'PLANTILLAS');
  const plantillas = sheetToObjects(shP, TABLE_DEFS.PLANTILLAS);
  const plantilla = plantillas.find(p => String(p.ID) === String(plantillaId));
  if (!plantilla) throw new Error('Plantilla no encontrada');
  let items;
  try { items = JSON.parse(plantilla.ItemsJSON || '[]'); } catch (e) { items = []; }
  const nuevosItems = items.map(it => ({ text: (it.text !== undefined ? it.text : it), checked: false }));
  const historial = {
    ID: Utilities.getUuid(),
    PlantillaID: plantillaId,
    Fecha: new Date().toISOString(),
    ItemsJSON: JSON.stringify(nuevosItems),
    NotaID: ''
  };
  const shH = getTable(ss, 'HISTORIAL');
  shH.appendRow(TABLE_DEFS.HISTORIAL.map(h => historial[h] !== undefined ? historial[h] : ''));
  return { ok: true, row: historial };
}

function toggleHistorialItem(sheetId, historialId, index) {
  const ss = getSS(sheetId);
  const sh = getTable(ss, 'HISTORIAL');
  const rowIdx = findRowIndexById(sh, historialId);
  if (rowIdx < 0) throw new Error('Registro de historial no encontrado');
  const colItems = TABLE_DEFS.HISTORIAL.indexOf('ItemsJSON') + 1;
  const raw = sh.getRange(rowIdx, colItems).getValue();
  let items;
  try { items = JSON.parse(raw || '[]'); } catch (e) { items = []; }
  if (items[index]) items[index].checked = !items[index].checked;
  sh.getRange(rowIdx, colItems).setValue(JSON.stringify(items));
  return { ok: true, items: items };
}

// Reemplazo completo de los items de un registro de historial — se usa al
// editar notas por item, agregar items sueltos al llenar, o marcar/desmarcar,
// todo en una sola operación en vez de una acción distinta por cada cosa.
function updateHistorialItems(sheetId, historialId, items) {
  const ss = getSS(sheetId);
  const sh = getTable(ss, 'HISTORIAL');
  const rowIdx = findRowIndexById(sh, historialId);
  if (rowIdx < 0) throw new Error('Registro de historial no encontrado');
  const colItems = TABLE_DEFS.HISTORIAL.indexOf('ItemsJSON') + 1;
  sh.getRange(rowIdx, colItems).setValue(JSON.stringify(items || []));
  return { ok: true, items: items };
}

// ── IA: Gemini principal + Claude de respaldo ──────────────────
function isGeminiQuotaError(data) {
  if (!data || !data.error) return false;
  const code = data.error.code;
  const msg = (data.error.message || '').toLowerCase();
  return code === 429 || msg.includes('quota') || msg.includes('rate limit') || msg.includes('resource_exhausted');
}

function callGemini(systemPrompt, userText) {
  const payload = { contents: [{ role: 'user', parts: [{ text: systemPrompt + '\n\n' + userText }] }] };
  const response = UrlFetchApp.fetch(GEMINI_API, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true });
  const data = JSON.parse(response.getContentText());
  if (response.getResponseCode() === 429 || isGeminiQuotaError(data) || data.error) return null;
  if (!data.candidates || !data.candidates[0]) return null;
  return data.candidates[0].content.parts[0].text.trim();
}

function callClaude(systemPrompt, userText) {
  const payload = {
    model: CLAUDE_MODEL, max_tokens: 2048,
    system: systemPrompt,
    messages: [{ role: 'user', content: userText }]
  };
  const response = UrlFetchApp.fetch(CLAUDE_API, {
    method: 'post', contentType: 'application/json',
    headers: { 'x-api-key': CLAUDE_KEY, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });
  const data = JSON.parse(response.getContentText());
  if (data.error) throw new Error('Claude: ' + (data.error.message || JSON.stringify(data.error)));
  if (!data.content || !data.content[0] || !data.content[0].text) throw new Error('Claude: respuesta inesperada');
  return data.content[0].text.trim();
}

function askAI(systemPrompt, userText) {
  try {
    const g = callGemini(systemPrompt, userText);
    if (g !== null) return { ok: true, text: g, modelo: 'gemini' };
  } catch (e) {}
  try {
    const c = callClaude(systemPrompt, userText);
    return { ok: true, text: c, modelo: 'claude' };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function aiAssist(mode, text, instruction) {
  if (!text) return { ok: false, error: 'Sin texto' };
  const prompts = {
    organizar: 'Eres un asistente de redacción. Reordena y limpia el siguiente texto (párrafos, puntuación, claridad) SIN cambiar el sentido ni el idioma. Devuelve únicamente el texto resultante, sin explicaciones ni comentarios.',
    redaccion: 'Eres un corrector de redacción. Mejora la redacción, ortografía y gramática del siguiente texto EN EL MISMO IDIOMA en que está escrito (no lo traduzcas). No cambies el sentido. Devuelve únicamente el texto corregido, sin explicaciones.',
    tesis: 'Eres un asistente académico. A partir del siguiente texto, propón una estructura de tesis/argumento (introducción, puntos clave, conclusión) usando Markdown con encabezados y viñetas. Devuelve únicamente el resultado.',
    procesar: 'Eres un asistente que resume y extrae los puntos clave del siguiente texto en una lista de viñetas Markdown. Devuelve únicamente el resultado.'
  };
  // "personalizada": el usuario escribe su propia instrucción en vez de
  // elegir uno de los modos fijos de arriba — se usa tal cual como
  // prompt del sistema.
  const systemPrompt = (mode === 'personalizada' && instruction) ? instruction : (prompts[mode] || prompts.organizar);
  return askAI(systemPrompt, text);
}

function translate(text, direction) {
  if (!text) return { ok: false, error: 'Sin texto' };
  // "auto" (un solo botón "Traducir ES↔EN", pedido explícito en vez de
  // dos botones ES→EN/EN→ES separados): el propio modelo detecta en qué
  // de los dos idiomas está el texto y traduce al otro — no hace falta
  // un paso de detección de idioma aparte, un LLM ya lo hace bien con
  // una instrucción clara.
  let dir;
  if (direction === 'en->es') dir = 'del inglés al español';
  else if (direction === 'es->en') dir = 'del español al inglés';
  else dir = 'detectando tú mismo si está en español o en inglés, y tradúcelo al OTRO idioma (si está en español, tradúcelo a inglés; si está en inglés, tradúcelo a español)';
  const systemPrompt = 'Traduce el siguiente texto ' + dir + '. Devuelve ÚNICAMENTE la traducción, sin explicaciones ni comillas.';
  return askAI(systemPrompt, text);
}

// Feedback de inglés: revisa gramática/ortografía en un texto en inglés y
// devuelve una "lección" breve en español, en un formato fijo que el
// frontend guarda como nota en el cuaderno "Inglés".
function aiEnglishLesson(text) {
  if (!text) return { ok: false, error: 'Sin texto' };
  const systemPrompt = 'Eres un profesor de inglés. Analiza el siguiente texto en inglés (puede tener errores) y da retroalimentación breve EN ESPAÑOL, en este formato exacto:\n' +
    '❌ Escribiste: <el texto original o un fragmento representativo, máximo 200 caracteres>\n' +
    '📘 Lección: <lista numerada breve de los errores encontrados y cómo corregirlos — gramática, ortografía, preposiciones, tiempos verbales, etc.>\n' +
    'Si el texto no tiene errores relevantes, dilo brevemente en la Lección. Devuelve ÚNICAMENTE ese formato de dos líneas, sin texto adicional.';
  return askAI(systemPrompt, text);
}

// ── Imágenes: subir a Drive (para incrustarlas en la nota) ──────
function getOrCreateImageFolder() {
  const props = PropertiesService.getScriptProperties();
  let folderId = props.getProperty('ORGANIZADOR_IMG_FOLDER_ID');
  if (folderId) { try { return DriveApp.getFolderById(folderId); } catch (e) {} }
  const folder = DriveApp.createFolder('Organizador — imágenes');
  props.setProperty('ORGANIZADOR_IMG_FOLDER_ID', folder.getId());
  return folder;
}

function uploadImage(sheetId, imageBase64, mimeType, fileName) {
  if (!imageBase64 || !mimeType) return { ok: false, error: 'Faltan datos de la imagen' };
  try {
    const folder = getOrCreateImageFolder();
    const bytes = Utilities.base64Decode(imageBase64);
    const blob = Utilities.newBlob(bytes, mimeType, fileName || ('imagen-' + new Date().getTime()));
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1200';
    return { ok: true, url: url, fileId: file.getId() };
  } catch (e) { return { ok: false, error: 'Drive: ' + e.message }; }
}

// ── Google Calendar: intercambio/refresco de tokens ──────────────
// El frontend obtiene un "code" de un solo uso vía Google Identity
// Services (popup, sin salir de la app — ux_mode:'popup' usa
// internamente redirect_uri:"postmessage", que NO necesita registrarse
// como URI de redirección en la consola de Google, ya que no es una URL
// real). Este intercambio SÍ necesita el client secret, por eso pasa
// por aquí y no por el navegador directamente.
function exchangeGoogleCode(code) {
  if (!code) return { ok: false, error: 'Falta el código de autorización' };
  if (!GOOGLE_CLIENT_SECRET) return { ok: false, error: 'Falta GOOGLE_CLIENT_SECRET en Propiedades del script' };
  const resp = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    payload: {
      code: code,
      client_id: GOOGLE_CLIENT_ID,
      client_secret: GOOGLE_CLIENT_SECRET,
      redirect_uri: 'postmessage',
      grant_type: 'authorization_code'
    },
    muteHttpExceptions: true
  });
  const data = JSON.parse(resp.getContentText());
  if (data.error) return { ok: false, error: data.error_description || data.error };
  return { ok: true, access_token: data.access_token, refresh_token: data.refresh_token || '', expires_in: data.expires_in };
}
// El refresh_token no expira solo (dura hasta que el usuario revoque el
// acceso desde su cuenta de Google, o Google lo invalide por
// inactividad prolongada) — mientras siga siendo válido, esto renueva
// el access_token (~1h) sin volver a pedirle nada a la persona.
function refreshGoogleToken(refreshToken) {
  if (!refreshToken) return { ok: false, error: 'Falta el refresh token' };
  if (!GOOGLE_CLIENT_SECRET) return { ok: false, error: 'Falta GOOGLE_CLIENT_SECRET en Propiedades del script' };
  const resp = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    payload: {
      refresh_token: refreshToken,
      client_id: GOOGLE_CLIENT_ID,
      client_secret: GOOGLE_CLIENT_SECRET,
      grant_type: 'refresh_token'
    },
    muteHttpExceptions: true
  });
  const data = JSON.parse(resp.getContentText());
  if (data.error) return { ok: false, error: data.error_description || data.error };
  return { ok: true, access_token: data.access_token, expires_in: data.expires_in };
}

// ── OCR con Claude Visión ───────────────────────────────────────
function ocrImage(imageBase64, mimeType) {
  if (!imageBase64 || !mimeType) return { ok: false, error: 'Faltan datos de la imagen' };
  try {
    const payload = {
      model: CLAUDE_MODEL, max_tokens: 2048,
      messages: [{
        role: 'user', content: [
          { type: 'image', source: { type: 'base64', media_type: mimeType, data: imageBase64 } },
          { type: 'text', text: 'Transcribe ÚNICAMENTE el texto que aparece en esta imagen, tal cual está escrito, conservando saltos de línea. Si no hay texto legible, responde solo: (sin texto detectado)' }
        ]
      }]
    };
    const response = UrlFetchApp.fetch(CLAUDE_API, {
      method: 'post', contentType: 'application/json',
      headers: { 'x-api-key': CLAUDE_KEY, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    const data = JSON.parse(response.getContentText());
    if (data.error) return { ok: false, error: 'Claude: ' + (data.error.message || JSON.stringify(data.error)) };
    if (!data.content || !data.content[0] || !data.content[0].text) return { ok: false, error: 'Respuesta inesperada de Claude' };
    return { ok: true, text: data.content[0].text.trim() };
  } catch (e) { return { ok: false, error: 'Error al leer la imagen: ' + e.message }; }
}
