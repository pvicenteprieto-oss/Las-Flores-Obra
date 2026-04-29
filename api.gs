// ============================================================
//  CONTROL DE OBRA LAS FLORES — Apps Script API v1
//  Ecopetro SA
//
//  INSTRUCCIONES:
//  1. Abrí tu Google Sheet
//  2. Extensiones > Apps Script
//  3. Pegá este código completo
//  4. Implementar > Nueva implementación
//     - Tipo: Aplicación web
//     - Ejecutar como: Yo
//     - Acceso: Cualquier persona
//  5. Copiá la URL generada y pegala en index.html donde dice TU_URL_DE_APPS_SCRIPT_AQUI
//
//  Hojas que se crean automáticamente:
//  usuarios, personal, equipos, cargas_diarias, precios, sesiones
// ============================================================

function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const p = e.parameter;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    initSheets(ss);
    let result;
    switch(p.action) {
      case 'getAll':            result = getAllData(ss); break;
      // Personal
      case 'addPersonal':       result = addRow(ss, 'personal', JSON.parse(p.data)); break;
      case 'updatePersonal':    result = updateRow(ss, 'personal', JSON.parse(p.data)); break;
      case 'deletePersonal':    result = deleteRow(ss, 'personal', JSON.parse(p.data).id); break;
      // Equipos
      case 'addEquipo':         result = addRow(ss, 'equipos', JSON.parse(p.data)); break;
      case 'updateEquipo':      result = updateRow(ss, 'equipos', JSON.parse(p.data)); break;
      case 'deleteEquipo':      result = deleteRow(ss, 'equipos', JSON.parse(p.data).id); break;
      // Cargas diarias
      case 'saveCargaDiaria':   result = saveCargaDiaria(ss, JSON.parse(p.data)); break;
      // Precios
      case 'savePrecios':       result = savePrecios(ss, JSON.parse(p.data)); break;
      // Usuarios
      case 'addUsuario':        result = addRow(ss, 'usuarios', JSON.parse(p.data)); break;
      case 'updateUsuario':     result = updateRow(ss, 'usuarios', JSON.parse(p.data)); break;
      case 'deleteUsuario':     result = deleteRow(ss, 'usuarios', JSON.parse(p.data).id); break;
      // Sesiones
      case 'addSesion':         result = addRow(ss, 'sesiones', JSON.parse(p.data)); break;

      default: result = { error: 'Acción desconocida: ' + p.action };
    }
    return out(result);
  } catch(err) {
    return out({ error: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

function out(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── INIT SHEETS ───────────────────────────────────────────────────
// Crea las hojas con sus cabeceras si no existen todavía
function initSheets(ss) {
  const SCHEMAS = {
    usuarios:       ['id','username','nombre','rol','pin'],
    personal:       ['id','nombre','categoria','activo'],
    equipos:        ['id','codigo','nombre','unidad','tipo','obs','activo'],
    cargas_diarias: ['id','fecha','asistencia','ausencia_motivos','equipos_qty','usuario','timestamp'],
    precios:        ['id','mes','codigo','precio'],
    sesiones:       ['id','usuario','nombre','fecha','dispositivo'],
  };
  Object.entries(SCHEMAS).forEach(([name, headers]) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1a3a6e').setFontColor('#ffffff');
    }
  });
}

// ── GET ALL ───────────────────────────────────────────────────────
function getAllData(ss) {
  return {
    usuarios:       sheetToObjects(ss, 'usuarios'),
    personal:       sheetToObjects(ss, 'personal'),
    equipos:        sheetToObjects(ss, 'equipos'),
    cargas_diarias: sheetToObjects(ss, 'cargas_diarias'),
    precios:        sheetToObjects(ss, 'precios'),
    sesiones:       sheetToObjects(ss, 'sesiones'),
  };
}

// ── CRUD GENÉRICO ─────────────────────────────────────────────────
function sheetToObjects(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function addRow(ss, name, obj) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return { error: 'Hoja no encontrada: ' + name };
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  sheet.appendRow(headers.map(h => obj[h] !== undefined ? obj[h] : ''));
  return { ok: true };
}

function updateRow(ss, name, obj) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return { error: 'Hoja no encontrada: ' + name };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');
  if (idCol === -1) return { error: 'Columna id no encontrada en ' + name };
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(obj.id)) {
      sheet.getRange(i + 1, 1, 1, headers.length)
        .setValues([headers.map(h => obj[h] !== undefined ? obj[h] : data[i][headers.indexOf(h)])]);
      return { ok: true };
    }
  }
  // Si no existe, lo agrega
  return addRow(ss, name, obj);
}

function deleteRow(ss, name, id) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return { error: 'Hoja no encontrada' };
  const data = sheet.getDataRange().getValues();
  const idCol = data[0].indexOf('id');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { error: 'Registro no encontrado' };
}

// ── CARGA DIARIA ──────────────────────────────────────────────────
// Upsert: si ya existe una carga para esa fecha, la reemplaza
function saveCargaDiaria(ss, obj) {
  const sheet = ss.getSheetByName('cargas_diarias');
  if (!sheet) return { error: 'Hoja cargas_diarias no encontrada' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const fechaCol = headers.indexOf('fecha');
  // Buscar si ya existe una fila con esa fecha
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][fechaCol]) === String(obj.fecha)) {
      // Reemplazar la fila existente
      sheet.getRange(i + 1, 1, 1, headers.length)
        .setValues([headers.map(h => obj[h] !== undefined ? obj[h] : data[i][headers.indexOf(h)])]);
      return { ok: true, action: 'updated' };
    }
  }
  // No existe: agregar nueva fila
  sheet.appendRow(headers.map(h => obj[h] !== undefined ? obj[h] : ''));
  return { ok: true, action: 'inserted' };
}

// ── PRECIOS ───────────────────────────────────────────────────────
// Recibe { mes, rows: [{id, mes, codigo, precio}] }
// Borra todas las filas del mes y las reescribe
function savePrecios(ss, payload) {
  const sheet = ss.getSheetByName('precios');
  if (!sheet) return { error: 'Hoja precios no encontrada' };
  const mes = payload.mes;
  const rows = payload.rows || [];

  // Leer datos existentes
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    // Solo hay cabecera, agregar todo directo
    rows.forEach(r => sheet.appendRow([r.id, r.mes, r.codigo, r.precio]));
    return { ok: true };
  }

  const headers = data[0];
  const mesCol = headers.indexOf('mes');

  // Eliminar de abajo hacia arriba las filas del mes indicado
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][mesCol]) === String(mes)) {
      sheet.deleteRow(i + 1);
    }
  }

  // Agregar las nuevas filas
  rows.forEach(r => {
    sheet.appendRow(headers.map(h => r[h] !== undefined ? r[h] : ''));
  });

  return { ok: true };
}
