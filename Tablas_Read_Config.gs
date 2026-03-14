/**
 * CONFIGURACIÓN GLOBAL
 */ 
const DATA_SS_MAP = {
  'ULTRA_DHO': '1GEr1V2EzAm1vpGNTmKHMGYGuv3mM2F8eH-kbot-1qX8',
  'CondominiumAdmin': '1tREeWG6QugdcGFfC8uy3vSG7Q6DSjVpuVBEtdR094eQ',
  'SPA1': '1yusSqOtLMleYo27LP_fW9aGwLYCDmRZWh654Aliag5o'
};

const CONFIG_SPREADSHEET_ID = '1s4N_pwkwPHMWXlNqcG9dQXm9_yg2jdKImkZdmghKIbs'; 
const CONFIG_SHEET_NAME = 'ConfigViewTB'; 
const USER_CONFIG_SHEET = 'ConfigView';   
const FRIENDLY_NAMES_SHEET = 'ConfigTB';
const DEFAULT_LIMIT = 20; 

/**
 * Función selectora dinámica de base de datos
 */
function getDataSpreadsheetId(userTienda) {
  return DATA_SS_MAP[userTienda] || DATA_SS_MAP['CondominiumAdmin'];
}

/**
 * Función principal: Maneja todas las peticiones GET
 */
function doGet(e) {
  try {
    const action = e.parameter.action;
    const userTienda = e.parameter.userTienda || 'DEFAULT';
    const userProfile = e.parameter.userProfile || 'INVITADO';
    const userId = e.parameter.userId;
    const userName = e.parameter.userName;
    const tableName = e.parameter.tableName || e.parameter.targetSheet;
    const isFullLoad = e.parameter.fullLoad === 'true';

    // --- ACCIONES DE CONFIGURACIÓN (NO requieren tableName) ---

    if (action === 'getAvailableTables') {
      return createJsonResponse(getAvailableTables(userTienda));
    }

    if (action === "getTableFriendlyNames") {
      return createJsonResponse(getTableFriendlyNames(userTienda)); 
    }

    // --- ACCIONES DE PERSISTENCIA (Requieren validaciones específicas) ---

    if (action === 'saveTableConfig') {
      // Extraemos explícitamente el nuevo parámetro ID de tienda
      const userTiendaID = e.parameter.userTiendaID || ''; 
      
      // Pasamos los 6 parámetros en el orden que definimos para la función
      return saveTableConfig(
        userId, 
        userName, 
        userTienda,      // Ya extraído al inicio del doGet
        userTiendaID, 
        tableName, 
        e.parameter.configData
      );
    }

    if (action === 'getTableConfig') {
      const userTiendaID = e.parameter.userTiendaID || '';
      
      // Ahora la búsqueda requiere los 4 elementos de la "llave de unicidad"
      return getTableConfig(userId, userTienda, userTiendaID, tableName);
    }

    // --- VALIDACIÓN DE TABLA PARA LECTURA DE DATOS ---
    // Se coloca aquí para no bloquear las acciones previas
    if (!tableName) {
      return createJsonResponse({ error: "Falta el parámetro 'tableName'" }, 400);
    }

    // --- SEGURIDAD DE PERFIL ---
    const REQUIRED_PROFILE = 'Admin';
    if (userProfile !== REQUIRED_PROFILE) {
      return createJsonResponse({ error: `Acceso denegado. Perfil ${REQUIRED_PROFILE} requerido.`, status: 403 });
    }

    // --- CONEXIÓN DINÁMICA A LOS DATOS ---
    const activeDataId = getDataSpreadsheetId(userTienda);
    const ss = SpreadsheetApp.openById(activeDataId);
    const sheet = ss.getSheetByName(tableName);
    
    if (!sheet) {
      return createJsonResponse({ error: `La tabla '${tableName}' no existe en el documento seleccionado.` }, 404);
    }

    // --- PROCESAR MAPEO DE COLUMNAS ---
    const { keyMap, visibilityMap } = getConfigMap(userTienda, tableName);
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow < 1) return createJsonResponse({ data: [], columnOrder: [] });

    const originalHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    const finalHeaders = [];         
    const finalDisplayMap = {};      
    const columnsToProcess = [];     

    originalHeaders.forEach((name, index) => {
      const cleanKey = String(name).trim().replace(/[^a-zA-Z0-9_]/g, '');
      const isVisible = visibilityMap.hasOwnProperty(name) ? visibilityMap[name] : true; 

      if (cleanKey && isVisible) {
        finalHeaders.push(cleanKey);
        finalDisplayMap[cleanKey] = keyMap[name] || name;
        columnsToProcess.push(index); 
      }
    });

    // --- LÓGICA DE EXTRACCIÓN DE FILAS ---
    const totalDataRows = lastRow > 1 ? lastRow - 1 : 0;
    const rowsToFetch = isFullLoad ? totalDataRows : Math.min(totalDataRows, DEFAULT_LIMIT);

    let data = [];
    if (totalDataRows > 0 && rowsToFetch > 0) {
      const startRow = isFullLoad ? 2 : Math.max(2, (lastRow - rowsToFetch) + 1);
      const values = sheet.getRange(startRow, 1, rowsToFetch, lastCol).getValues();
      
      data = values.map(row => {
        const obj = {};
        columnsToProcess.forEach((colIdx, i) => {
          obj[finalHeaders[i]] = row[colIdx];
        });
        return obj;
      }).reverse();
    }

    return createJsonResponse({
      data: data,
      columnOrder: finalHeaders,
      displayMap: finalDisplayMap,
      dbUsed: userTienda,
      pagination: { 
        totalRows: totalDataRows, 
        fetchedRows: data.length,
        type: isFullLoad ? "FULL" : "PREVIEW" 
      }
    });

  } catch (err) {
    return createJsonResponse({ error: err.toString() }, 500);
  }
}

/**
 * Obtiene el mapeo de nombres amigables desde la hoja "ConfigTB"
 */
function getTableFriendlyNames(appTienda) {
  try {
    const ssConfig = SpreadsheetApp.openById(CONFIG_SPREADSHEET_ID);
    const sheet = ssConfig.getSheetByName(FRIENDLY_NAMES_SHEET);
    if (!sheet) return { success: false, message: "Hoja " + FRIENDLY_NAMES_SHEET + " no encontrada" };

    const data = sheet.getDataRange().getValues();
    const rows = data.slice(1);
    const configMap = {};
    
    rows.forEach(row => {
        const tienda = String(row[0]).trim(); 
        const nombreTecnico = String(row[1]).trim(); 
        const nombreAmigable = String(row[2]).trim();
        
        if (!appTienda || tienda === appTienda) {
            configMap[nombreTecnico] = {
                label: nombreAmigable
            };
        }
    });

    return { success: true, data: configMap };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Persistencia de configuración con marca de tiempo automática (Columna 7)
 * Estructura: [0]ID, [1]Nombre, [2]Tienda, [3]Tienda_ID, [4]Nombre_Tabla, [5]Config_JSON, [6]Fecha_Actualizacion
 */
function saveTableConfig(userId, userName, userTienda, userTiendaID, tableName, configData) {
  const ss = SpreadsheetApp.openById(CONFIG_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(USER_CONFIG_SHEET) || ss.insertSheet(USER_CONFIG_SHEET);
  
  const data = sheet.getDataRange().getValues();
  let foundRow = -1;
  const now = new Date(); // Generamos la fecha y hora actual en el servidor

  // Buscamos si ya existe la configuración para este usuario/tienda/tabla
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == userId && 
        data[i][2] == userTienda && 
        data[i][3] == userTiendaID && 
        data[i][4] == tableName) {
      foundRow = i + 1;
      break;
    }
  }

  if (foundRow !== -1) {
    // ACTUALIZACIÓN
    sheet.getRange(foundRow, 2).setValue(userName); 
    sheet.getRange(foundRow, 6).setValue(configData);
    sheet.getRange(foundRow, 7).setValue(now); // Columna 7: Marca de tiempo
  } else {
    // NUEVO REGISTRO
    // Agregamos la fila con la fecha al final
    sheet.appendRow([userId, userName, userTienda, userTiendaID, tableName, configData, now]);
  }
  
  // Para asegurar el formato dd/mm/yyyy hh:mm:ss, le damos formato a la columna 7
  sheet.getRange("G:G").setNumberFormat("dd/mm/yyyy HH:mm:ss");

  return createJsonResponse({ success: true });
}

function getTableConfig(userId, userTienda, userTiendaID, tableName) {
  const ss = SpreadsheetApp.openById(CONFIG_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(USER_CONFIG_SHEET);
  if (!sheet) return createJsonResponse([]);
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    // La búsqueda para obtener la configuración también debe incluir los nuevos filtros
    if (data[i][0] == userId && 
        data[i][2] == userTienda && 
        data[i][3] == userTiendaID && 
        data[i][4] == tableName) {
      return createJsonResponse(JSON.parse(data[i][5]));
    }
  }
  return createJsonResponse([]);
}

/**
 * Lista de tablas permitidas por Tienda
 */
function getAvailableTables(userTienda) {
  const ssConfig = SpreadsheetApp.openById(CONFIG_SPREADSHEET_ID);
  const sheet = ssConfig.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues().slice(1);
  const tableNames = new Set();
  rows.forEach(row => {
    if (String(row[0]).trim() === userTienda && row[1]) {
      tableNames.add(String(row[1]).trim());
    }
  });
  return Array.from(tableNames);
}

/**
 * Mapeo de cabeceras originales a nombres visibles y visibilidad
 */
function getConfigMap(userTienda, sheetName) {
  const ssConfig = SpreadsheetApp.openById(CONFIG_SPREADSHEET_ID);
  const sheet = ssConfig.getSheetByName(CONFIG_SHEET_NAME);
  const keyMap = {};
  const visibilityMap = {};
  if (sheet) {
    const rows = sheet.getDataRange().getValues().slice(1);
    rows.forEach(row => {
      if (String(row[0]).trim() === userTienda && String(row[1]).trim() === sheetName) {
        const originalHeader = String(row[2]).trim();
        keyMap[originalHeader] = String(row[3] || originalHeader).trim();
        visibilityMap[originalHeader] = !!row[4];
      }
    });
  }
  return { keyMap, visibilityMap };
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
