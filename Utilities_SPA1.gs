// Acceso rápido al documento de datos
//const getDataSheet = (name) => SpreadsheetApp.openById(DATA_SS_ID).getSheetByName(name);



/**
 * Obtiene nombres amigables para el selector de tablas
 */
function getTableFriendlyNames(appTienda) {
  try {
    const ssConfig = SpreadsheetApp.openById(CONFIG_SS_ID);
    const sheet = ssConfig.getSheetByName("ConfigTB");
    if (!sheet) return { success: false, message: "Hoja ConfigTB no encontrada" };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);

    // Mapeamos los datos para que el frontend reciba un diccionario útil
    // Filtramos por AppTienda (si se proporciona)
    const configMap = {};
    
    rows.forEach(row => {
      const tienda = String(row[0]).trim();
      const nombreTecnico = String(row[1]).trim();
      const nombreAmigable = String(row[2]).trim();
      
      if (!appTienda || tienda === appTienda) {
        configMap[nombreTecnico] = {
          // Si nombreAmigable está vacío, mandamos el técnico con la marca de error
          label: nombreAmigable ? nombreAmigable : `${nombreTecnico} (Sin nombre)`,
          c1: row[3] || "",
          c2: row[4] || ""
        };
      }
    });

    return { success: true, data: configMap };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Procesa todas las hojas del documento DATA_SS_ID y extrae los encabezados
 * para consolidarlos en una tabla dentro de 'hojaX'.
 */
function generateHeadersInventory() {
  const TARGET_SHEET_NAME = 'hojaX';
  
  try {
    const ss = SpreadsheetApp.openById(DATA_SS_ID);
    const sheets = ss.getSheets();
    let inventoryData = [];

    // Iterar por cada hoja del documento
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      
      // Evitar procesar la hoja de destino para no crear bucles de datos
      if (sheetName === TARGET_SHEET_NAME) return;

      // Obtener la primera fila (encabezados)
      // getRange(fila, columna, numFilas, numColumnas)
      const lastColumn = sheet.getLastColumn();
      
      if (lastColumn > 0) {
        const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
        
        headers.forEach(header => {
          if (header !== "") {
            inventoryData.push([sheetName, header]);
          }
        });
      }
    });

    // Gestión de la hoja de destino 'hojaX'
    let targetSheet = ss.getSheetByName(TARGET_SHEET_NAME);
    if (!targetSheet) {
      targetSheet = ss.insertSheet(TARGET_SHEET_NAME);
    } else {
      targetSheet.clearContents(); // Limpiar contenido previo
    }

    // Insertar encabezados de la nueva tabla
    targetSheet.getRange(1, 1, 1, 2).setValues([["Hoja", "Columna"]]);
    targetSheet.getRange(1, 1, 1, 2).setFontWeight("bold");

    // Insertar los datos recolectados
    if (inventoryData.length > 0) {
      targetSheet.getRange(2, 1, inventoryData.length, 2).setValues(inventoryData);
    }

    Logger.log("Inventario generado con éxito en " + TARGET_SHEET_NAME);
    
  } catch (e) {
    Logger.log("Error: " + e.toString());
  }
}

/**
 * Procesa registros (Crear/Editar) de forma dinámica preservando la integridad
 * de todas las tablas y normalizando formatos numéricos.
 */
function handleDynamicDataTD(params, mode) {
  // --- DETECCIÓN DE IMAGEN ---
  for (var key in params) {
    if (key.toUpperCase().includes('IMAGEN')) {
      var val = params[key];
      if (typeof val === 'string' && val.indexOf('data:image') === 0) {
        var nombreArchivo = "IMG_" + new Date().getTime() + ".jpg";
        var driveUrl = uploadImageToDrive(val, nombreArchivo);
        
        if (driveUrl && !driveUrl.startsWith("Error")) {
          // Guardamos la fórmula para que veas la miniatura en el Excel/Sheets
          params[key] = '=IMAGE("' + driveUrl + '")';
        }
      }
    }
  }
  // --- FIN DETECCIÓN ---

  const ssData = SpreadsheetApp.openById(DATA_SS_ID);
  const sheet = ssData.getSheetByName(params.TABLA_DESTINO);
  
  if (!sheet) return createJsonResponse({ success: false, message: 'Tabla no encontrada: ' + params.TABLA_DESTINO });

  const range = sheet.getDataRange();
  const fullData = range.getValues();      // Valores resultantes
  const fullFormulas = range.getFormulas(); // Fórmulas originales
  const headers = fullData[0];
  const timestamp = Utilities.formatDate(new Date(), "GMT-4", "dd/MM/yyyy HH:mm:ss");
  
  const tablePrefix = params.TABLA_DESTINO.split('_')[0].toUpperCase();
  const campoClave = params.CAMPO_CLAVE || (tablePrefix + "ID");
  
  const idColIndex = headers.indexOf(campoClave);
  if (idColIndex === -1) return createJsonResponse({ success: false, message: 'Falta columna clave: ' + campoClave });

  let rowIndex = -1;
  let newGeneratedId = null;

  if (mode === "REGISTER") {
    newGeneratedId = generateNextIDInternal(fullData, tablePrefix);
  } else {
    const rawIdValue = params[campoClave] || params.ID_VALUE;
    const valorBusqueda = Number(String(rawIdValue).replace(/[.,\s]/g, ''));

    for (let i = 1; i < fullData.length; i++) {
      const cellValue = Number(String(fullData[i][idColIndex]).replace(/[.,\s]/g, ''));
      if (cellValue === valorBusqueda) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex === -1) return createJsonResponse({ success: false, message: 'ID no hallado.' });
    if (mode === "DELETE") return moveRowToHistory(ssData, sheet, rowIndex, headers, params);
  }

  // --- 2. PREPARACIÓN DE FILA (INTELIGENCIA DE FÓRMULAS) ---
  const rowValues = (mode === "EDIT") ? [...fullData[rowIndex - 1]] : new Array(headers.length).fill("");
  const rowFormulas = (mode === "EDIT") ? [...fullFormulas[rowIndex - 1]] : new Array(headers.length).fill("");

  headers.forEach((header, index) => {
    const cleanH = header.trim();
    const upperH = cleanH.toUpperCase();
    
    // A. Asignación de Llave Primaria
    if (cleanH === campoClave && mode === "REGISTER") {
      rowValues[index] = newGeneratedId;
    } 
    // B. Auditoría
    else if (upperH.endsWith("REGISTROUSER")) {
      rowValues[index] = params[cleanH] || params.currentUser || "UserSys";
    } 
    else if (upperH.endsWith("REGISTRODATA")) {
      rowValues[index] = timestamp;
    } 
    // C. Datos del Frontend vs Fórmulas
    else if (params[cleanH] !== undefined) {
      let val = params[cleanH];
      
      // Si el valor es una cadena que parece número (y no es un ID o Código)
      if (typeof val === "string" && val.trim() !== "" && !upperH.endsWith("IDNOMBRE") && !upperH.endsWith("ID")) {
        // Si el frontend envía el valor "limpio" (ej. "1250.50"), 
        // nos aseguramos de que Google Sheets lo trate como número
        if (!isNaN(val) && val.includes('.')) {
          val = parseFloat(val);
        } else if (!isNaN(val)) {
          val = Number(val);
        }
      }
      rowValues[index] = val;
    }
    // D. LÓGICA CRÍTICA: Si el campo NO viene en el payload (Calculado o omitido)
    else if (mode === "EDIT") {
      // Si la celda original tenía una fórmula, la PRESERVAMOS sobre el valor estático
      if (rowFormulas[index] && rowFormulas[index].toString().startsWith('=')) {
        rowValues[index] = rowFormulas[index];
      }
      // Si no es fórmula, rowValues[index] ya tiene el valor estático de fullData[rowIndex-1]
    }
  });

  // 3. PERSISTENCIA
  try {
    if (mode === "REGISTER") {
      sheet.appendRow(rowValues);
    } else {
      // Escribimos la fila completa. Aquellas posiciones con "=" serán tratadas como fórmulas por Sheets.
      sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowValues]);
    }

    const responseObj = {};
    headers.forEach((h, i) => responseObj[h.trim()] = rowValues[i]);

    return createJsonResponse({ 
      success: true, 
      message: mode === "EDIT" ? 'Registro actualizado.' : 'Creado correctamente.',
      data: responseObj 
    });
  } catch (e) {
    return createJsonResponse({ success: false, message: 'Error: ' + e.toString() });
  }
}

function moveRowToHistory(ss, sourceSheet, rowIndex, headers, params) {
  const historySheet = ss.getSheetByName("TD999_BORRADOS");
  if (!historySheet) return createJsonResponse({ success: false, message: 'Tabla TD999_BORRADOS no hallada.' });

  // 1. Obtener datos actuales antes de borrar
  const rowDataArray = sourceSheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  const rowDataObj = {};
  headers.forEach((h, i) => rowDataObj[h.trim()] = rowDataArray[i]);

  const timestamp = Utilities.formatDate(new Date(), "GMT-4", "dd/MM/yyyy HH:mm:ss");
  const tablePrefix = params.TABLA_DESTINO.split('_')[0].toUpperCase();
  const idDocumento = rowDataObj[tablePrefix + "ID"] || "N/A";

  // 2. Generar ID Correlativo para TD999ID (Busca el máximo para permitir orden descendente)
  const lastRow = historySheet.getLastRow();
  let nextId = 999001;
  if (lastRow > 1) {
    // Obtenemos todos los IDs de la primera columna
    const allIds = historySheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const maxId = Math.max(...allIds.filter(id => !isNaN(id)));
    if (maxId >= 999001) nextId = maxId + 1;
  }

  // 3. Estructura: TD999ID, TD999IDDOC, TD999DATAJSON, TD999RegistroUser, TD999RegistroData
  const historyRow = [
    nextId,           // TD999ID
    idDocumento,      // TD999IDDOC
    JSON.stringify(rowDataObj), 
    params.usuario_id || "User", 
    timestamp
  ];

  try {
    // 4. Insertar el registro
    historySheet.appendRow(historyRow);

    // 5. ORDENAR DESCENDENTE (Por la columna 1: TD999ID)
    const newLastRow = historySheet.getLastRow();
    if (newLastRow > 1) {
      const lastCol = historySheet.getLastColumn();
      // Aplicamos el sort a todo el rango de datos (excluyendo encabezado)
      historySheet.getRange(2, 1, newLastRow - 1, lastCol)
                  .sort({ column: 1, ascending: false });
    }

    // 6. Eliminar de la hoja original
    sourceSheet.deleteRow(rowIndex);

    return createJsonResponse({ 
      success: true, 
      message: 'Registro eliminado correctamente.' 
    });
  } catch (e) {
    return createJsonResponse({ success: false, message: 'Error en archivo: ' + e.toString() });
  }
}

function generateNextIDInternal(fullData, prefix) {
  const numericPrefix = prefix.replace(/\D/g, "");
  const rangeStart = parseInt(numericPrefix + "1001");
  const rangeEnd = parseInt(numericPrefix + "9999");
  
  const ids = fullData.slice(1).map(row => {
    if (!row[0]) return null;
    const cleanId = String(row[0]).replace(/[.,\s]/g, "");
    const numId = parseInt(cleanId);
    return isNaN(numId) ? null : numId;
  }).filter(id => id !== null && id >= rangeStart && id <= rangeEnd);

  const maxId = ids.length === 0 ? rangeStart - 1 : Math.max(...ids);
  const nextId = maxId + 1;

  if (nextId > rangeEnd) throw new Error("Rango agotado para " + prefix);
  return nextId;
}

function syncAndGetMasterFields() {
  const ssConfig = SpreadsheetApp.openById(CONFIG_SS_ID);
  const ssData = SpreadsheetApp.openById(DATA_SS_ID);
  const configData = ssConfig.getSheetByName(CONFIG_SHEET_NAME).getDataRange().getValues().slice(1);

  const fieldCounts = {};
  configData.forEach(row => {
    const field = String(row[2]).trim();
    if (field) fieldCounts[field] = (fieldCounts[field] || 0) + 1;
  });

  const sharedFields = Object.keys(fieldCounts).filter(f => fieldCounts[f] > 1);
  const masterStructure = [];

  sharedFields.forEach(field => {
    const fieldPrefix = field.substring(0, 5).toUpperCase();
    const baseRow = configData.find(row => {
      const tName = String(row[1]).toUpperCase();
      return String(row[2]) === field && tName.startsWith(fieldPrefix);
    });

    if (baseRow) {
      const baseTableName = baseRow[1];
      const values = extractUniqueValues(ssData, baseTableName, field);
      masterStructure.push({
        Nombre_Tabla: baseTableName,
        Encabezado_Tabla: field,
        Valores_Encabezado: values
      });
    }
  });

  return masterStructure;
}

function extractUniqueValues(ss, sheetName, colName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const idx = data[0].indexOf(colName);
  if (idx === -1) return [];
  return [...new Set(data.slice(1).map(r => r[idx]).filter(c => c !== ""))].sort();
}

function forzarPermisos() {
  const folder = DriveApp.getFolderById("1NzhsEJy51DQCPxOYDuK2rC8mXAICeOPF");
  Logger.log("Acceso concedido a: " + folder.getName());
}

function uploadImageToDrive(base64Data, fileName) {
  try {
    var folderId = "1NzhsEJy51DQCPxOYDuK2rC8mXAICeOPF";
    var folder = DriveApp.getFolderById(folderId);
    
    var parts = base64Data.split(',');
    var contentType = parts[0].split(':')[1].split(';')[0];
    var decoded = Utilities.base64Decode(parts[1]);
    var blob = Utilities.newBlob(decoded, contentType, fileName);
    
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Este formato de URL es el que permite que Google Sheets renderice la imagen
    var directLink = "https://drive.google.com/uc?export=download&id=" + file.getId();
    
    return directLink;
  } catch (e) {
    console.error("Fallo en Drive: " + e.toString());
    return "Error: " + e.toString();
  }
}

/**
 * Obtiene el diccionario de valores permitidos para los campos de la App.
 * Se filtra por App_Tienda para que el usuario solo descargue lo que le compete.
 */
function getDictionaryData(appTienda) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("TV002_DICCIONARIO");
  
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data.shift(); // Quitamos los encabezados del Excel
  
  // Mapeamos los datos a objetos JSON
  const dictionary = data.map(row => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  }).filter(item => {
    // Filtramos para que solo traiga los del usuario actual o globales
    return item.App_Tienda === appTienda || item.App_Tienda === 'GLOBAL';
  });

  return dictionary;
}
