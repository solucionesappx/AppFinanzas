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
      const tienda = row[0]; // AppTienda
      const nombreTecnico = row[1]; // Nombre_Tabla (ej: TD101_BASIC)
      const nombreAmigable = row[2]; // Descripción_Tabla (ej: PRINCIPAL)
      
      if (!appTienda || tienda === appTienda) {
        configMap[nombreTecnico] = {
          label: nombreAmigable,
          c1: row[3], // ConfgTB01 (opcional para usos futuros)
          c2: row[4]  // ConfgTB02
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

  saveToValoresTable(ssData, masterStructure);
  return masterStructure;
}

function saveToValoresTable(ss, data) {
  let sheet = ss.getSheetByName('TV001_VALORES') || ss.insertSheet('TV001_VALORES');
  
  // 1. Preparar la matriz completa empezando por los encabezados
  const output = [['Nombre_Tabla', 'Encabezado_Tabla', 'Valores_Encabezado']];
  
  // 2. Si hay datos, transformarlos y agregarlos a la matriz
  if (data && data.length > 0) {
    const rows = data.map(item => [
      item.Nombre_Tabla, 
      item.Encabezado_Tabla, 
      JSON.stringify(item.Valores_Encabezado)
    ]);
    output.push(...rows);
  }

  // 3. Limpiar el contenido previo (solo valores, mantiene formatos si los hay)
  sheet.clearContents();

  // 4. Escribir todo el bloque desde la fila 1, columna 1
  // Esto garantiza que los encabezados se escriban siempre en la línea 1
  sheet.getRange(1, 1, output.length, 3).setValues(output);
  
  // 5. Opcional: Proteger los encabezados (Negrita y Congelar fila)
  sheet.getRange(1, 1, 1, 3).setFontWeight("bold");
  if (sheet.getFrozenRows() === 0) sheet.setFrozenRows(1);
}

function extractUniqueValues(ss, sheetName, colName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const idx = data[0].indexOf(colName);
  if (idx === -1) return [];
  return [...new Set(data.slice(1).map(r => r[idx]).filter(c => c !== ""))].sort();
}
