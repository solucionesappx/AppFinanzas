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
          params[key] = '=IMAGE("' + driveUrl + '")';
        }
      }
    }
  }

  const ssData = SpreadsheetApp.openById(DATA_SS_ID);
  const sheet = ssData.getSheetByName(params.TABLA_DESTINO);
  if (!sheet) return createJsonResponse({ success: false, message: 'Tabla no encontrada: ' + params.TABLA_DESTINO });

  const range = sheet.getDataRange();
  const fullData = range.getValues();
  const fullFormulas = range.getFormulas();
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

  // --- 2. PREPARACIÓN DE FILA ---
  const rowValues = (mode === "EDIT") ? [...fullData[rowIndex - 1]] : new Array(headers.length).fill("");
  const rowFormulas = (mode === "EDIT") ? [...fullFormulas[rowIndex - 1]] : new Array(headers.length).fill("");

  headers.forEach((header, index) => {
    const cleanH = header.trim();
    const upperH = cleanH.toUpperCase();
    
    if (cleanH === campoClave && mode === "REGISTER") {
      rowValues[index] = newGeneratedId;
    } 
    else if (upperH.endsWith("IDOWNER")) {
      if (mode === "REGISTER") rowValues[index] = params.currentUser || ""; 
    }
    else if (upperH.endsWith("REGISTROUSER")) {
      rowValues[index] = params.currentUser || "UserSys";
    } 
    else if (upperH.endsWith("REGISTRODATA")) {
      rowValues[index] = timestamp;
    }
    else if (params[cleanH] !== undefined) {
      let val = params[cleanH];
      if (typeof val === "string" && val.trim() !== "" && !upperH.endsWith("IDNOMBRE") && !upperH.endsWith("ID")) {
        if (!isNaN(val) && val.includes('.')) {
          val = parseFloat(val);
        } else if (!isNaN(val)) {
          val = Number(val);
        }
      }
      rowValues[index] = val;
    }
    else if (mode === "EDIT") {
      if (rowFormulas[index] && rowFormulas[index].toString().startsWith('=')) {
        rowValues[index] = rowFormulas[index];
      }
    }
  });

// --- 3. LÓGICA DE TRIPLE REGISTRO (TRANSFERENCIAS) ---
  try {
    if (mode === "REGISTER") {
      // Verificamos si CUENTA2 tiene datos para disparar los asientos vinculados
      const colCuenta2Index = headers.findIndex(h => h.toUpperCase().endsWith("CUENTA2"));
      const valCuenta2OriginalID = colCuenta2Index !== -1 ? rowValues[colCuenta2Index] : null;

      // Identificamos el nombre de la propiedad que viene del frontend (ej: TD101ALIASCUENTA2)
      const keyAlias2 = tablePrefix + "ALIASCUENTA2";
      const aliasCapturado = params[keyAlias2];

      // Si detecto que CUENTA2 no está vacío, cambio su valor por ALIASCUENTA2 antes de guardar el principal
      if (valCuenta2OriginalID && String(valCuenta2OriginalID).trim() !== "" && aliasCapturado) {
        rowValues[colCuenta2Index] = aliasCapturado;
      }

      // Guardamos primero el registro principal (TRANSFERENCIA)
      sheet.appendRow(rowValues);

      // Si valCuenta2 tenía datos, procedemos con los otros 2 registros/asientos
      if (valCuenta2OriginalID && String(valCuenta2OriginalID).trim() !== "") {
        const colMovIndex = headers.findIndex(h => h.toUpperCase().endsWith("MOVIMIENTO"));
        const colCuentaIndex = headers.findIndex(h => h.toUpperCase().endsWith("CUENTA") && !h.toUpperCase().endsWith("CUENTA2"));
        const colAliasIndex = headers.findIndex(h => h.toUpperCase().endsWith("ALIASCUENTA"));
        const colRefIndex = headers.findIndex(h => h.toUpperCase().endsWith("REF"));

        // --- Asiento DÉBITO (Salida de dinero) ---
        let rowDebito = [...rowValues];
        rowDebito[idColIndex] = newGeneratedId + 1; // ID siguiente
        if (colMovIndex !== -1) rowDebito[colMovIndex] = "DÉBITO";
        if (colRefIndex !== -1) rowDebito[colRefIndex] = "TRANSFERENCIA " + newGeneratedId;
        if (colCuenta2Index !== -1) rowDebito[colCuenta2Index] = ""; // Limpiar receptora
        sheet.appendRow(rowDebito);

        // --- Asiento CRÉDITO (Entrada de dinero) ---
        let rowCredito = [...rowValues];
        rowCredito[idColIndex] = newGeneratedId + 2; // ID subsiguiente
        if (colMovIndex !== -1) rowCredito[colMovIndex] = "CRÉDITO";
        if (colRefIndex !== -1) rowCredito[colRefIndex] = "TRANSFERENCIA " + newGeneratedId;
        
        // 1. La cuenta destino (ID original de CUENTA2) se convierte en la principal (CUENTA)
        if (colCuentaIndex !== -1 && valCuenta2OriginalID) {
            rowCredito[colCuentaIndex] = valCuenta2OriginalID;
        }

        // 2. El Alias capturado se convierte en el ALIASCUENTA principal
        if (colAliasIndex !== -1 && aliasCapturado) {
            rowCredito[colAliasIndex] = aliasCapturado;
        }
        
        // 3. Limpiar la referencia a la cuenta secundaria
        if (colCuenta2Index !== -1) rowCredito[colCuenta2Index] = ""; 
        
        sheet.appendRow(rowCredito);
      }
    } else {
      // Escribimos la fila completa para el modo EDIT. 
      // Aquellas posiciones con "=" serán tratadas como fórmulas por Sheets.
      sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowValues]);
    }

    const responseObj = {};
    headers.forEach((h, i) => responseObj[h.trim()] = rowValues[i]);

    return createJsonResponse({ 
      success: true, 
      message: mode === "REGISTER" ? 'Procesado con asientos' : 'Actualizado', 
      data: responseObj 
    });

  } catch (e) {
    return createJsonResponse({ success: false, message: 'Error en persistencia: ' + e.toString() });
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
// NO USAR 
function contarYListarCuentasPorTienda() {
  const ssConfig = SpreadsheetApp.openById(CONFIG_SS_ID);
  const ssData = SpreadsheetApp.openById(DATA_SS_ID);
  
  // 1. Obtener TODOS los usuarios de la tienda "SPA1"
  const sheetUsuarios = ssConfig.getSheetByName("USUARIOS");
  const dataUsuarios = sheetUsuarios.getDataRange().getValues();
  const headersUsuarios = dataUsuarios[0];
  
  const colUserTiendaRef = headersUsuarios.indexOf("Usuario_Tienda");
  const colUserTiendaIDRef = headersUsuarios.indexOf("Usuario_Tienda_ID");
  
  // CAMBIO CLAVE: .filter() en lugar de .find() para no detenerse en el primero
  const usuariosSPA1 = dataUsuarios.slice(1).filter(row => row[colUserTiendaRef] === "SPA1");
  
  if (usuariosSPA1.length === 0) {
    Logger.log("❌ Error: No se encontraron usuarios para la tienda 'SPA1' en la tabla USUARIOS.");
    return;
  }
  
  Logger.log("👥 Usuarios encontrados para SPA1: " + usuariosSPA1.length);

  // 2. Cargar tabla TV002_DICCIONARIO
  const sheetDict = ssData.getSheetByName("TV002_DICCIONARIO");
  const dataDict = sheetDict.getDataRange().getValues();
  const headersDict = dataDict[0];
  
  const colDataTiendaNombre = headersDict.indexOf("Usuario_Tienda");
  const colDataTiendaID = headersDict.indexOf("Usuario_Tienda_ID");

  if (colDataTiendaNombre === -1 || colDataTiendaID === -1) {
    Logger.log("❌ ERROR: No se encontraron las columnas necesarias en TV002_DICCIONARIO.");
    return;
  }

  // 3. Barrido de cada usuario encontrado
  usuariosSPA1.forEach((usuario) => {
    const idTiendaBuscado = String(usuario[colUserTiendaIDRef]).trim();
    const nombreTiendaBuscado = "SPA1";

    Logger.log("--------------------------------------------------");
    Logger.log("🔍 Procesando Usuario_ID: " + idTiendaBuscado);

    // Filtrar en el diccionario para ESTE Usuario_ID específico
    const registrosEncontrados = dataDict.slice(1).filter(row => {
      const cumpleNombre = String(row[colDataTiendaNombre]).trim() === nombreTiendaBuscado;
      const cumpleID = String(row[colDataTiendaID]).trim() === idTiendaBuscado;
      return cumpleNombre && cumpleID;
    });

    // 4. Reporte por Usuario
    if (registrosEncontrados.length > 0) {
      Logger.log("✅ Coincidencias para " + idTiendaBuscado + ": " + registrosEncontrados.length);
      
      registrosEncontrados.forEach((fila, index) => {
        let detalleFila = {};
        headersDict.forEach((header, i) => {
          detalleFila[header] = fila[i];
        });
        Logger.log("   -> [" + (index + 1) + "]: " + JSON.stringify(detalleFila));
      });
    } else {
      Logger.log("⚠️ Sin registros en TV002_DICCIONARIO para el ID: " + idTiendaBuscado);
    }
  });
} 

function ejecutarActualizacionDiccionario(params) {
  const ssData = SpreadsheetApp.openById(DATA_SS_ID);
  const sheetDict = ssData.getSheetByName("TV002_DICCIONARIO");
  const dataDict = sheetDict.getDataRange().getValues();
  const headersDict = dataDict[0];

  // 1. Mapeo de índices según tu estructura
  const colTiendaNombre = headersDict.indexOf("Usuario_Tienda");
  const colTabla = headersDict.indexOf("Nombre_Tabla");
  const colCampo = headersDict.indexOf("Encabezado_Tabla");
  const colValores = headersDict.indexOf("Valores_Encabezado");
  const colTiendaID = headersDict.indexOf("Usuario_Tienda_ID");

  // 2. Búsqueda de fila existente
  let filaEncontrada = -1;
  const idBuscado = String(params.usuarioId).trim();
  const tablaBuscada = String(params.nombreTabla).trim();
  const campoBuscado = String(params.encabezadoTabla).trim();

  for (let i = 1; i < dataDict.length; i++) {
    if (String(dataDict[i][colTiendaID]).trim() === idBuscado && 
        String(dataDict[i][colTabla]).trim() === tablaBuscada && 
        String(dataDict[i][colCampo]).trim() === campoBuscado) {
      filaEncontrada = i + 1;
      break;
    }
  }

  // 3. Preparar los datos (Asegurar que todos los campos tengan valor)
  const tiendaNombre = params.userTienda || "SPA1"; // Viene del frontend o default
  const valoresJson = params.nuevoValor; // Ya viene como JSON string del frontend

  if (filaEncontrada !== -1) {
    // ACTUALIZAR REGISTRO EXISTENTE
    const rango = sheetDict.getRange(filaEncontrada, 1, 1, headersDict.length);
    const valoresFila = [];
    valoresFila[colTiendaNombre] = tiendaNombre;
    valoresFila[colTabla] = tablaBuscada;
    valoresFila[colCampo] = campoBuscado;
    valoresFila[colValores] = valoresJson;
    valoresFila[colTiendaID] = idBuscado;
    
    // Aplicamos los valores uno por uno en sus columnas correspondientes
    sheetDict.getRange(filaEncontrada, colTiendaNombre + 1).setValue(tiendaNombre);
    sheetDict.getRange(filaEncontrada, colTabla + 1).setValue(tablaBuscada);
    sheetDict.getRange(filaEncontrada, colCampo + 1).setValue(campoBuscado);
    sheetDict.getRange(filaEncontrada, colValores + 1).setValue(valoresJson);
    sheetDict.getRange(filaEncontrada, colTiendaID + 1).setValue(idBuscado);

    console.log("✅ Registro actualizado en TV002_DICCIONARIO");
  } else {
    // CREAR NUEVO REGISTRO (Append)
    const nuevaFila = [];
    nuevaFila[colTiendaNombre] = tiendaNombre;
    nuevaFila[colTabla] = tablaBuscada;
    nuevaFila[colCampo] = campoBuscado;
    nuevaFila[colValores] = valoresJson;
    nuevaFila[colTiendaID] = idBuscado;
    
    sheetDict.appendRow(nuevaFila);
    console.log("✨ Nuevo registro creado en TV002_DICCIONARIO");
  }

  return ContentService.createTextOutput(JSON.stringify({ 
    success: true, 
    message: filaEncontrada !== -1 ? "Actualizado" : "Creado" 
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Borra en cascada el registro maestro de TRANSFERENCIA y sus asientos hijos.
 */
function handleDeleteTransferCascade(params) {
  const ssData = SpreadsheetApp.openById(DATA_SS_ID);
  const sheet = ssData.getSheetByName(params.TABLA_DESTINO);
  
  if (!sheet) return createJsonResponse({ success: false, message: 'Tabla no encontrada.' });

  const tablePrefix = params.TABLA_DESTINO.split('_')[0].toUpperCase();
  const idColName = tablePrefix + "ID";
  const refColName = tablePrefix + "REF";
  const masterId = params[idColName]; // El ID de la fila TRANSFERENCIA
  const refPattern = params.REF_CASCADE; // Ejemplo: "TRANSFERENCIA #105"

  if (!masterId || !refPattern) {
    return createJsonResponse({ success: false, message: 'Faltan datos para el borrado en cascada.' });
  }

  const range = sheet.getDataRange();
  const data = range.getValues();
  const headers = data[0];
  
  const idColIndex = headers.indexOf(idColName);
  const refColIndex = headers.indexOf(refColName);

  if (idColIndex === -1 || refColIndex === -1) {
    return createJsonResponse({ success: false, message: 'Columnas ID o REF no encontradas.' });
  }

  let rowsDeleted = 0;

  // Recorremos de abajo hacia arriba para no alterar los índices al eliminar
  for (let i = data.length - 1; i >= 1; i--) {
    const currentRow = data[i];
    const rowId = String(currentRow[idColIndex]);
    const rowRef = String(currentRow[refColIndex]);

    // REGLA: Borrar si el ID coincide (Maestro) O si la REF coincide (Hijos)
    if (rowId === String(masterId) || rowRef === String(refPattern)) {
      // Si usas historial, llamamos a tu función existente:
      moveRowToHistory(ssData, sheet, i + 1, headers, params);
      rowsDeleted++;
    }
  }

  return createJsonResponse({ 
    success: true, 
    message: 'Cascada completada. ' + rowsDeleted + ' registros procesados.' 
  });
}

function buscarDatoCuentaG(aliasBuscado) {
  if (!aliasBuscado) return "";
  
  const ss = SpreadsheetApp.openById(DATA_SS_ID);
  const sheetCuentas = ss.getSheetByName("TD102_CUENTAS");
  if (!sheetCuentas) return aliasBuscado; // Si no existe la tabla, devolvemos el original

  const data = sheetCuentas.getDataRange().getValues();
  
  // Buscamos en la Columna A (índice 0) y devolvemos la Columna G (índice 6)
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toUpperCase() === String(aliasBuscado).trim().toUpperCase()) {
      return data[i][6]; // Columna G
    }
  }
  
  return aliasBuscado; // Si no lo encuentra, deja lo que venía
}
