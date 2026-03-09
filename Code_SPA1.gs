const DATA_SS_ID = '1yusSqOtLMleYo27LP_fW9aGwLYCDmRZWh654Aliag5o'; 
const CONFIG_SS_ID = '1s4N_pwkwPHMWXlNqcG9dQXm9_yg2jdKImkZdmghKIbs'; 
const CONFIG_SHEET_NAME = 'ConfigViewTB';

//DriveApp.getFiles();

/**
 * Función Principal Receptora de Apps Script - Versión Completa con Diccionario
 */
function doGet(e) {
  try {
    const action = e.parameter.action;
    const appTienda = e.parameter.appTienda;
    const userTienda = e.parameter.userTienda || 'DEFAULT';

    // Acción para obtener nombres amigables (Selector de tablas)
    if (action === "getTableFriendlyNames") {
      const result = getTableFriendlyNames(appTienda || userTienda);
      return createJsonResponse(result);
    }

    const tableName = e.parameter.tableName || e.parameter.sheet;
    if (!tableName) throw new Error("Parámetro 'tableName' omitido.");

    const ignoreVisibility = e.parameter.ignoreVisibility === 'true'; 

    const ssData = SpreadsheetApp.openById(DATA_SS_ID);
    const ssConfig = SpreadsheetApp.openById(CONFIG_SS_ID);
    
    const configSheet = ssConfig.getSheetByName(CONFIG_SHEET_NAME);
    const dataSheet = ssData.getSheetByName(tableName);

    if (!dataSheet) throw new Error("La tabla '" + tableName + "' no existe.");
    if (!configSheet) throw new Error("La hoja de configuración no existe.");

    // --- 1. OBTENER CONFIGURACIÓN DE COLUMNAS ---
    const configRows = configSheet.getDataRange().getValues();
    const configData = configRows.slice(1);
    const configMap = {};
    const availableTables = [];
    const fullConfigForFrontend = [];

    configData.forEach(row => {
        const rowAppTienda = String(row[0]).trim();
        const nombreTabla = String(row[1]).trim();
        
        // Filtro para tablas disponibles por usuario
        if (rowAppTienda === userTienda && !availableTables.includes(nombreTabla)) {
          availableTables.push(nombreTabla);
        }

        // Validación de coincidencia de tabla y tienda
        if (nombreTabla === tableName && rowAppTienda === userTienda) {
          const idColumna = String(row[2]).trim();
          const upperColId = idColumna.toUpperCase();
          const tablePrefix = tableName.split('_')[0].toUpperCase();
          const esID = (upperColId === `${tablePrefix}ID`);

          const configObj = {
            ID_Columna: idColumna,
            Nombre_Encabezado: String(row[3] || idColumna).trim(),
            Visible_Encabezado: String(row[4] || "").trim(),
            Justificado_Campo: String(row[5] || "left").trim().toLowerCase(),
            Es_Obligatorio: !esID && String(row[6] || "").trim().toLowerCase() === "x", 
            Es_Calculado: !esID && String(row[6] || "").trim().toLowerCase() === "calc"
          };
          configMap[idColumna] = configObj;
          fullConfigForFrontend.push(configObj);
        }
    });

    // --- 2. PROCESAR DATOS DE LA TABLA ---
    const fullData = dataSheet.getDataRange().getValues();
    if (fullData.length === 0) throw new Error("La tabla está vacía.");
    
    const originalHeaders = fullData[0];
    const tablePrefix = tableName.split('_')[0].toUpperCase();
    const finalHeaders = [];
    const finalDisplayMap = {};
    const finalAlignMap = {};
    const colIndexesToFetch = [];

    originalHeaders.forEach((headerName, index) => {
      const cleanH = String(headerName).trim();
      const upperH = cleanH.toUpperCase();
      const config = configMap[cleanH];
      
      const isPK = upperH === `${tablePrefix}ID`;
      const isAuditField = upperH.endsWith("REGISTROUSER") || upperH.endsWith("REGISTRODATA");
      const isTypeReg = upperH.endsWith("TYPEREG");
      
      if (ignoreVisibility || (config && config.Visible_Encabezado !== "") || isPK || isAuditField || isTypeReg) {
        finalHeaders.push(cleanH);
        finalDisplayMap[cleanH] = (config && config.Nombre_Encabezado) ? config.Nombre_Encabezado : cleanH;
        finalAlignMap[cleanH] = (config && config.Justificado_Campo) ? config.Justificado_Campo : 'left';
        colIndexesToFetch.push(index);
      }
    });

    const jsonData = fullData.slice(1).map(row => {
      const obj = {};
      colIndexesToFetch.forEach((colIdx, i) => { obj[finalHeaders[i]] = row[colIdx]; });
      return obj;
    });

    // --- 3. SINCRONIZACIÓN MAESTRA ---
    const masterFields = typeof syncAndGetMasterFields === "function" ? syncAndGetMasterFields(ssData) : []; 

    // --- 4. CARGA DEL DICCIONARIO DINÁMICO (TV002_DICCIONARIO) ---
    const dictionaryData = [];
    try {
      const dictSheet = ssData.getSheetByName("TV002_DICCIONARIO");
      if (dictSheet) {
        const dictValues = dictSheet.getDataRange().getValues();
        const dictRows = dictValues.slice(1);

        dictRows.forEach(row => {
          const dictAppTienda = String(row[0]).trim();
          // Filtramos para que el usuario solo reciba su tienda o valores globales
          if (dictAppTienda === userTienda || dictAppTienda === 'GLOBAL') {
            dictionaryData.push({
              Nombre_Tabla: String(row[1]).trim(),
              Encabezado_Tabla: String(row[2]).trim(),
              Valores_Encabezado: String(row[3] || "[]").trim()
            });
          }
        });
      }
    } catch (dictErr) {
      console.error("Error cargando diccionario: " + dictErr.toString());
      // No lanzamos error para no interrumpir el flujo principal de datos
    }

    // --- 5. RESPUESTA FINAL UNIFICADA ---
    return createJsonResponse({
      success: true,
      data: jsonData,
      columnOrder: finalHeaders,
      displayMap: finalDisplayMap,
      alignMap: finalAlignMap,
      fullConfig: fullConfigForFrontend,
      availableTables: availableTables,
      masterFields: masterFields,
      dictionary: dictionaryData // Inyectado para que el frontend actualice VALUE_DICTIONARY
    });

  } catch (err) {
    return createJsonResponse({ success: false, message: err.toString() });
  }
}

function doPost(e) {
  try {
    let params;

    // 1. DETERMINAR EL ORIGEN DE LOS DATOS
    // Si el contenido empieza con '{', intentamos JSON. Si no, usamos e.parameter
    if (e.postData && e.postData.contents && e.postData.contents.charAt(0) === '{') {
      params = JSON.parse(e.postData.contents);
    } else {
      // e.parameter ya contiene los datos parseados si vienen como formulario
      params = e.parameter;
    }

    // 2. VERIFICACIÓN DE EMERGENCIA
    if (!params || Object.keys(params).length === 0) {
      throw new Error("No se recibieron parámetros en el Backend.");
    }

    // 3. PROCESAMIENTO DE IMAGEN (Aquí es donde ocurre la magia del Drive)
    params = detectarYProcesarImagenes(params);

    // 4. EJECUCIÓN DE ACCIONES
    const action = params.action;
    let result;

    if (action === "registerDynamicDataTD") {
      result = handleDynamicDataTD(params, "REGISTER");
    } else if (action === "editDynamicDataTD") {
      result = handleDynamicDataTD(params, "EDIT");
    } else if (action === "deleteDynamicDataTD") {
      result = handleDynamicDataTD(params, "DELETE");
    } else {
      throw new Error("Acción desconocida: " + action);
    }

    syncAndGetMasterFields();
    return result;

  } catch (err) {
    // Retornamos el error en formato JSON para que el Frontend lo muestre bonito
    return ContentService.createTextOutput(JSON.stringify({ 
      success: false, 
      message: "Falla en doPost: " + err.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Esta función debe estar en tu archivo .gs también
function detectarYProcesarImagenes(datos) {
  // Creamos una copia para no alterar el original mientras iteramos
  let nuevosDatos = {};
  for (let key in datos) {
    nuevosDatos[key] = datos[key];
  }

  for (let key in nuevosDatos) {
    if (key.toUpperCase().includes('IMAGEN')) {
      let val = nuevosDatos[key];
      if (typeof val === 'string' && val.indexOf('data:image') === 0) {
        let nombreArchivo = "IMG_" + new Date().getTime() + ".jpg";
        let driveUrl = uploadImageToDrive(val, nombreArchivo);
        
        if (driveUrl && !driveUrl.startsWith("Error")) {
          // Reemplazamos el Base64 por la fórmula de imagen
          nuevosDatos[key] = '=IMAGE("' + driveUrl + '")';
        }
      }
    }
  }
  return nuevosDatos;
}

/**
 * Utilidad para responder en formato JSON
 */
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}



