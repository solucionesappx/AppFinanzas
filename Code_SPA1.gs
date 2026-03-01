const DATA_SS_ID = '1yusSqOtLMleYo27LP_fW9aGwLYCDmRZWh654Aliag5o'; 
const CONFIG_SS_ID = '1s4N_pwkwPHMWXlNqcG9dQXm9_yg2jdKImkZdmghKIbs'; 
const CONFIG_SHEET_NAME = 'ConfigViewTB';

// Nombres técnicos exactos de tus hojas de datos
const HOJAS = {
  TPE: 'TD101_TPE',
  VALENCIAS: 'TD102_VALENCIAS',
  NOMENCLATURA: 'TD201_NOMENCLATURA',
  EXCEPCIONES: 'TD202_EXCEPCIONES',
  REACCIONES: 'TD301_REACCIONES'
};

/**
 * Función Principal Receptora de Apps Script
 */
function doGet(e) {
  try {
    const action = e.parameter.action;
    const appTienda = e.parameter.appTienda;
    const userTienda = e.parameter.userTienda || 'DEFAULT';

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

    // 1. Obtener Configuración de Columnas
    const configRows = configSheet.getDataRange().getValues();
    const configData = configRows.slice(1);
    const configMap = {};
    const availableTables = [];
    const fullConfigForFrontend = [];

    configData.forEach(row => {
        const rowAppTienda = String(row[0]).trim();
        const nombreTabla = String(row[1]).trim();
        
        // Filtro para tablas disponibles (Correcto)
        if (rowAppTienda === userTienda && !availableTables.includes(nombreTabla)) {
          availableTables.push(nombreTabla);
        }

        // CORRECCIÓN: Validar que el nombre de la tabla coincida Y que pertenezca al userTienda (App ID)
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

    // 2. Procesar Datos de la Tabla
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

    // 3. Sincronización Maestra
    const masterFields = typeof syncAndGetMasterFields === "function" ? syncAndGetMasterFields(ssData) : []; 

    return createJsonResponse({
      success: true,
      data: jsonData,
      columnOrder: finalHeaders,
      displayMap: finalDisplayMap,
      alignMap: finalAlignMap,
      fullConfig: fullConfigForFrontend,
      availableTables: availableTables,
      masterFields: masterFields 
    });

  } catch (err) {
    return createJsonResponse({ success: false, message: err.toString() });
  }
}

function doPost(e) {
  try {
    const params = e.parameter;
    const action = params.action;
    let result;

    if (action === "registerDynamicDataTD") result = handleDynamicDataTD(params, "REGISTER");
    else if (action === "editDynamicDataTD") result = handleDynamicDataTD(params, "EDIT");
    else if (action === "deleteDynamicDataTD") result = handleDynamicDataTD(params, "DELETE"); // Nueva acción
    else throw new Error("Acción desconocida");

    syncAndGetMasterFields();
    return result;
  } catch (err) {
    return createJsonResponse({ success: false, message: err.toString() });
  }
}

/**
 * Utilidad para responder en formato JSON
 */
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

