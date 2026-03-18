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

    const usuarioId = String(e.parameter.usuarioId || "").trim();
    const perfil = String(e.parameter.perfil || "").trim();
    const isAdmin = (perfil.toLowerCase() === "admin");

    if (action === "getTableFriendlyNames") {
      const result = getTableFriendlyNames(appTienda || userTienda);
      return createJsonResponse(result);
    }

    if (action === "getDashboardSummary") {
      const ssData = SpreadsheetApp.openById(DATA_SS_ID);
      const timeZone = ssData.getSpreadsheetTimeZone();
      
      // --- 1. TD101_MAIN (OPERACIONES) ---
      const sheet101 = ssData.getSheetByName("TD101_MAIN");
      const data101 = sheet101 ? sheet101.getDataRange().getValues() : [];
      let resumenOps = [];

      if (data101.length > 1) {
        const headers101 = data101[0];
        
        // Mapeo robusto de índices
        const idx = {
          fecha:  headers101.findIndex(h => h.toUpperCase().includes("FECHA")),
          monto:  headers101.findIndex(h => h.toUpperCase().includes("IMPORTE")),
          mov:    headers101.findIndex(h => h.toUpperCase().includes("MOVIMIENTO")),
          type:   headers101.findIndex(h => h.toUpperCase().includes("TYPEREG")),
          cuenta: headers101.findIndex(h => h.toUpperCase().includes("CUENTA") && !h.toUpperCase().includes("ALIAS")),
          alias:  headers101.findIndex(h => h.toUpperCase().includes("ALIASCUENTA")),
          owner:  headers101.findIndex(h => h.toUpperCase().includes("IDOWNER"))
        };

        resumenOps = data101.slice(1)
          .filter(row => {
            const rowOwner = String(row[idx.owner] || "").trim();
            return isAdmin || rowOwner === usuarioId;
          })
          .map(row => {
            // Normalización de Fecha
            let fechaVal = row[idx.fecha];
            let fechaStr = (fechaVal instanceof Date) 
              ? Utilities.formatDate(fechaVal, timeZone, "dd/MM/yyyy") 
              : String(fechaVal || "");

            // Normalización de Importe (Asegurar que sea número para el Dashboard)
            let importeRaw = row[idx.monto];
            let importeNum = (typeof importeRaw === 'number') ? importeRaw : parseFloat(String(importeRaw).replace(/[^\d.-]/g, '')) || 0;

            return {
              FECHA: fechaStr,
              IMPORTE: importeNum,
              MOVIMIENTO: String(row[idx.mov] || "").trim().toUpperCase(),
              TYPEREG: String(row[idx.type] || "").trim().toUpperCase(),
              CUENTA: String(row[idx.cuenta] || "").trim(),
              ALIAS_CUENTA: idx.alias !== -1 ? String(row[idx.alias] || "").trim() : "Sin Alias"
            };
          });
      }

      // --- 2. TD102_CUENTAS (MAESTRO) ---
      const sheet102 = ssData.getSheetByName("TD102_CUENTAS");
      const data102 = sheet102 ? sheet102.getDataRange().getValues() : [];
      let resumenCuentas = [];

      if (data102.length > 1) {
        const headers102 = data102[0];
        
        const idxCta = {
          id:     headers102.findIndex(h => h.toUpperCase().endsWith("ID")), 
          banco:  headers102.findIndex(h => h.toUpperCase().includes("BANCO")),
          tipo:   headers102.findIndex(h => h.toUpperCase().includes("TIPO")),
          moneda: headers102.findIndex(h => h.toUpperCase().includes("MONEDA")),
          num:    headers102.findIndex(h => h.toUpperCase().includes("NUMERO")),
          owner:  headers102.findIndex(h => h.toUpperCase().includes("IDOWNER"))
        };

        resumenCuentas = data102.slice(1)
          .filter(row => {
            const rowOwner = String(row[idxCta.owner] || "").trim();
            return isAdmin || rowOwner === usuarioId;
          })
          .map(row => {
            const banco  = String(row[idxCta.banco] || '').trim();
            const tipo   = String(row[idxCta.tipo] || '').trim();
            const moneda = String(row[idxCta.moneda] || '').trim();
            
            return {
              ID: String(row[idxCta.id] || "").trim(),
              NUMERO: String(row[idxCta.num] || "").trim(),
              ALIAS: `${banco} ${tipo} ${moneda}`.replace(/\s+/g, ' ').trim() || "Cuenta sin nombre",
              MONEDA: moneda
            };
          });
      }

      return createJsonResponse({
        success: true,
        operaciones: resumenOps,
        cuentas: resumenCuentas
      });
    }

    const tableName = e.parameter.tableName || e.parameter.sheet;
    if (!tableName) throw new Error("Parámetro 'tableName' o 'action' omitido.");

    const ignoreVisibility = e.parameter.ignoreVisibility === 'true'; 

    const ssData = SpreadsheetApp.openById(DATA_SS_ID);
    const ssConfig = SpreadsheetApp.openById(CONFIG_SS_ID);
    
    const configSheet = ssConfig.getSheetByName(CONFIG_SHEET_NAME);
    const dataSheet = ssData.getSheetByName(tableName);

    if (!dataSheet) throw new Error("La tabla '" + tableName + "' no existe.");
    if (!configSheet) throw new Error("La hoja de configuración no existe.");

    const configRows = configSheet.getDataRange().getValues();
    const configData = configRows.slice(1);
    const configMap = {};
    const availableTables = [];
    const fullConfigForFrontend = [];

    configData.forEach(row => {
        const rowAppTienda = String(row[0]).trim();
        const nombreTabla = String(row[1]).trim();
        if (rowAppTienda === userTienda && !availableTables.includes(nombreTabla)) {
          availableTables.push(nombreTabla);
        }
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

    const fullData = dataSheet.getDataRange().getValues();
    if (fullData.length === 0) throw new Error("La tabla está vacía.");
    
    const originalHeaders = fullData[0];
    const tablePrefix = tableName.split('_')[0].toUpperCase();
    const colOwnerIdx = originalHeaders.indexOf(`${tablePrefix}IDOWNER`);

    let rowsToProcess = fullData.slice(1);
    if (!isAdmin && colOwnerIdx !== -1) {
      rowsToProcess = rowsToProcess.filter(row => String(row[colOwnerIdx]).trim() === usuarioId);
    }

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
      const isOwnerField = upperH === `${tablePrefix}IDOWNER`;
      
      if (ignoreVisibility || (config && config.Visible_Encabezado !== "") || isPK || isAuditField || isTypeReg || isOwnerField) {
        finalHeaders.push(cleanH);
        finalDisplayMap[cleanH] = (config && config.Nombre_Encabezado) ? config.Nombre_Encabezado : cleanH;
        finalAlignMap[cleanH] = (config && config.Justificado_Campo) ? config.Justificado_Campo : 'left';
        colIndexesToFetch.push(index);
      }
    });

    const jsonData = rowsToProcess.map(row => {
      const obj = {};
      colIndexesToFetch.forEach((colIdx, i) => { obj[finalHeaders[i]] = row[colIdx]; });
      return obj;
    });

    const masterFields = typeof syncAndGetMasterFields === "function" ? syncAndGetMasterFields(ssData) : []; 

    const dictionaryData = [];
    try {
      const dictSheet = ssData.getSheetByName("TV002_DICCIONARIO");
      if (dictSheet) {
        const dictValues = dictSheet.getDataRange().getValues();
        const dictHeaders = dictValues[0];
        const colDictOwnerIdx = dictHeaders.indexOf("Usuario_Tienda_ID");
        const dictRows = dictValues.slice(1);

        dictRows.forEach(row => {
          const dictAppTienda = String(row[0]).trim();
          const dictOwnerId = String(row[colDictOwnerIdx]).trim();
          const matchesTienda = (dictAppTienda === userTienda || dictAppTienda === "ALL");
          const ownerIdUpper = dictOwnerId.toUpperCase();
          const matchesOwner = (isAdmin || dictOwnerId === usuarioId || ownerIdUpper === "ALL");

          if (matchesTienda && matchesOwner) {
            dictionaryData.push({
              Nombre_Tabla: String(row[1]).trim(),
              Encabezado_Tabla: String(row[2]).trim(),
              Valores_Encabezado: String(row[3] || "[]").trim(),
              Usuario_Tienda_ID: dictOwnerId
            });
          }
        });
      }
    } catch (dictErr) {
      console.error("Error cargando diccionario: " + dictErr.toString());
    }

    return createJsonResponse({
      success: true,
      data: jsonData,
      columnOrder: finalHeaders,
      displayMap: finalDisplayMap,
      alignMap: finalAlignMap,
      fullConfig: fullConfigForFrontend,
      availableTables: availableTables,
      masterFields: masterFields,
      dictionary: dictionaryData 
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
        } else if (action === "deleteTransferCascade") { // <--- NUEVA ACCIÓN
          result = handleDeleteTransferCascade(params);
        } else if (action === "actualizarValorEnDiccionario") {
          result = ejecutarActualizacionDiccionario(params);
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



