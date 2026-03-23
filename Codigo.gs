// Cuentas de Casa - Apps Script v3
// Al implementar: Acceso = "Cualquier usuario"

function doGet(e) {
  return respond(handleRequest(e));
}

function doPost(e) {
  return respond(handleRequest(e));
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleRequest(e) {
  try {
    var data = {};
    var action = '';
    var debugInfo = {receivedParams: null, receivedPost: null};

    // Leer payload
    if (e.parameter && e.parameter.payload) {
      debugInfo.receivedParams = 'payload param found, length=' + e.parameter.payload.length;
      try { data = JSON.parse(e.parameter.payload); } catch(x) {
        return {error: 'JSON parse error: ' + x.toString(), raw: e.parameter.payload.substring(0,100)};
      }
      action = data.action || '';
    } else if (e.postData && e.postData.contents) {
      debugInfo.receivedPost = 'postData found';
      try { data = JSON.parse(e.postData.contents); } catch(x) { data = {}; }
      action = data.action || '';
    } else if (e.parameter) {
      action = e.parameter.action || '';
      data = e.parameter;
      debugInfo.receivedParams = JSON.stringify(Object.keys(e.parameter));
    } else {
      return {error: 'No data received', debug: debugInfo};
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === 'getAll')          return getAllData(ss);
    if (action === 'addRegistro')     return addRow(ss, 'Registros', data.row);
    if (action === 'addEfectivo')     return addRow(ss, 'Efectivo', data.row);
    if (action === 'editRegistro')    return editRow(ss, 'Registros', data.original, data.updated);
    if (action === 'editEfectivo')    return editRow(ss, 'Efectivo', data.original, data.updated);
    if (action === 'deleteRegistro')  return deleteRow(ss, 'Registros', data.row);
    if (action === 'deleteEfectivo')  return deleteRow(ss, 'Efectivo', data.row);
    if (action === 'updateGastoFijo') return updateGastoFijo(ss, data);

    return {ok: true, ping: true, receivedAction: action, debug: debugInfo};
  } catch(err) {
    return {error: err.toString()};
  }
}

function getSheetData(sh) {
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return [];
  return sh.getRange(1, 1, lastRow, lastCol).getValues();
}

function getAllData(ss) {
  var regSh = getOrCreateSheet(ss, 'Registros',
    ['Fecha valor','Importe','Concepto','Entidad',
     'Nombre de producto','Tipo de producto','Tipo de movimiento','Tipo de gasto','Categoria']);
  var efSh = getOrCreateSheet(ss, 'Efectivo',
    ['Fecha','Concepto','Descripcion','Importe','Tipo']);
  var gfSh = getOrCreateSheet(ss, 'Gastos Fijos',
    ['Categoria','Concepto','Ultimo cobro','Importe','Frecuencia']);

  if (gfSh.getLastRow() < 2) {
    var gfData = [
      ['Agua','Recibo urbide arabako ur','2026-03-04',-49.31,'Bimestral'],
      ['Comunidad','Cuota comunidad de vecinos','2026-03-06',-48.76,'Mensual'],
      ['Comunidad','Trf. garajes barrio la llana','2026-03-06',-3.00,'Mensual'],
      ['Cuota Eroski','Rbo eroski s. coop.','2026-03-06',-12.99,'Mensual'],
      ['Gas','Recibo ned suministro gl','2026-03-12',-127.76,'Bimestral'],
      ['Hipoteca','Cuota ptmo 852117208-6','2026-02-23',-754.00,'Mensual'],
      ['Internet','Recibo xfera moviles s.','2026-03-06',-63.21,'Mensual'],
      ['Otros seguros','Rbo the phone house','2026-02-04',-16.00,'Mensual'],
      ['Otros seguros','Recibo occident gco','2026-03-05',-16.81,'Mensual'],
      ['Prestamos','Rbo rci banque s a.','2026-03-02',-311.15,'Mensual'],
      ['Seguro Vida','Rbo allianz vida','2026-01-02',-172.17,'Anual'],
      ['Seguro Vida','Recibo allianz hogar','2026-01-02',-579.26,'Anual'],
      ['Seguro auto','Recibo qualitas auto','2026-02-25',-694.59,'Anual'],
      ['Sindicato','L.a.b.','2026-03-02',-16.31,'Mensual'],
      ['Suministros','Recibo low cost power','2026-03-02',-96.85,'Mensual'],
      ['Suscripciones','Recibo paypal europe 1','2026-02-24',-6.99,'Mensual'],
      ['Suscripciones','Recibo paypal europe 2','2026-01-27',-9.90,'Mensual']
    ];
    for (var g = 0; g < gfData.length; g++) {
      gfSh.appendRow(gfData[g]);
    }
  }

  return {
    registros:   sheetToObjects(regSh),
    efectivo:    sheetToObjects(efSh),
    gastosFijos: sheetToObjects(gfSh)
  };
}

function getOrCreateSheet(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
  }
  return sh;
}

function sheetToObjects(sh) {
  var vals = getSheetData(sh);
  if (vals.length < 2) return [];
  var hdrs = vals[0];
  var result = [];
  for (var i = 1; i < vals.length; i++) {
    var row = vals[i];
    var allEmpty = true;
    for (var c = 0; c < row.length; c++) {
      if (row[c] !== '' && row[c] !== null && row[c] !== undefined) {
        allEmpty = false;
        break;
      }
    }
    if (allEmpty) continue;
    var obj = {};
    for (var h = 0; h < hdrs.length; h++) {
      obj[hdrs[h]] = row[h];
    }
    result.push(obj);
  }
  return result;
}

function toDateStr(v) {
  if (!v && v !== 0) return '';
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(v).trim();
  var parts;
  if (s.length === 10 && s.charAt(2) === '/' && s.charAt(5) === '/') {
    parts = s.split('/');
    return parts[2] + '-' + parts[1] + '-' + parts[0];
  }
  if (s.length >= 10 && s.charAt(4) === '-') {
    return s.substring(0, 10);
  }
  return s;
}

function addRow(ss, sheetName, row) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return {error: 'Sheet not found: ' + sheetName};

  var r = row;
  if (typeof r === 'string') {
    try { r = JSON.parse(r); } catch(x) {}
  }
  if (!Array.isArray(r)) return {error: 'row must be array, got: ' + typeof r};

  var processed = [];
  for (var i = 0; i < r.length; i++) {
    var v = r[i];
    if (i === 0 && typeof v === 'string' && v.length >= 10 && v.charAt(4) === '-') {
      var p = v.split('-');
      processed.push(new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10)));
    } else {
      processed.push(v);
    }
  }
  sh.appendRow(processed);
  return {ok: true};
}

function findRowIndex(vals, key) {
  var keyFecha = toDateStr(key[0]);
  var keyImp   = parseFloat(key[1]);
  var keyConc  = String(key[2] || '').trim().toLowerCase();

  var debugRows = [];
  for (var i = 1; i < vals.length; i++) {
    var row = vals[i];
    var allEmpty = true;
    for (var c = 0; c < row.length; c++) {
      if (row[c] !== '' && row[c] !== null) { allEmpty = false; break; }
    }
    if (allEmpty) continue;

    var rowFecha = toDateStr(row[0]);
    var rowImp   = parseFloat(row[1]);
    var rowConc  = String(row[2] || '').trim().toLowerCase();

    if (debugRows.length < 3) {
      debugRows.push({rowFecha: rowFecha, rowImp: rowImp, rowConc: rowConc.substring(0,20)});
    }

    if (rowFecha === keyFecha && Math.abs(rowImp - keyImp) < 0.005 && rowConc === keyConc) {
      return i;
    }
  }
  // Return debug info so we can see what went wrong
  return {notFound: -1, looking: {keyFecha: keyFecha, keyImp: keyImp, keyConc: keyConc.substring(0,20)}, firstRows: debugRows};
}

function editRow(ss, sheetName, original, updated) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return {error: 'Sheet not found: ' + sheetName};

  if (typeof original === 'string') original = JSON.parse(original);
  if (typeof updated === 'string')  updated  = JSON.parse(updated);

  var vals = getSheetData(sh);
  var idxResult = findRowIndex(vals, original);
  if (typeof idxResult === 'object') return {error: 'Row not found', debug: idxResult};

  var processed = [];
  for (var i = 0; i < updated.length; i++) {
    var v = updated[i];
    if (i === 0 && typeof v === 'string' && v.length >= 10 && v.charAt(4) === '-') {
      var p = v.split('-');
      processed.push(new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10)));
    } else {
      processed.push(v);
    }
  }

  sh.getRange(idxResult + 1, 1, 1, processed.length).setValues([processed]);
  return {ok: true, row: idxResult + 1};
}

function deleteRow(ss, sheetName, rowData) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return {error: 'Sheet not found: ' + sheetName};

  if (typeof rowData === 'string') rowData = JSON.parse(rowData);

  var vals = getSheetData(sh);
  var idxResult = findRowIndex(vals, rowData);
  if (typeof idxResult === 'object') return {error: 'Row not found', debug: idxResult};

  sh.deleteRow(idxResult + 1);
  return {ok: true};
}

function updateGastoFijo(ss, data) {
  var sh = ss.getSheetByName('Gastos Fijos');
  if (!sh) return {error: 'Sheet not found'};

  var keyword = String(data.keyword || '').toLowerCase();
  var vals    = getSheetData(sh);

  for (var i = 1; i < vals.length; i++) {
    var conc = String(vals[i][1] || '').toLowerCase();
    if (keyword && conc.indexOf(keyword) >= 0) {
      var parts = String(data.fecha || '').split('-');
      if (parts.length === 3) {
        var d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
        sh.getRange(i + 1, 3).setValue(d).setNumberFormat('yyyy-MM-dd');
        sh.getRange(i + 1, 4).setValue(parseFloat(data.importe));
      }
    }
  }
  return {ok: true};
}