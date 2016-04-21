/*

Microservicio implementado mediante una App de Google Script, consistente en la obtención de los datos de una Hoja de cálculo
de Google (o libro) en formato Json.

Parámetros url:
---------------

@param id {string} Identificador de la hoja de cálculo (Obligatorio)
@param sheet {string} Nombre de la hoja del libro que se quiere obtener. (Opcional) Ejm: 'Hoja 1'
@param cell {string}  Identificador de celda. (Opcional) Ejem: 'A1:B2'

*/


function doGet(request) {
  
  // Id de spreadsheet (Obligatorio)
  var id = request.parameters.id;
  
  // Nombre de hoja (Opcional)
  var sheet = request.parameters.sheet;
  
  // Celda de donde extraer datos. (Opcional)
  var cell = request.parameters.cell;

  var output = ContentService.createTextOutput();
  var data = {};  
  var ss = SpreadsheetApp.openById(id);
  
  if (sheet) {
    if (cell) {
      // Si se ha pasado un parametro que define una celda obtiene el valor que contiene dicha celda
      data = ss.getSheetByName(sheet).getRange(cell).getValue();
    } else {
      // Si no se proporciona ningun parametro de celda se bbtienen todos los valores de la hoja
      data[sheet] = readData_(ss, sheet);
    }
  } else {
    // Si no se especifican ni hoja ni celda, obtiene todas las hojas del libro, salvo las que comienzan con '_'
    ss.getSheets().forEach(function(oSheet, iIndex) {
      var sName = oSheet.getName();
      if (! sName.match(/^_/)) {
        data[sName] = readData_(ss, sName);
      }
    })
  }
  var result = cell ? data : JSON.stringify(data);
  // En caso de que se quiera usar Jsonp:
  //var callback = request.parameters.callback;
  var callback = 'jsonpCallback';
  if (callback == undefined) {
    output.setContent(result);
    output.setMimeType(cell ? ContentService.MimeType.TEXT : ContentService.MimeType.JSON);
  }
  else {
    output.setContent(callback + "(" + result + ")");
    output.setContent(result);
    output.setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return output;
}


function readData_(ss, sheetname, properties) {
  
  if (typeof properties == "undefined") {
    properties = getHeaderRow_(ss, sheetname);
    properties = properties.map(function(p) { return p.replace(/\s+/g, '_'); });
  }
  
  var rows = getDataRows_(ss, sheetname);
  var data = [];
  for (var r = 0, l = rows.length; r < l; r++) {
    var row = rows[r];
    var record = {};
    for (var p in properties) {
      record[properties[p]] = convert_(row[p]);
    }
    data.push(record);
  }
  return data;
}


function convert_(value) {
  if (value === "true") return true;
  if (value === "false") return false;
  return value;
}


function getDataRows_(ss, sheetname) {
  
  var sh = ss.getSheetByName(sheetname);
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
}


function getHeaderRow_(ss, sheetname) {
  
  var sh = ss.getSheetByName(sheetname);
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
}
  
function test(){
  
  var parameters = {id:'1qZZMIcp-tzTEdaY4-jiIGaix8fzfJ5_FNx4yN4NqwA4', sheet: 'Apps'};
  var request = {parameters: parameters};
  var result = doGet(request);
}
