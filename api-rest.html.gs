//https://script.google.com/home/projects/1wS7nZ8DVjQzir-jcwCmmKQ30zowrZyfcb1ssAuPCEAVu_Cu8trTqYIfa/exec

function doPost(e) {
  var request = JSON.parse(e);
  console.log(request);
}
function doGeta(e) {
  console.log(e)
  var request = JSON.parse(e);
  console.log(request);
}

var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Datos');

//https://script.google.com/macros/s/AKfycby3dTQEd-BmDzx_TIHZZXBuGo2ipYub672zxwHed3g/dev
function doGet(e) {
  console.log(e)
  var action = e.parameter.action;

  if (action = "getUsers") {
    return getUsers(e);

  }
  return JSON.stringify({message: "Hola Mundo"})
}

function getUsers(e) {

  var rows = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getValues;
  var data = [];

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var record = {};

    record['ID'] = row[0];
    record['Marca temporal'] = row[1];
    record['Dirección de correo electrónico'] = row[2];

    data.push(record);

  }
  var result = JSON.stringify(data);

  return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.JSON)

}
