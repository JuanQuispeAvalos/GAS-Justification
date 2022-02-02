//Configuración
const colFecCrea = 1;
const colEmail = 2
const colApellido = 4;
const colNombre = 5;
const colDni = 3;
const colCarrera = 6;
const colCiclo = 7;
const colUnidad = 9;
const colFecIna = 15;
const colMotivo = 16;
const colEvidencia = 17;
const colPdf = 19;
const colEnviado = 20;

//Identificaciones
var plantillaID = "1dMFGwJRv4E5Ba0bAiPA19I6FdS8ydoVeDK47E6HFrwM";
var pdfID = "1OWdc93-RCQOXL0JZEC2vQxVyGjEE00WU";
var tempID = "1HV3bYbwyAc66n6tXJNB-UXZWT8k8JxS1";

function generarPDFMasivo() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Datos');
  var alumnos = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getValues();
  alumnos.forEach(function (alumno, i) {
    var estudiante = {}
    estudiante.fecCrea = Utilities.formatDate(alumno[colFecCrea - 1], "GMT-5", "d 'de' MMMM 'del' yyyy");
    estudiante.email = alumno[colEmail - 1];
    estudiante.apellido = alumno[colApellido - 1];
    estudiante.nombre = alumno[colNombre - 1];
    estudiante.dni = alumno[colDni - 1];
    estudiante.carrera = alumno[colCarrera - 1];
    estudiante.ciclo = alumno[colCiclo - 1];
    estudiante.unidad = alumno[colUnidad - 1];
    estudiante.fecIna = Utilities.formatDate(alumno[colFecIna - 1], "GMT5", "dd/mm/yyyy");
    estudiante.motivo = alumno[colMotivo - 1];
    estudiante.evidencia = alumno[colEvidencia - 1];
    if (!alumno[colPdf - 1]) {
      var urls = generarPDF(estudiante);
      hoja.getRange(i + 2, colPdf).setValue(urls.urlPDF);
      estudiante.pdf = urls.urlPDF;
      enviarMail(estudiante);
      hoja.getRange(i + 2, colEnviado).setValue("Enviado");
    }
  })
}

function generarPDFIndividual() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Datos');
  var filaActiva = SpreadsheetApp.getActiveRange().getRow();
  var estudiante = {}
  estudiante.fecCrea = hoja.getRange(filaActiva, colFecCrea).getDisplayValue();
  estudiante.email = hoja.getRange(filaActiva, colEmail).getValue();
  estudiante.apellido = hoja.getRange(filaActiva, colApellido).getValue();
  estudiante.nombre = hoja.getRange(filaActiva, colNombre).getValue();
  estudiante.dni = hoja.getRange(filaActiva, colDni).getValue();
  estudiante.carrera = hoja.getRange(filaActiva, colCarrera).getValue();
  estudiante.ciclo = hoja.getRange(filaActiva, colCiclo).getValue();
  estudiante.unidad = hoja.getRange(filaActiva, colUnidad).getValue();
  estudiante.fecIna = hoja.getRange(filaActiva, colFecIna).getDisplayValue();
  estudiante.motivo = hoja.getRange(filaActiva, colMotivo).getValue();
  estudiante.evidencia = hoja.getRange(filaActiva, colEvidencia).getValue();
  var urls = generarPDF(estudiante);
  hoja.getRange(filaActiva, colPdf).setValue(urls.urlPDF);
  estudiante.pdf = urls.urlPDF;
  enviarMail(estudiante);
  hoja.getRange(filaActiva, colEnviado).setValue("Enviado");
}

function generarPDF(estudiante) {

  //Identificaciones
  var plantillaID = "1dMFGwJRv4E5Ba0bAiPA19I6FdS8ydoVeDK47E6HFrwM";
  var pdfID = "1OWdc93-RCQOXL0JZEC2vQxVyGjEE00WU";
  var tempID = "1HV3bYbwyAc66n6tXJNB-UXZWT8k8JxS1";

  //Conexiones
  var archivoPlantilla = DriveApp.getFileById(plantillaID);
  var carpetaPDF = DriveApp.getFolderById(pdfID);
  var carpetaDocs = DriveApp.getFolderById(tempID);
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Datos');

  //Crear Documento 
  var copiaArchivoPlantilla = archivoPlantilla.makeCopy(carpetaDocs);
  var copiaID = copiaArchivoPlantilla.getId();
  var nombreDoc = "JustificaciónInasistencia" + "_" + "####-YYYY"
  copiaArchivoPlantilla.setName(nombreDoc);
  var doc = DocumentApp.openById(copiaID);
  doc.setName(nombreDoc);

  //Remplazar variables
  doc.getBody().replaceText('{{last_name}}', estudiante.apellido);
  doc.getBody().replaceText('{{name}}', estudiante.nombre);
  doc.getBody().replaceText('{{ciclo}}', estudiante.ciclo);
  doc.getBody().replaceText('{{date}}', estudiante.fecIna);
  doc.getBody().replaceText('{{reason}}', estudiante.motivo);
  doc.getBody().replaceText('{{carrera}}', estudiante.carrera);
  doc.getBody().replaceText('{{unidad}}', estudiante.unidad);
  doc.getBody().replaceText('{{evidence}}', estudiante.evidencia);
  doc.getBody().replaceText('{{dni}}', estudiante.dni);
  doc.getBody().replaceText('{{creationdate}}', estudiante.fecCrea);

  doc.saveAndClose();

  const pdfBlob = copiaArchivoPlantilla.getAs(MimeType.PDF);
  var PDFcreado = carpetaPDF.createFile(pdfBlob);
  PDFcreado.addViewer(estudiante.email);
  var urls = {}
  urls.urlPDF = PDFcreado.getUrl();
  urls.urlDoc = doc.getUrl();
  return urls
}

function enviarMail(estudiante) {
  var mensaje = "Justificación de inasistencia a la clase de la Unidad Didáctica " +estudiante.unidad +" Descárgalo aquí " +estudiante.pdf
  MailApp.sendEmail(estudiante.email, "JUSTIFICACIÓN DE INASISTENCIA Nº ####-YYYY", mensaje)

}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu("Reportes");
  menu.addItem("Generar Justificación", "generarPDFIndividual")
    .addItem("Generar Justificación masiva", "generarPDFMasivo")
    .addToUi();
}
