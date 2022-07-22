//Configuración
const colId = 1;
const colFecCrea = 2;
const colEmail = 3;
const colApellido = 5;
const colNombre = 6;
const colDni = 4;
const colCarrera = 7;
const colCiclo = 8;
const colUnidad = 29;
const colFecIna = 16;
const colMotivo = 17;
const colEvidencia = 18;
const colUrlEvidencia = 19;
const colPdf = 26;
const colEnviado = 27;
const colResponsable = 31;
const colRecuperacion = 30;

//Identificaciones
var plantillaID = "1dMFGwJRv4E5Ba0bAiPA19I6FdS8ydoVeDK47E6HFrwM";
var pdfID = "1OWdc93-RCQOXL0JZEC2vQxVyGjEE00WU";
var tempID = "1HV3bYbwyAc66n6tXJNB-UXZWT8k8JxS1";

function generarPDFMasivo() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Datos');
  var alumnos = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getValues();
  alumnos.forEach(function (alumno, i) {
    var estudiante = {}
    estudiante.id = alumno[colId - 1];
    estudiante.fecCrea = Utilities.formatDate(alumno[colFecCrea - 1], "GMT-5", "d 'de' MMMM 'del' yyyy");
    estudiante.email = alumno[colEmail - 1];
    estudiante.responsable = alumno[colResponsable - 1];
    estudiante.urlevidencia = alumno[colUrlEvidencia - 1];
    estudiante.recuperacion = alumno[colRecuperacion - 1];
    estudiante.apellido = alumno[colApellido - 1];
    estudiante.nombre = alumno[colNombre - 1];
    estudiante.dni = alumno[colDni - 1];
    estudiante.carrera = alumno[colCarrera - 1];
    estudiante.ciclo = alumno[colCiclo - 1];
    estudiante.unidad = alumno[colUnidad - 1];
    estudiante.fecIna = Utilities.formatDate(alumno[colFecIna - 1], "GMT5", "dd/MM/YYYY");
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
  estudiante.id = hoja.getRange(filaActiva, colId).getValue();
  estudiante.fecCrea = hoja.getRange(filaActiva, colFecCrea).getDisplayValue();
  estudiante.email = hoja.getRange(filaActiva, colEmail).getValue();
  estudiante.responsable = hoja.getRange(filaActiva, colResponsable).getValue();
  estudiante.urlevidencia = hoja.getRange(filaActiva, colUrlEvidencia).getValue();
  estudiante.recuperacion = hoja.getRange(filaActiva, colRecuperacion).getValue();
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

  //Crear Documento 

  var now = new Date
  var year = now.getFullYear();
  var copiaArchivoPlantilla = archivoPlantilla.makeCopy(carpetaDocs);
  var copiaID = copiaArchivoPlantilla.getId();
  var nombreDoc = "JustificaciónInasistencia" + "_" + estudiante.id + "-" + year
  copiaArchivoPlantilla.setName(nombreDoc);
  var doc = DocumentApp.openById(copiaID);
  doc.setName(nombreDoc);

  //Remplazar variables

  doc.getBody().replaceText('{{id}}', estudiante.id);
  doc.getBody().replaceText('{{fecid}}', year);
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
  if (estudiante.recuperacion == "SI") {
    doc.getBody().replaceText('{{recover}}', "(R)");
  } else {
    doc.getBody().replaceText('{{recover}}', " ");
  }

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
  var correo = estudiante.email + "," + estudiante.responsable;
  var now = new Date
  var year = now.getFullYear();

  var textoHtml = HtmlService.createHtmlOutputFromFile("correo").getContent();
  textoHtml = textoHtml.replace("{{estudiante.pdf}}", estudiante.pdf)
  textoHtml = textoHtml.replace("{{estudiante.evidencia}}", estudiante.urlevidencia)
  textoHtml = textoHtml.replace("{{estudiante.id}}", estudiante.id)
  textoHtml = textoHtml.replace("{{year}}", year)
  
  var mensaje = "Justificación de inasistencia a la clase de la Unidad Didáctica " + estudiante.unidad + " Descárgalo aquí " + estudiante.pdf
  MailApp.sendEmail({
    to: correo,
    subject: "JUSTIFICACIÓN DE INASISTENCIA Nº " + estudiante.id + "-" + year,
    htmlBody: textoHtml
  })
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu("Reportes");
  menu.addItem("Generar Justificación", "generarPDFIndividual")
    .addItem("Generar Justificación masiva", "generarPDFMasivo")
    .addToUi();
}
