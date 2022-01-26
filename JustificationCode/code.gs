function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('AutoFill Docs');
  menu.addItem('Create New Docs', 'createNewGoogleDocs')
  menu.addToUi();
}

function createNewGoogleDocs() {
  //Este valor debe ser la identificación de su plantilla de documento 
  const googleDocTemplate = DriveApp.getFileById('1mtCYqcxGglV8Uh9GhkbmcxtVrmWZKHAzhyuu404W82w');
  
  //Este valor debe ser la identificación de la carpeta donde se almacenaran los documentos
  const destinationFolder = DriveApp.getFolderById('1-c3MxmNQupCo-i9Ldv6Ux1m1MlxxZm0C');

  //Este valor debe ser la identificación de la carpeta donde se almacenaran los PDFs
  const detionationFolderPdfs = DriveApp.getFolderById('1MPytBwARxz2U6K3TqYM3bZaWQd0KPz-a');

  //aqui guardamos la hoja como variable 
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Data')
  
  //Ahora obtenemos todos los valores como una matriz
  const rows = sheet.getDataRange().getValues();
  
  //Empezamos a procesar cada fila de la hoja de cálculo 
  rows.forEach(function(row, index){
    //Aquí verificamos si esta fila son los encabezados, si es así lo saltamos 
    if (index === 0) return;

    //Aquí comprobamos si ya se ha generado un documento mirando 'Enlace del documento', si es así lo omitimos 
    if (row[19]) return;

    //Usando los datos de la fila en un literal de plantilla, hacemos una copia de nuestro documento de plantilla en nuestra    Carpeta de destino 
    const copy = googleDocTemplate.makeCopy(`${row[3]}, ${row[4]} Justification` , destinationFolder)

    //Una vez que tenemos la copia, la abrimos usando DocumentApp 
    const doc = DocumentApp.openById(copy.getId())

    //Todo el contenido vive en el cuerpo, así que lo obtenemos para editarlo 
    const body = doc.getBody();

    //En esta línea, hacemos un formato de fecha 
    const friendlyDate = new Date(row[6]).toLocaleDateString();
    
    //En estas líneas, reemplazamos nuestros tokens de reemplazo con valores de nuestra fila de hoja de cálculo 
    body.replaceText('{{last_name}}', row[3]);
    body.replaceText('{{name}}', row[4]);
    body.replaceText('{{ciclo}}', row[5]);
    body.replaceText('{{date}}', friendlyDate);
    body.replaceText('{{reason}}', row[9]);
    body.replaceText('{{dni}}', row[2]);
    
    //Hacemos permanentes nuestros cambios guardando y cerrando el documento 
    doc.saveAndClose();

    //Almacenar la url de nuestro nuevo documento en una variable 
    const url = doc.getUrl();

    //Escriba ese valor de nuevo en la columna 'Enlace del documento' en la hoja de cálculo. 
    sheet.getRange(index + 1, 20).setValue(url)
    
  })
  
}