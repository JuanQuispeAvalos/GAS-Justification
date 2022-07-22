const header = {
  id: colId - 1,
  email: colEmail - 1,
  documentId: colDni - 1,
  lastName: colApellido - 1,
  firstName: colNombre - 1,
  career: colCarrera - 1,
  semester: colCiclo - 1,
  date: colFecIna - 1,
  motive: colMotivo - 1,
  evidencyType: colEvidencia - 1,
  evidency: colUrlEvidencia - 1,
  report: colPdf - 1,
  satus: colEnviado - 1,
  course: colUnidad - 1,
  recovery: colRecuperacion - 1,
  responsibles: colResponsable - 1,
};

class DataItem {
  constructor({ id, email, documentId, lastName, firstName, career, semester, date, motive, evidencyType, evidency, report, isSent, course, recovery, responsibles }) {
    this.id = id;
    this.email = email;
    this.documentId = documentId;
    this.lastName = lastName;
    this.firstName = firstName;
    this.career = career;
    this.semester = semester;
    this.date = date;
    this.motive = motive;
    this.evidencyType = evidencyType;
    this.evidency = evidency;
    this.report = report;
    this.isSent = isSent;
    this.course = course ? String(course).split(",") : null;
    this.recovery = recovery;
    this.responsibles = responsibles ? String(responsibles).split(",") : null;
  }
}

function MERGE_COLUMNS(datos = [[]], separator = ",") {
  const result = datos.map(row => {
    return row
      .filter(col => col)
      .join(separator);
  })
  return result;
}

function CLEAN_DATA(sheetName, range) {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadSheet.getSheetByName(sheetName);
  const values = sheet.getRange(range).getValues();
  const cells = values.reduce((store, row, pos) => {
    const item = mapToDataItem(row);
    if (pos == 0) store.push(Object.keys(item).map(key => key.toUpperCase()))
    if (item.responsibles) item.responsibles.forEach((responsible, index) => {
      store.push(
        mapToArray({
          ...item,
          responsibles: responsible,
          course: `${item.course[index]}`.trim()
        })
      )
    })
    return store
  }, [])
  return cells;
}

function mapToArray(item = new DataItem()) {
  return Object.values(item)
}

function mapToDataItem(row = new Array()) {
  return new DataItem({
    id: row[header.id],
    email: row[header.email],
    documentId: row[header.documentId],
    lastName: row[header.lastName],
    firstName: row[header.firstName],
    career: row[header.career],
    semester: row[header.semester],
    date: row[header.date],
    motive: row[header.motive],
    evidencyType: row[header.evidencyType],
    evidency: row[header.evidency],
    report: row[header.report],
    isSent: row[header.satus],
    course: row[header.course],
    recovery: row[header.recovery],
    responsibles: row[header.responsibles],
  });
}

function main() {
  console.log(CLEAN_DATA("Datos", "A2:AE10"));
}
