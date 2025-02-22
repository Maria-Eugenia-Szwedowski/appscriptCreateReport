// Programa realizado por María Eugenia Szwedowski - Cel +598 91074131
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('MENU PERSONAL')
    .addItem('Buscar Datos', 'buscador')
    .addToUi();
}


function buscador() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('buscador')
      .setTitle('Buscar Datos')
      .setWidth(300);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Buscar Datos');
}


function generateFilesFromTemplate2() {
  var TEMPLATE_FILE_ID = '________________FILE_ID________';
  var DESTINATION_FOLDER_ID = '____________FOLDER_ID____________';
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var formValues = dataRange.getValues();
  var docTemplate = DocumentApp.openById(TEMPLATE_FILE_ID);
  var folder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
  var headers = formValues[0];
 
  for (var rowIndex = 1; rowIndex < formValues.length; rowIndex++) {
    var formRow = formValues[rowIndex];
    var fileName = formRow[2]; // Ubicar el Unique ID para que no se repitan los nombres de los archivos. Acá esta en columna 3
   
    // Crear una copia del documento de plantilla
    var copy = DriveApp.getFileById(docTemplate.getId()).makeCopy();
    var copyDoc = DocumentApp.openById(copy.getId());
    var copyDocBody = copyDoc.getBody();
   
    // Reemplazar los marcadores en el documento con los datos de la fila
    for (var columnIndex = 0; columnIndex < formRow.length; columnIndex++) {
      var placeholder = "<<" + headers[columnIndex] + ">>";
      var value = formRow[columnIndex];
      copyDocBody.replaceText(placeholder, value);
    }
   
    // Guardar y cerrar el documento de copia
    copyDoc.saveAndClose();
   
    // Crear un archivo Docs a partir de la copia
    var newFileDoc = copy.setName(fileName + ".doc");
   
    // Crear un archivo PDF a partir de la copia y convertirlo a PDF
    var newFilePdf = copy.getAs("application/pdf").setName(fileName + ".pdf");
   
    // Agregar los archivos a la carpeta de destino
    //folder.createFile(newFileDoc);
    folder.createFile(newFilePdf);
   
    // Mover la copia del documento a la papelera
    //DriveApp.getFileById(copy.getId()).setTrashed(true);
  }
}




function getDataList() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getRange(2, 3, sheet.getLastRow() - 1); // Tercera columna sin encabezados
  var dataValues = dataRange.getValues();
  var dataList = dataValues.map(function(row) {
    return row[0];
  });
  return dataList;
}
