// Función para mostrar un cuadro de diálogo que permite al usuario ingresar el ID del archivo CSV
function showPrompt() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Convertir CSV a Google Sheets', 'Por favor ingresa el ID del archivo CSV:', ui.ButtonSet.OK_CANCEL);

  // Procesar la respuesta del usuario
  if (response.getSelectedButton() == ui.Button.OK) {
    var fileId = response.getResponseText();
    try {
      convertCsvToGoogleSheet(fileId);
    } catch (e) {
      ui.alert('Ocurrió un error: ' + e.message);
    }
  } else {
    ui.alert('Operación cancelada.');
  }
}




function convertCsvToGoogleSheet(fileId) {
  try {
    // Obtener el archivo CSV por ID
    var file = DriveApp.getFileById(fileId);
    var csvData = Utilities.parseCsv(file.getBlob().getDataAsString());

    // Crear una nueva hoja de cálculo de Google Sheets
    var newSheet = SpreadsheetApp.create(file.getName().replace('.csv', ''));
    var sheet = newSheet.getActiveSheet();

    // Escribir los datos CSV en la nueva hoja de cálculo
    sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);

    // Obtener el ID y URL de la nueva hoja de cálculo
    var newSheetId = newSheet.getId();
    var newSheetUrl = 'https://docs.google.com/spreadsheets/d/' + newSheetId;
    Logger.log('Spreadsheet ID: ' + newSheetId);

    showSuccessMessage(newSheetUrl, newSheet.getUrl(), newSheet.getName());
    return newSheet.getName();
  } catch (e) {
    Logger.log('Error: ' + e.message);
    throw e;
  }
}



function showSuccessMessage(sheetUrl, fileUrl, fileName) {
  var htmlOutput = HtmlService.createHtmlOutput(
    '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '<base target="_top">' +
    '<style>' +
    'body { font-family: Arial, sans-serif; text-align: center; }' +
    '.success-container { padding: 20px; border: 2px solid #4CAF50; background-color: #f9fff9; border-radius: 10px; margin-top: 20px; }' +
    '.success-icon { font-size: 50px; color: #4CAF50; }' +
    '.file-link { display: block; margin-top: 20px; text-decoration: none; color: #4CAF50; font-size: 18px; }' +
    '.file-link:hover { text-decoration: underline; }' +
    '</style>' +
    '</head>' +
    '<body>' +
    '<div class="success-container">' +
    '<div class="success-icon">✔️</div>' +
    '<h2>Archivo convertido con éxito</h2>' +
    '<p>El archivo ha sido convertido y guardado como:</p>' +
    '<a id="sheetLink" class="file-link" target="_blank">Abrir archivo Google Sheets</a>' +
    '</div>' +
    '<script>' +
    'document.getElementById("sheetLink").href="' + sheetUrl + '";' +
    '</script>' +
    '</body>' +
    '</html>'
  ).setWidth(400).setHeight(300);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Archivo Convertido');
}
