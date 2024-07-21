function copiar_hoja_en_nuevo_libro() {
  var sourceSpreadsheetId = 'ID' 
  //Reemplaza esto con el ID del libro de Google Sheets

  var sourceSheetName = 'Base'; 
  // Reemplaza esto con el nombre de la hoja en el libro de Google 

  var targetSheetName = sourceSheetName ; 
  // Reemplaza esto con el nombre de la hoja en el libro asociado a este script

  copySheetData(sourceSpreadsheetId, sourceSheetName, targetSheetName);
  
}

function copySheetData(sourceSpreadsheetId, sourceSheetName, targetSheetName) {
  // Obtener el libro de origen y la hoja de origen
  var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
  
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert('La hoja "' + sourceSheetName + '" no existe en el libro de Google Sheets con ID: ' + sourceSpreadsheetId);
    return;
  }
  
  // Obtener todos los datos de la hoja de origen
  var data = sourceSheet.getDataRange().getValues();
  
  // Obtener el libro actual (el asociado a este script) y la hoja destino
  var targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);
  
  // Si la hoja destino ya existe, eliminarla
  if (targetSheet) {
    targetSpreadsheet.deleteSheet(targetSheet);
  }
  
  // Crear una nueva hoja con el nombre de la hoja de origen
  targetSheet = targetSpreadsheet.insertSheet(targetSheetName);
  
  // Pegar los datos en la hoja nueva
  targetSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  
  SpreadsheetApp.getUi().alert('Los datos de la hoja "' + sourceSheetName + '" del libro de Google Sheets con ID: ' + sourceSpreadsheetId + ' han sido copiados a una nueva hoja "' + targetSheetName + '" en este libro.');
}
