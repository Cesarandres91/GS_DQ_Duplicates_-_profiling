function replicateWithDataType() {
  // Obtener el libro activo y la hoja 'Base'
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var baseSheet = ss.getSheetByName('Base');
  if (!baseSheet) {
    SpreadsheetApp.getUi().alert('La hoja "Base" no existe.');
    return;
  }

  // Crear o limpiar la hoja 'Datatype'
  var datatypeSheet = ss.getSheetByName('Datatype');
  if (!datatypeSheet) {
    datatypeSheet = ss.insertSheet('Datatype');
  } else {
    datatypeSheet.clear();
  }

  // Obtener todos los datos de la hoja 'Base'
  var data = baseSheet.getDataRange().getValues();
  var headers = data[0];
  var values = data.slice(1);
  var newHeaders = [];
  
  // Crear nuevas cabeceras con "_type" añadido
  headers.forEach(function(header) {
    newHeaders.push(header);
    newHeaders.push(header + '_type');
  });
  datatypeSheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);

  // Iterar sobre las filas y determinar los tipos de datos
  var newData = values.map(function(row) {
    var newRow = [];
    row.forEach(function(value, colIndex) {
      newRow.push(value);
      newRow.push(determineDataType(value));
    });
    return newRow;
  });

  // Escribir los datos nuevos en la hoja 'Datatype'
  datatypeSheet.getRange(2, 1, newData.length, newData[0].length).setValues(newData);
}

function determineDataType(value) {
  if (value === null || value === '') {
    return 'Null';
  }
  if (!isNaN(value)) {
    return 'Numeric';
  }
  if (/^[a-zA-Z0-9]+$/.test(value)) {
    return 'Alphanumeric';
  }
  if (typeof value === 'string') {
    return 'Text';
  }
  return 'Unknown';
}
