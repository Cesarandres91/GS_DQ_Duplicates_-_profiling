function findDuplicates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseSheet = ss.getSheetByName('Base');
  const duplicatesSheet = ss.getSheetByName('Duplicados');
  
  if (!baseSheet || !duplicatesSheet) {
    SpreadsheetApp.getUi().alert('Las hojas "Base" o "Duplicados" no existen.');
    return;
  }

  // Limpiar la hoja Duplicados
  duplicatesSheet.clear();
  duplicatesSheet.getRange('A1:C1').setValues([['Field', 'Value', 'Count']]);

  const data = baseSheet.getDataRange().getValues();
  const headers = data[0];
  const values = data.slice(1);
  const result = [];

  headers.forEach((header, colIndex) => {
    const columnValues = values.map(row => row[colIndex]);
    const counts = columnValues.reduce((acc, value) => {
      if (value) {
        acc[value] = (acc[value] || 0) + 1;
      }
      return acc;
    }, {});

    for (const [key, count] of Object.entries(counts)) {
      if (count >= 2) {
        result.push([header, key, count]);
      }
    }
  });

  // Ordenar resultados
  result.sort((a, b) => {
    if (a[0] === b[0]) {
      if (b[2] === a[2]) {
        return a[1].localeCompare(b[1]);
      }
      return b[2] - a[2];
    }
    return a[0].localeCompare(b[0]);
  });

  // Escribir resultados en la hoja Duplicados
  if (result.length > 0) {
    duplicatesSheet.getRange(2, 1, result.length, result[0].length).setValues(result);
  }
}
