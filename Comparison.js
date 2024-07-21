function compareSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet1 = ss.getSheetByName('Base');
  const sheet2 = ss.getSheetByName('Base2');
  const comparisonSheet = ss.getSheetByName('Comparison') || ss.insertSheet('Comparison');
  
  if (!sheet1 || !sheet2) {
    SpreadsheetApp.getUi().alert('Las hojas "Base" o "Base2" no existen.');
    return;
  }

  // Limpiar la hoja Comparison
  comparisonSheet.clear();
  comparisonSheet.getRange('A1:D1').setValues([['Status', 'Variable', 'Old Value', 'New Value']]);

  const data1 = sheet1.getDataRange().getValues();
  const data2 = sheet2.getDataRange().getValues();
  const headers = data1[0];
  
  const values1 = data1.slice(1);
  const values2 = data2.slice(1);

  const map1 = createDataMap(values1, headers);
  const map2 = createDataMap(values2, headers);

  const comparison = [];

  // Comparar data1 con data2
  for (const [key, value] of Object.entries(map1)) {
    if (map2[key]) {
      if (JSON.stringify(value) !== JSON.stringify(map2[key])) {
        comparison.push(['Changed', key, JSON.stringify(value), JSON.stringify(map2[key])]);
      }
    } else {
      comparison.push(['Deleted', key, JSON.stringify(value), null]);
    }
  }

  // Identificar nuevos datos en data2
  for (const [key, value] of Object.entries(map2)) {
    if (!map1[key]) {
      comparison.push(['New', key, null, JSON.stringify(value)]);
    }
  }

  // Escribir resultados en la hoja Comparison
  if (comparison.length > 0) {
    comparisonSheet.getRange(2, 1, comparison.length, comparison[0].length).setValues(comparison);
  }
}

function createDataMap(values, headers) {
  const dataMap = {};
  
  values.forEach(row => {
    const key = row.join('|'); // Crear una clave única basada en la concatenación de valores de fila
    const data = {};
    headers.forEach((header, index) => {
      data[header] = row[index];
    });
    dataMap[key] = data;
  });
  
  return dataMap;
}
