function dataProfiling() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseSheet = ss.getSheetByName('Base');
  const profilingSheet = ss.getSheetByName('Profiling');
  
  if (!baseSheet || !profilingSheet) {
    SpreadsheetApp.getUi().alert('Las hojas "Base" o "Profiling" no existen.');
    return;
  }

  // Limpiar la hoja Profiling
  profilingSheet.clear();
  profilingSheet.getRange('A1:T1').setValues([[
    'Variable', 'Random_Example', 'Total_Records', 'Nulls', 'Blanks', 'Distinct',
    'Max_Value', 'Min_Value', 'Range', 'Negatives', 'Positives', 'Equal_Zero',
    'Average', 'Median', 'Variance', 'Standard_Deviation', 'Quartiles', 'Max_Length',
    'Min_Length', 'Mode'
  ]]);

  const data = baseSheet.getDataRange().getValues();
  const headers = data[0];
  const values = data.slice(1);
  const profile = [];

  headers.forEach((header, colIndex) => {
    const colValues = values.map(row => row[colIndex]);
    const colValuesNoNulls = colValues.filter(val => val !== '' && val !== null);
    const data = {
      'Variable': header,
      'Random_Example': colValuesNoNulls.length ? colValuesNoNulls[Math.floor(Math.random() * colValuesNoNulls.length)] : null,
      'Total_Records': colValues.length,
      'Nulls': colValues.filter(val => val === null).length,
      'Blanks': colValues.filter(val => val === '').length,
      'Distinct': [...new Set(colValues)].filter(val => val !== '' && val !== null).length,
      'Max_Length': Math.max(...colValues.map(val => val.toString().length)),
      'Min_Length': Math.min(...colValues.map(val => val.toString().length)),
      'Mode': colValuesNoNulls.length ? getMode(colValuesNoNulls) : null
    };

    if (colValuesNoNulls.length && !isNaN(colValuesNoNulls[0])) {
      const numValues = colValuesNoNulls.map(Number);
      data['Max_Value'] = Math.max(...numValues);
      data['Min_Value'] = Math.min(...numValues);
      data['Range'] = data['Max_Value'] - data['Min_Value'];
      data['Negatives'] = numValues.filter(val => val < 0).length;
      data['Positives'] = numValues.filter(val => val > 0).length;
      data['Equal_Zero'] = numValues.filter(val => val === 0).length;
      data['Average'] = average(numValues);
      data['Median'] = median(numValues);
      data['Variance'] = variance(numValues);
      data['Standard_Deviation'] = Math.sqrt(data['Variance']);
      data['Quartiles'] = JSON.stringify({
        '0.25': percentile(numValues, 25),
        '0.5': percentile(numValues, 50),
        '0.75': percentile(numValues, 75)
      });
    } else {
      data['Max_Value'] = data['Min_Value'] = data['Range'] = data['Negatives'] = data['Positives'] = data['Equal_Zero'] = null;
      data['Average'] = data['Median'] = data['Variance'] = data['Standard_Deviation'] = data['Quartiles'] = null;
    }

    profile.push(data);
  });

  const profileValues = profile.map(row => [
    row['Variable'], row['Random_Example'], row['Total_Records'], row['Nulls'], row['Blanks'], row['Distinct'],
    row['Max_Value'], row['Min_Value'], row['Range'], row['Negatives'], row['Positives'], row['Equal_Zero'],
    row['Average'], row['Median'], row['Variance'], row['Standard_Deviation'], row['Quartiles'],
    row['Max_Length'], row['Min_Length'], row['Mode']
  ]);

  profilingSheet.getRange(2, 1, profileValues.length, profileValues[0].length).setValues(profileValues);
}

function getMode(arr) {
  return arr.sort((a,b) =>
    arr.filter(v => v===a).length
    - arr.filter(v => v===b).length
  ).pop();
}

function average(arr) {
  return arr.reduce((a, b) => a + b, 0) / arr.length;
}

function median(arr) {
  const sorted = arr.slice().sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 !== 0 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
}

function variance(arr) {
  const avg = average(arr);
  return average(arr.map(num => Math.pow(num - avg, 2)));
}

function percentile(arr, p) {
  arr.sort((a, b) => a - b);
  const index = (p / 100) * (arr.length - 1);
  const lower = Math.floor(index);
  const upper = lower + 1;
  const weight = index % 1;
  if (upper >= arr.length) return arr[lower];
  return arr[lower] * (1 - weight) + arr[upper] * weight;
}
