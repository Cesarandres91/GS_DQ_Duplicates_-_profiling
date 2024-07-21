# Data quality & profiling - Scripts para Googlesheets

Este repositorio contiene una serie de scripts en Google Apps Script diseñados para extender la funcionalidad de las hojas de cálculo de Google, permitiendo operaciones como la identificación de duplicados, el perfilado de datos y la comparación de hojas de cálculo.

## Funciones

### `findDuplicates()`

Identifica valores duplicados en una hoja llamada 'Base' y escribe los resultados en una hoja llamada 'Duplicados'.

**Pre-requisitos:**
- Una hoja de cálculo de Google con una hoja llamada 'Base' y otra llamada 'Duplicados'.

**Funcionamiento:**
- Busca duplicados en cada columna de la hoja 'Base'.
- Los duplicados encontrados se registran en la hoja 'Duplicados', incluyendo el campo, valor y conteo de veces que se repite el valor.

### `dataProfiling()`

Realiza un perfilado de datos de la hoja llamada 'Base' y escribe un resumen estadístico en una hoja llamada 'Profiling'.

**Pre-requisitos:**
- Una hoja de cálculo de Google con una hoja llamada 'Base' y otra llamada 'Profiling'.

**Funcionamiento:**
- Calcula estadísticas básicas como el máximo, mínimo, promedio, mediana, y más por cada columna.
- Los resultados se escriben en la hoja 'Profiling'.

### `compareSheets()`

Compara dos hojas dentro de la misma hoja de cálculo llamadas 'Base' y 'Base2' y registra las diferencias en una hoja nueva llamada 'Comparison'.

**Pre-requisitos:**
- Una hoja de cálculo de Google con las hojas llamadas 'Base' y 'Base2'.

**Funcionamiento:**
- Identifica cambios, eliminaciones y adiciones entre las dos hojas.
- Los resultados se escriben en la hoja 'Comparison'.

### `showPrompt()`

Muestra un cuadro de diálogo para que el usuario ingrese el ID de un archivo CSV, que luego se convertirá a una hoja de cálculo de Google Sheets.

### `convertCsvToGoogleSheet(fileId)`

Convierte un archivo CSV identificado por su `fileId` en una hoja de cálculo de Google Sheets.

### `replicateWithDataType()`

Crea o actualiza una hoja llamada 'Datatype' en la misma hoja de cálculo, replicando los datos de la hoja 'Base' e incluyendo el tipo de dato de cada celda.

## Instalación

1. Abre tu hoja de cálculo de Google.
2. Navega a `Extensiones > Apps Script`.
3. Copia y pega el código proporcionado en este repositorio en el editor de Apps Script.
4. Guarda y ejecuta las funciones como necesites.

## Contribuciones

Las contribuciones a este proyecto son bienvenidas. Por favor, asegúrate de seguir las mejores prácticas de desarrollo y mantén un código limpio y documentado.

## Licencia

Este proyecto está bajo la licencia MIT. Consulta el archivo `LICENSE` para más detalles.
