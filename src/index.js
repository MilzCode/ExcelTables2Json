import ExcelJS from 'exceljs';

/**
 * Convierte un archivo Excel a JSON.
 * 
 * @param {File} _file - Archivo Excel a procesar.
 * @param {string} _tableName - Nombre de la tabla dentro del Excel a extraer.
 * @param {boolean} [_isColumnsObjects=false] - Si es `true`, retorna datos estructurados en objetos con columnas como claves.
 * @returns {Promise<Object|null>} Objeto JSON con los datos extraídos o `null` si no se encuentra la tabla o hay errores.
 */
const ExcelTables2Json = async (_file, _tableName, _isColumnsObjects = false) => {
  if (!_file) {
    // Si no se pasa el archivo, retornar null.
    return null;
  }

  try {
    const workbook = new ExcelJS.Workbook(); // Crear una instancia del Workbook de ExcelJS.
    const reader = new FileReader(); // Crear un lector para el archivo.

    // Devolver una Promesa para manejar el procesamiento asincrónico.
    return new Promise((resolve, reject) => {
      reader.onload = async (e) => {
        try {
          const arrayBuffer = e.target.result;
          // Cargar el archivo Excel en el Workbook.
          await workbook.xlsx.load(arrayBuffer);

          // Obtener todas las hojas del libro.
          const sheets = workbook.worksheets;

          let returnData = null;

          // Recorrer todas las hojas para encontrar la tabla especificada.
          sheets.forEach((sheet) => {
            const tables = sheet.tables;
            const tableNames = Object.keys(tables);

            // Verificar si la tabla existe.
            if (tableNames.includes(_tableName)) {
              const table = tables[_tableName].table;
              const tableRange = table.tableRef; // Ejemplo: 'B2:C5'.
              const [start, end] = tableRange.split(':');
              const [startCol, startRow] = start.split(/(\d+)/);
              const [endCol, endRow] = end.split(/(\d+)/);

              returnData = {};

              if (_isColumnsObjects) {
                // Estructurar los datos con columnas como objetos.
                const columns = [];
                for (let j = startCol.charCodeAt(0); j <= endCol.charCodeAt(0); j++) {
                  const cell = sheet.getCell(String.fromCharCode(j) + startRow);
                  columns.push(cell.value);
                }

                const data = {};
                for (let i = parseInt(startRow) + 1; i <= parseInt(endRow); i++) {
                  for (let j = startCol.charCodeAt(0); j <= endCol.charCodeAt(0); j++) {
                    const cell = sheet.getCell(String.fromCharCode(j) + i);
                    if (!data[columns[j - startCol.charCodeAt(0)]]) {
                      data[columns[j - startCol.charCodeAt(0)]] = [];
                    }
                    data[columns[j - startCol.charCodeAt(0)]].push(cell.value);
                  }
                }
                returnData.columns = columns;
                returnData.data = data;
              } else {
                // Estructurar los datos como filas.
                const data = [];
                for (let i = parseInt(startRow); i <= parseInt(endRow); i++) {
                  const row = [];
                  for (let j = startCol.charCodeAt(0); j <= endCol.charCodeAt(0); j++) {
                    const cell = sheet.getCell(String.fromCharCode(j) + i);
                    row.push(cell.value);
                  }
                  data.push(row);
                }
                returnData.data = data;
              }
            }
          });

          resolve(returnData || null); // Resolver con los datos o `null` si no se encuentra la tabla.
        } catch (err) {
          reject(null); // Manejo de errores en la carga o procesamiento.
        }
      };

      reader.readAsArrayBuffer(_file); // Leer el archivo como ArrayBuffer.
    });
  } catch (err) {
    return null; // Si ocurre un error al intentar procesar el archivo.
  }
};

export default ExcelTables2Json;
