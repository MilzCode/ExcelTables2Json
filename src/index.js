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
    return null;
  }

  try {
    const workbook = new ExcelJS.Workbook();
    const reader = new FileReader();

    return new Promise((resolve, reject) => {
      reader.onload = async (e) => {
        try {
          const arrayBuffer = e.target.result;
          await workbook.xlsx.load(arrayBuffer);
          const sheets = workbook.worksheets;
          let returnData = null;

          sheets.forEach((sheet) => {
            const tables = sheet.tables;
            const tableNames = Object.keys(tables);

            if (tableNames.includes(_tableName)) {
              const table = tables[_tableName].table;
              const tableRange = table.tableRef;
              const [start, end] = tableRange.split(':');
              const [startCol, startRow] = start.split(/(\d+)/);
              const [endCol, endRow] = end.split(/(\d+)/);
              returnData = {};

              if (_isColumnsObjects) {
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

          resolve(returnData || null);
        } catch (err) {
          reject(null);
        }
      };

      reader.readAsArrayBuffer(_file);
    });
  } catch (err) {
    return null;
  }
};

export default ExcelTables2Json;