# ExcelTables2Json

**ExcelTables2Json** es una librería que convierte tablas dentro de archivos Excel (.xlsx) en objetos JSON, permitiendo al usuario elegir entre dos formatos de salida: un arreglo de arreglos o un objeto donde las llaves son los nombres de las columnas.

## 🚀 Características

- Convierte una tabla específica en un archivo Excel a JSON.
- Soporta dos formatos de salida:
  1. **Array de arreglos:** Las filas son representadas como arreglos.
  2. **Objetos con llaves:** Los nombres de las columnas son las llaves y los datos de las filas son los valores.
- Fácil de usar con JavaScript o TypeScript.

## 📦 Instalación

Usa npm para instalar la librería:

```bash
npm install exceltables2json
```

## 📖 Uso

```javascript
import ExcelTables2Json from 'exceltables2json';


//file example e.target.files[0]
const processExcel = async (file) => {
  
  const tableName = 'MyTable'; // Nombre de la tabla dentro del archivo Excel
  const isColumnsObjects = true; // Cambiar a `false` para obtener un array de arreglos

  const result = await ExcelTables2Json(file, tableName, isColumnsObjects);

  console.log(result);
};

```
##Ejemplo de salida tipo 1
```json
{
  "data": [
    ["Header1", "Header2"],
    ["Row1Col1", "Row1Col2"],
    ["Row2Col1", "Row2Col2"]
  ]
}
```

##Ejemplo de salida tipo 2
```json
{
  "columns": ["Header1", "Header2"],
  "data": {
    "Header1": ["Row1Col1", "Row2Col1"],
    "Header2": ["Row1Col2", "Row2Col2"]
  }
}
```

