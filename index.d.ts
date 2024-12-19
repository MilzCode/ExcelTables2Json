declare module "exceltables2json" {
  /**
   * Convierte una tabla de un archivo Excel a formato JSON.
   * @param file - Archivo Excel a procesar.
   * @param tableName - Nombre de la tabla dentro del archivo Excel.
   * @param isColumnsObjects - Si `true`, los datos se retornan como un objeto donde las llaves son los nombres de las columnas. Si `false`, como un array de arreglos.
   * @returns Una promesa que resuelve en un objeto JSON o `null` si ocurre un error.
   */
  function ExcelTables2Json(
    file: File,
    tableName: string,
    isColumnsObjects?: boolean
  ): Promise<ExcelTableResult | null>;

  interface ExcelTableResult {
    /**
     * Nombres de las columnas (si `isColumnsObjects` es `true`).
     */
    columns?: string[];

    /**
     * Datos de la tabla.
     * - Como un array de arrays si `isColumnsObjects` es `false`.
     * - Como un objeto con columnas como llaves si `isColumnsObjects` es `true`.
     */
    data: Array<Array<any>> | Record<string, any[]>;
  }

  export default ExcelTables2Json;
}