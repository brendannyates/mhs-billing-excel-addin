// Minimal ExcelScript surface used by office-scripts/*.ts (signatures follow the official API) — for local typechecking only.
declare namespace ExcelScript {
  type Value = string | number | boolean;
  interface Range {
    getValues(): Value[][];
    setValues(values: Value[][]): void;
    getCell(row: number, column: number): Range;
    setValue(value: Value): void;
  }
  interface Table { resize(range: Range): void; }
  interface Worksheet {
    getUsedRange(valuesOnly?: boolean): Range;
    getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number): Range;
    getTables(): Table[];
  }
  interface Workbook { getWorksheet(name: string): Worksheet; }
}
