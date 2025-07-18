// Define a type for a row object that allows string indexing
// This is used in isExcelColumnPopulated to correctly type dynamic access
export interface ExcelRow {
  [key: string]: any; // Allows any string to be used as an index, returning any type
}

// Define a minimal interface for XLSX.CellObject to ensure type safety when accessing cell.v
export interface ExcelCell {
  v?: any; // The raw value of the cell (value, not formatted text)
  w?: string; // The formatted text of the cell
  t?: string; // The type of the cell (e.g., 'n' for number, 's' for string, 'b' for boolean, 'd' for date)
  // Add other properties if needed, e.g., s for style, f for formula etc.
}

/**
 * Interface for the result object returned by findSpecialCharacterCells.
 */
export interface SpecialCharacterCellResult {
  columnName: string; // The header name of the column, or column letter if no header
  cellAddress: string; // The Excel cell address (e.g., "A1", "B5")
  cellValue: string; // The actual value of the cell that contains special characters
}
