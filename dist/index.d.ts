/**
 * Reads an Excel file (from a File object) and returns the total number of rows in a specified sheet.
 * This count includes header rows.
 *
 * @param file The File object representing the Excel file.
 * @param sheetName (Optional) The name of the sheet to read. If not provided, the first sheet will be used.
 * @returns A Promise that resolves with the total number of rows, or rejects with an error.
 */
export declare function getExcelRowCount(file: File, sheetName?: string): Promise<number>;
/**
 * Reads an Excel file (from a File object) and returns the headers (first row) of a specified sheet.
 *
 * @param file The File object representing the Excel file.
 * @param sheetName (Optional) The name of the sheet to read. If not provided, the first sheet will be used.
 * @returns A Promise that resolves with an array of header strings, or rejects with an error.
 */
export declare function getExcelHeaders(file: File, sheetName?: string): Promise<string[]>;
/**
 * Reads an Excel file (from a File object) and checks if a specific column (identified by its header name)
 * contains any non-empty values in its data rows.
 *
 * @param file The File object representing the Excel file.
 * @param headerName The exact name of the header column to check.
 * @param sheetName (Optional) The name of the sheet to read. If not provided, the first sheet will be used.
 * @returns A Promise that resolves with `true` if the column has at least one non-empty value, `false` otherwise.
 * Rejects with an error if the file or sheet is not found, or if the header name does not exist.
 */
export declare function isExcelColumnPopulated(file: File, headerName: string, sheetName?: string): Promise<boolean>;
