/**
 * Interface for the result object returned by findSpecialCharacterCells.
 */
interface SpecialCharacterCellResult {
    columnName: string;
    cellAddress: string;
    cellValue: string;
}
/**
 * Helper function to read a File object as an ArrayBuffer.
 * This is necessary because XLSX.read expects an ArrayBuffer for browser environments.
 * @param file The File object from a user input.
 * @returns A Promise that resolves with the ArrayBuffer of the file.
 */
export declare function readFileAsArrayBuffer(file: File): Promise<ArrayBuffer>;
/**
 * Reads an Excel file (from a File, ArrayBuffer, or Base64 string) and returns the count of rows up to the maximum row
 * that contains at least one non-empty value in any column within the specified sheet.
 * This effectively gives the "last data row number" of the sheet.
 *
 * @param data The input data which can be a File object, an ArrayBuffer, or a Base64 encoded string. Can be null if no data is provided.
 * @param sheetName (Optional) The name of the sheet to read. If not provided, the first sheet will be used.
 * @returns A Promise that resolves with the count of rows up to the last data row, or rejects with an error.
 */
export declare function getExcelRowCount(data: File | ArrayBuffer | string | null, sheetName?: string): Promise<number>;
/**
 * Reads an Excel file (from a File, ArrayBuffer, or Base64 string) and returns the headers (first row) of a specified sheet.
 *
 * @param data The input data which can be a File object, an ArrayBuffer, or a Base64 encoded string. Can be null if no data is provided.
 * @param sheetName (Optional) The name of the sheet to read. If not provided, the first sheet will be used.
 * @returns A Promise that resolves with an array of header strings, or rejects with an error.
 */
export declare function getExcelHeaders(data: File | ArrayBuffer | string | null, sheetName?: string): Promise<string[]>;
/**
 * Reads an Excel file (from a File, ArrayBuffer, or Base64 string) and checks if a specific column (identified by its header name)
 * contains any non-empty values in its data rows.
 *
 * @param data The input data which can be a File object, an ArrayBuffer, or a Base64 encoded string. Can be null if no data is provided.
 * @param headerName The exact name of the header column to check.
 * @param sheetName (Optional) The name of the sheet to read. If not provided, the first sheet will be used.
 * @returns A Promise that resolves with `true` if the column has at least one non-empty value, `false` otherwise.
 * Rejects with an error if the data or sheet is not found, or if the header name does not exist.
 */
export declare function isExcelColumnPopulated(data: File | ArrayBuffer | string | null, headerName: string, sheetName?: string): Promise<boolean>;
/**
 * Reads an Excel file and finds cells that contain special characters.
 * Special characters are defined as anything that is not an alphanumeric character or common whitespace.
 *
 * @param data The input data (File, ArrayBuffer, or Base64 string). Can be null if no data is provided.
 * @param sheetName (Optional) The name of the sheet to read. If not provided, the first sheet will be used.
 * @returns A Promise that resolves with an array of SpecialCharacterCellResult objects,
 * each containing the column name, cell address, and the cell value.
 * Rejects with an error if the data or sheet is not found.
 */
export declare function findSpecialCharacterCells(data: File | ArrayBuffer | string | null, sheetName?: string): Promise<SpecialCharacterCellResult[]>;
/**
 * Reads an Excel file and finds any duplicate headers in a specified sheet.
 * It can check for duplicates in a case-sensitive or case-insensitive manner.
 *
 * @param data The input data (File, ArrayBuffer, or Base64 string). Can be null if no data is provided.
 * @param sheetName (Optional) The name of the sheet to read. If not provided, the first sheet will be used.
 * @param caseInsensitive (Optional) If true, performs a case-insensitive check for duplicates. Defaults to true.
 * @returns A Promise that resolves with an array of strings, where each string is a header name that appears more than once.
 * Rejects with an error if the data or sheet is not found.
 */
export declare function findDuplicateHeaders(data: File | ArrayBuffer | string | null, sheetName?: string, caseInsensitive?: boolean): Promise<string[]>;
export {};
