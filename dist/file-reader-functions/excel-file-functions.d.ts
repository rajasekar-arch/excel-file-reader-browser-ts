import * as XLSX from 'xlsx';
/**
 * Helper function to read a File object as an ArrayBuffer.
 * This is necessary because XLSX.read expects an ArrayBuffer for browser environments.
 * @param file The File object from a user input.
 * @returns A Promise that resolves with the ArrayBuffer of the file.
 */
export declare function readFileAsArrayBuffer(file: File): Promise<ArrayBuffer>;
/**
 * Processes the input data (File, ArrayBuffer, or Base64 string) and returns an XLSX.WorkBook object.
 * This internal helper centralizes the data parsing logic.
 * @param data The input data which can be a File, ArrayBuffer, or Base64 string.
 * @returns A Promise that resolves with the XLSX.WorkBook object.
 */
export declare function getWorkbookFromData(data: File | ArrayBuffer | string): Promise<XLSX.WorkBook>;
/**
 * Reads an Excel file and returns the number of columns from the header row.
 *
 * @param data The input data (File, ArrayBuffer, or Base64 string)
 * @param sheetName Optional sheet name. Defaults to the first sheet.
 * @returns A Promise resolving to the number of columns.
 */
export declare function getExcelColumnCount(data: File | ArrayBuffer | string | null, sheetName?: string): Promise<number>;
