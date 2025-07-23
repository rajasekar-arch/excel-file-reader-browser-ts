import * as XLSX from 'xlsx';
/**
 * Reads a File as ArrayBuffer in chunks (to handle huge files).
 * @param file The File object from input.
 * @param chunkSize The chunk size (default 4MB).
 * @returns A Promise that resolves with the ArrayBuffer of the file.
 */
export declare function readFileAsArrayBuffer(file: File, chunkSize?: number): Promise<ArrayBuffer>;
/**
 * Gets an XLSX.WorkBook object from File, ArrayBuffer, or Base64 string.
 */
export declare function getWorkbookFromData(data: File | ArrayBuffer | string): Promise<XLSX.WorkBook>;
/**
 * Efficiently retrieves the column count from the first row of an Excel sheet.
 * It processes only the first row instead of loading all rows.
 */
export declare function getExcelColumnCount(data: File | ArrayBuffer | string | null, sheetName?: string): Promise<number>;
