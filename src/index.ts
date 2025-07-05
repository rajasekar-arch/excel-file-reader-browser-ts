// src/index.ts

import * as XLSX from 'xlsx';

/**
 * Helper function to read a File object as an ArrayBuffer.
 * This is necessary because XLSX.read expects an ArrayBuffer for browser environments.
 * @param file The File object from a user input.
 * @returns A Promise that resolves with the ArrayBuffer of the file.
 */
async function readFileAsArrayBuffer(file: File): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e: ProgressEvent<FileReader>) => {
      if (e.target?.result instanceof ArrayBuffer) {
        resolve(e.target.result);
      } else {
        reject(new Error("Failed to read file as ArrayBuffer."));
      }
    };
    reader.onerror = (err) => reject(new Error(`File reading error: ${reader.error?.message || 'Unknown error'}`));
    reader.readAsArrayBuffer(file);
  });
}

/**
 * Reads an Excel file (from a File object) and returns the total number of rows in a specified sheet.
 * This count includes header rows.
 *
 * @param file The File object representing the Excel file.
 * @param sheetName (Optional) The name of the sheet to read. If not provided, the first sheet will be used.
 * @returns A Promise that resolves with the total number of rows, or rejects with an error.
 */
export async function getExcelRowCount(file: File, sheetName?: string): Promise<number> {
  try {
    const arrayBuffer = await readFileAsArrayBuffer(file);
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });

    const targetSheetName = sheetName || workbook.SheetNames[0];

    if (!workbook.SheetNames.includes(targetSheetName)) {
      throw new Error(`Sheet '${targetSheetName}' not found in the Excel file.`);
    }

    const worksheet = workbook.Sheets[targetSheetName];

    if (!worksheet || !worksheet['!ref']) {
      return 0; // Empty sheet or no defined range
    }

    const range = XLSX.utils.decode_range(worksheet['!ref']);
    const rowCount = range.e.r + 1;

    return rowCount;

  } catch (error: any) {
    throw new Error(`Failed to get row count from Excel file: ${error.message}`);
  }
}

/**
 * Reads an Excel file (from a File object) and returns the headers (first row) of a specified sheet.
 *
 * @param file The File object representing the Excel file.
 * @param sheetName (Optional) The name of the sheet to read. If not provided, the first sheet will be used.
 * @returns A Promise that resolves with an array of header strings, or rejects with an error.
 */
export async function getExcelHeaders(file: File, sheetName?: string): Promise<string[]> {
  try {
    const arrayBuffer = await readFileAsArrayBuffer(file);
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });

    const targetSheetName = sheetName || workbook.SheetNames[0];

    if (!workbook.SheetNames.includes(targetSheetName)) {
      throw new Error(`Sheet '${targetSheetName}' not found in the Excel file.`);
    }

    const worksheet = workbook.Sheets[targetSheetName];

    if (!worksheet || !worksheet['!ref']) {
      return []; // Empty sheet or no defined range, thus no headers
    }

    const headers: string[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, range: 0, raw: false });

    if (headers.length > 0 && Array.isArray(headers[0])) {
      return headers[0].filter(header => typeof header === 'string' && header.trim() !== '');
    } else {
      return []; // No headers found
    }

  } catch (error: any) {
    throw new Error(`Failed to get headers from Excel file: ${error.message}`);
  }
}

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
export async function isExcelColumnPopulated(file: File, headerName: string, sheetName?: string): Promise<boolean> {
  try {
    const arrayBuffer = await readFileAsArrayBuffer(file);
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });

    const targetSheetName = sheetName || workbook.SheetNames[0];

    if (!workbook.SheetNames.includes(targetSheetName)) {
      throw new Error(`Sheet '${targetSheetName}' not found in the Excel file.`);
    }

    const worksheet = workbook.Sheets[targetSheetName];

    if (!worksheet || !worksheet['!ref']) {
      return false; // Empty sheet, no data, so column is not populated
    }

    const jsonData: any[] = XLSX.utils.sheet_to_json(worksheet, { raw: false });

    if (jsonData.length === 0) {
      return false; // No data rows, so column is not populated
    }

    const headers = Object.keys(jsonData[0]);
    if (!headers.includes(headerName)) {
        throw new Error(`Header '${headerName}' not found in the Excel sheet.`);
    }

    for (const row of jsonData) {
      const value = row[headerName];
      if (value !== null && value !== undefined && String(value).trim() !== '') {
        return true; // Found at least one populated cell in the column
      }
    }

    return false; // No populated cells found in the column
  } catch (error: any) {
    throw new Error(`Failed to check column population in Excel file: ${error.message}`);
  }
}