// src/index.ts

import * as XLSX from 'xlsx';

// Define a type for a row object that allows string indexing
// This is used in isExcelColumnPopulated to correctly type dynamic access
interface ExcelRow {
  [key: string]: any; // Allows any string to be used as an index, returning any type
}

// Define a minimal interface for XLSX.CellObject to ensure type safety when accessing cell.v
interface ExcelCell {
  v?: any; // The raw value of the cell (value, not formatted text)
  w?: string; // The formatted text of the cell
  t?: string; // The type of the cell (e.g., 'n' for number, 's' for string, 'b' for boolean, 'd' for date)
  // Add other properties if needed, e.g., s for style, f for formula etc.
}

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
 * Reads an Excel file (from a File object) and returns the count of rows up to the maximum row
 * that contains at least one non-empty value in any column within the specified sheet.
 * This effectively gives the "last data row number" of the sheet.
 *
 * @param file The File object representing the Excel file. Can be null if no file is selected.
 * @param sheetName (Optional) The name of the sheet to read. If not provided, the first sheet will be used.
 * @returns A Promise that resolves with the count of rows up to the last data row, or rejects with an error.
 */
export async function getExcelRowCount(file: File | null, sheetName?: string): Promise<number> {
  if (!file) {
    throw new Error("No file provided. Please select an Excel file.");
  }

  try {
    const arrayBuffer: ArrayBuffer = await readFileAsArrayBuffer(file);
    const workbook: XLSX.WorkBook = XLSX.read(arrayBuffer, { type: 'array' });

    const targetSheetName: string = sheetName || workbook.SheetNames[0];

    if (!workbook.SheetNames.includes(targetSheetName)) {
      throw new Error(`Sheet '${targetSheetName}' not found in the Excel file.`);
    }

    const worksheet: XLSX.WorkSheet = workbook.Sheets[targetSheetName];

    // If the worksheet has no defined range (!ref), it's considered empty.
    if (!worksheet || !worksheet['!ref']) {
      return 0;
    }

    let maxRowWithData: number = 0;
    const range: XLSX.Range = XLSX.utils.decode_range(worksheet['!ref']);

    // Iterate through all cells within the detected range of the worksheet
    for (let R: number = range.s.r; R <= range.e.r; ++R) { // R is 0-based row index
      for (let C: number = range.s.c; C <= range.e.c; ++C) { // C is 0-based column index
        const cellAddress: string = XLSX.utils.encode_cell({ r: R, c: C });
        const cell: ExcelCell | undefined = worksheet[cellAddress]; // Use ExcelCell interface

        // Check if the cell exists and has a non-empty value
        if (cell && cell.v !== undefined && cell.v !== null) {
          const cellValue: string = String(cell.v).trim();
          if (cellValue !== '') {
            // Update maxRowWithData if this cell's row is higher than previously found
            // R is 0-based, so R + 1 gives the actual row number
            maxRowWithData = Math.max(maxRowWithData, R + 1);
          }
        }
      }
    }

    return maxRowWithData;

  } catch (error: any) {
    throw new Error(`Failed to get row count from Excel file: ${error.message}`);
  }
}

/**
 * Reads an Excel file (from a File object) and returns the headers (first row) of a specified sheet.
 *
 * @param file The File object representing the Excel file. Can be null if no file is selected.
 * @param sheetName (Optional) The name of the sheet to read. If not provided, the first sheet will be used.
 * @returns A Promise that resolves with an array of header strings, or rejects with an error.
 */
export async function getExcelHeaders(file: File | null, sheetName?: string): Promise<string[]> {
  if (!file) {
    throw new Error("No file provided. Please select an Excel file.");
  }

  try {
    const arrayBuffer: ArrayBuffer = await readFileAsArrayBuffer(file);
    const workbook: XLSX.WorkBook = XLSX.read(arrayBuffer, { type: 'array' });

    const targetSheetName: string = sheetName || workbook.SheetNames[0];

    if (!workbook.SheetNames.includes(targetSheetName)) {
      throw new Error(`Sheet '${targetSheetName}' not found in the Excel file.`);
    }

    const worksheet: XLSX.WorkSheet = workbook.Sheets[targetSheetName];

    if (!worksheet || !worksheet['!ref']) {
      return []; // Empty sheet or no defined range, thus no headers
    }

    const headers: string[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, range: 0, raw: false });

    if (headers.length > 0 && Array.isArray(headers[0])) {
      return headers[0].filter((header: any) => typeof header === 'string' && header.trim() !== '');
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
 * @param file The File object representing the Excel file. Can be null if no file is selected.
 * @param headerName The exact name of the header column to check.
 * @param sheetName (Optional) The name of the sheet to read. If not provided, the first sheet will be used.
 * @returns A Promise that resolves with `true` if the column has at least one non-empty value, `false` otherwise.
 * Rejects with an error if the file or sheet is not found, or if the header name does not exist.
 */
export async function isExcelColumnPopulated(file: File | null, headerName: string, sheetName?: string): Promise<boolean> {
  if (!file) {
    throw new Error("No file provided. Please select an Excel file.");
  }

  try {
    const arrayBuffer: ArrayBuffer = await readFileAsArrayBuffer(file);
    const workbook: XLSX.WorkBook = XLSX.read(arrayBuffer, { type: 'array' });

    const targetSheetName: string = sheetName || workbook.SheetNames[0];

    if (!workbook.SheetNames.includes(targetSheetName)) {
      throw new Error(`Sheet '${targetSheetName}' not found in the Excel file.`);
    }

    const worksheet: XLSX.WorkSheet = workbook.Sheets[targetSheetName];

    if (!worksheet || !worksheet['!ref']) {
      return false; // Empty sheet, no data, so column is not populated
    }

    // Explicitly type jsonData as an array of objects with string index signatures
    const jsonData: ExcelRow[] = XLSX.utils.sheet_to_json(worksheet, { raw: false });

    if (jsonData.length === 0) {
      return false; // No data rows, so column is not populated
    }

    // Check if the header exists in the first data row (which represents the header row in this context)
    const headers: string[] = Object.keys(jsonData[0]);
    if (!headers.includes(headerName)) {
        throw new Error(`Header '${headerName}' not found in the Excel sheet.`);
    }

    // Iterate through data rows (skipping the header row, as sheet_to_json already handles it)
    for (const row of jsonData) {
      // Now 'row' is typed as ExcelRow, allowing string indexing
      const value: any = row[headerName];
      // Check if the value is not null, undefined, and not an empty string after trimming
      if (value !== null && value !== undefined && String(value).trim() !== '') {
        return true; // Found at least one populated cell in the column
      }
    }

    return false; // No populated cells found in the column
  } catch (error: any) {
    throw new Error(`Failed to check column population in Excel file: ${error.message}`);
  }
}
