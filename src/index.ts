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
 * Interface for the result object returned by findSpecialCharacterCells.
 */
interface SpecialCharacterCellResult {
  columnName: string; // The header name of the column, or column letter if no header
  cellAddress: string; // The Excel cell address (e.g., "A1", "B5")
  cellValue: string; // The actual value of the cell that contains special characters
}

/**
 * Helper function to read a File object as an ArrayBuffer.
 * This is necessary because XLSX.read expects an ArrayBuffer for browser environments.
 * @param file The File object from a user input.
 * @returns A Promise that resolves with the ArrayBuffer of the file.
 */
export async function readFileAsArrayBuffer(file: File): Promise<ArrayBuffer> {
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
 * Processes the input data (File, ArrayBuffer, or Base64 string) and returns an XLSX.WorkBook object.
 * This internal helper centralizes the data parsing logic.
 * @param data The input data which can be a File, ArrayBuffer, or Base64 string.
 * @returns A Promise that resolves with the XLSX.WorkBook object.
 */
async function getWorkbookFromData(data: File | ArrayBuffer | string): Promise<XLSX.WorkBook> {
  if (data instanceof File) {
    const arrayBuffer = await readFileAsArrayBuffer(data);
    return XLSX.read(arrayBuffer, { type: 'array' });
  } else if (data instanceof ArrayBuffer) {
    return XLSX.read(data, { type: 'array' });
  } else if (typeof data === 'string') {
    // Assume string is Base64 encoded Excel data
    return XLSX.read(data, { type: 'base64' });
  } else {
    throw new Error("Unsupported data type provided. Expected File, ArrayBuffer, or Base64 string.");
  }
}

/**
 * Reads an Excel file (from a File, ArrayBuffer, or Base64 string) and returns the count of rows up to the maximum row
 * that contains at least one non-empty value in any column within the specified sheet.
 * This effectively gives the "last data row number" of the sheet.
 *
 * @param data The input data which can be a File object, an ArrayBuffer, or a Base64 encoded string. Can be null if no data is provided.
 * @param sheetName (Optional) The name of the sheet to read. If not provided, the first sheet will be used.
 * @returns A Promise that resolves with the count of rows up to the last data row, or rejects with an error.
 */
export async function getExcelRowCount(data: File | ArrayBuffer | string | null, sheetName?: string): Promise<number> {
  if (!data) {
    throw new Error("No data provided. Please provide a File, ArrayBuffer, or Base64 string.");
  }

  try {
    const workbook: XLSX.WorkBook = await getWorkbookFromData(data);

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
    throw new Error(`Failed to get row count from Excel data: ${error.message}`);
  }
}

/**
 * Reads an Excel file (from a File, ArrayBuffer, or Base64 string) and returns the headers (first row) of a specified sheet.
 *
 * @param data The input data which can be a File object, an ArrayBuffer, or a Base64 encoded string. Can be null if no data is provided.
 * @param sheetName (Optional) The name of the sheet to read. If not provided, the first sheet will be used.
 * @returns A Promise that resolves with an array of header strings, or rejects with an error.
 */
export async function getExcelHeaders(data: File | ArrayBuffer | string | null, sheetName?: string): Promise<string[]> {
  if (!data) {
    throw new Error("No data provided. Please provide a File, ArrayBuffer, or Base64 string.");
  }

  try {
    const workbook: XLSX.WorkBook = await getWorkbookFromData(data);
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
    throw new Error(`Failed to get headers from Excel data: ${error.message}`);
  }
}

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
export async function isExcelColumnPopulated(data: File | ArrayBuffer | string | null, headerName: string, sheetName?: string): Promise<boolean> {
  if (!data) {
    throw new Error("No data provided. Please provide a File, ArrayBuffer, or Base64 string.");
  }

  try {
    const workbook: XLSX.WorkBook = await getWorkbookFromData(data);

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
    throw new Error(`Failed to check column population in Excel data: ${error.message}`);
  }
}

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
export async function findSpecialCharacterCells(
  data: File | ArrayBuffer | string | null,
  sheetName?: string
): Promise<SpecialCharacterCellResult[]> {
  if (!data) {
    throw new Error("No data provided. Please provide a File, ArrayBuffer, or Base64 string.");
  }

  try {
    const workbook: XLSX.WorkBook = await getWorkbookFromData(data);
    const targetSheetName: string = sheetName || workbook.SheetNames[0];

    if (!workbook.SheetNames.includes(targetSheetName)) {
      throw new Error(`Sheet '${targetSheetName}' not found in the Excel file.`);
    }

    const worksheet: XLSX.WorkSheet = workbook.Sheets[targetSheetName];

    if (!worksheet || !worksheet['!ref']) {
      return []; // Empty sheet or no defined range, so no special characters
    }

    const results: SpecialCharacterCellResult[] = [];
    const range: XLSX.Range = XLSX.utils.decode_range(worksheet['!ref']);

    // Get headers to map column index to column name for the results
    const rawHeaders: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, range: 0, raw: false });
    const headers: string[] = (rawHeaders.length > 0 && Array.isArray(rawHeaders[0]))
      ? rawHeaders[0].filter((h: any) => typeof h === 'string' && h.trim() !== '')
      : [];

    // Regex to detect special characters (anything not a Unicode letter, Unicode number, or common whitespace)
    // The 'u' flag is for Unicode support.
    const specialCharRegex = /[^\p{L}\p{N}\s]/u;

    for (let R: number = range.s.r; R <= range.e.r; ++R) { // R is 0-based row index
      for (let C: number = range.s.c; C <= range.e.c; ++C) { // C is 0-based column index
        const cellAddress: string = XLSX.utils.encode_cell({ r: R, c: C });
        const cell: ExcelCell | undefined = worksheet[cellAddress];

        // Check if the cell exists and has a value that is not null/undefined
        if (cell && cell.v !== undefined && cell.v !== null) {
          const cellValue: string = String(cell.v); // Convert to string for regex test. Do not trim here.

          // Test if the cell value contains any special characters
          if (specialCharRegex.test(cellValue)) {
            // Determine column name: use header if available, otherwise use Excel column letter (e.g., 'A', 'B')
            const columnName: string = headers[C] || XLSX.utils.encode_col(C);
            results.push({
              columnName: columnName,
              cellAddress: cellAddress,
              cellValue: cellValue // Store the original value including whitespace/newlines/tabs
            });
          }
        }
      }
    }

    return results;

  } catch (error: any) {
    throw new Error(`Failed to find special character cells in Excel data: ${error.message}`);
  }
}

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
export async function findDuplicateHeaders(
  data: File | ArrayBuffer | string | null,
  sheetName?: string,
  caseInsensitive: boolean = true
): Promise<string[]> {
  if (!data) {
    throw new Error("No data provided. Please provide a File, ArrayBuffer, or Base64 string.");
  }

  try {
    const workbook: XLSX.WorkBook = await getWorkbookFromData(data);
    const targetSheetName: string = sheetName || workbook.SheetNames[0];

    if (!workbook.SheetNames.includes(targetSheetName)) {
      throw new Error(`Sheet '${targetSheetName}' not found in the Excel file.`);
    }

    const worksheet: XLSX.WorkSheet = workbook.Sheets[targetSheetName];

    if (!worksheet || !worksheet['!ref']) {
      return []; // Empty sheet or no defined range, so no headers to check
    }

    // Get headers from the first row
    const rawHeaders: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, range: 0, raw: false });
    let headers: string[] = [];

    if (rawHeaders.length > 0 && Array.isArray(rawHeaders[0])) {
      // Filter out non-string or empty headers and convert to string
      headers = rawHeaders[0].filter((h: any) => typeof h === 'string' && h.trim() !== '');
    } else {
      return []; // No headers found
    }

    const seenHeaders = new Set<string>();
    const duplicateHeaders: string[] = [];

    for (const header of headers) {
      const compareHeader = caseInsensitive ? header.toLowerCase() : header;
      if (seenHeaders.has(compareHeader)) {
        // Add the original header name to the list of duplicates
        if (!duplicateHeaders.includes(header)) { // Avoid adding the same original header multiple times
          duplicateHeaders.push(header);
        }
      } else {
        seenHeaders.add(compareHeader);
      }
    }

    return duplicateHeaders;

  } catch (error: any) {
    throw new Error(`Failed to find duplicate headers in Excel data: ${error.message}`);
  }
}
