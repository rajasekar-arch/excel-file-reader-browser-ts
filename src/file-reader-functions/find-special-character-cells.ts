import * as XLSX from 'xlsx';

import { ExcelCell, SpecialCharacterCellResult } from "../interfaces/excel.interface";
import { getWorkbookFromData } from './excel-file-functions';


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

