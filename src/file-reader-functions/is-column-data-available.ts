import * as XLSX from 'xlsx';
import { getWorkbookFromData } from "./excel-file-functions";
import { ExcelRow } from '../interfaces/excel.interface';


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

