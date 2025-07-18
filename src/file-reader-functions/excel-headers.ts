import * as XLSX from 'xlsx';
import { getWorkbookFromData } from "./excel-file-functions";


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
