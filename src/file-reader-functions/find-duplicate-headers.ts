import * as XLSX from 'xlsx';
import { getWorkbookFromData } from './excel-file-functions';

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

  } catch (error: unknown) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    throw new Error(`Failed to find duplicate headers in Excel data: ${errorMessage}`);
  }
}