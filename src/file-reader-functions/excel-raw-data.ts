import * as XLSX from 'xlsx';
import { getWorkbookFromData } from "./excel-file-functions";


/**
 * Reads an Excel file and returns the raw data as a 2D array (rows × columns).
 *
 * @param data The input data (File, ArrayBuffer, or Base64 string)
 * @param sheetName Optional sheet name. Defaults to the first sheet.
 * @returns A Promise resolving to a 2D array of strings.
 */
export async function getExcelRawData(
  data: File | ArrayBuffer | string | null,
  sheetName?: string
): Promise<string[][]> {
  if (!data) {
    throw new Error("No data provided. Please provide a File, ArrayBuffer, or Base64 string.");
  }

  const workbook: XLSX.WorkBook = await getWorkbookFromData(data);
  const targetSheetName: string = sheetName || workbook.SheetNames[0];

  if (!workbook.SheetNames.includes(targetSheetName)) {
    throw new Error(`Sheet '${targetSheetName}' not found in the Excel file.`);
  }

  const worksheet: XLSX.WorkSheet = workbook.Sheets[targetSheetName];
  if (!worksheet || !worksheet['!ref']) return [];

  const sheetData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });

  // Convert all cell values to strings and normalize undefined/null to ''
  const formattedData: string[][] = sheetData.map(row =>
    row.map(cell => (cell !== null && cell !== undefined) ? String(cell?.trim()) : '')
  );

  return formattedData;
}

