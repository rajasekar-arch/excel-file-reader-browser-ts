import * as XLSX from 'xlsx';
import { getWorkbookFromData } from "./excel-file-functions";
import { ExcelCell } from '../interfaces/excel.interface';


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
