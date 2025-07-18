import * as XLSX from 'xlsx';

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
export async function getWorkbookFromData(data: File | ArrayBuffer | string): Promise<XLSX.WorkBook> {
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
 * Reads an Excel file and returns the number of columns from the header row.
 *
 * @param data The input data (File, ArrayBuffer, or Base64 string)
 * @param sheetName Optional sheet name. Defaults to the first sheet.
 * @returns A Promise resolving to the number of columns.
 */
export async function getExcelColumnCount(
  data: File | ArrayBuffer | string | null,
  sheetName?: string
): Promise<number> {
  if (!data) {
    throw new Error("No data provided. Please provide a File, ArrayBuffer, or Base64 string.");
  }

  const workbook: XLSX.WorkBook = await getWorkbookFromData(data);
  const targetSheetName: string = sheetName || workbook.SheetNames[0];

  if (!workbook.SheetNames.includes(targetSheetName)) {
    throw new Error(`Sheet '${targetSheetName}' not found in the Excel file.`);
  }

  const worksheet: XLSX.WorkSheet = workbook.Sheets[targetSheetName];
  if (!worksheet || !worksheet['!ref']) return 0;

  const sheetData: (string | number | boolean | null | undefined)[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false }) as (string | null | undefined)[][];
  return sheetData[0]?.length || 0;
}
