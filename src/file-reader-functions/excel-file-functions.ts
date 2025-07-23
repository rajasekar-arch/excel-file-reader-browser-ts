import * as XLSX from 'xlsx';

/**
 * Reads a File as ArrayBuffer in chunks (to handle huge files).
 * @param file The File object from input.
 * @param chunkSize The chunk size (default 4MB).
 * @returns A Promise that resolves with the ArrayBuffer of the file.
 */
export async function readFileAsArrayBuffer(file: File, chunkSize = 0.1 * 1024 * 1024): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const fileSize = file.size;
    let offset = 0;
    const chunks: Uint8Array[] = [];
    const reader = new FileReader();

    reader.onload = (e: ProgressEvent<FileReader>) => {
      if (e.target?.result instanceof ArrayBuffer) {
        chunks.push(new Uint8Array(e.target.result));
        offset += chunkSize;
        if (offset < fileSize) {
          readNextChunk();
        } else {
          // Combine chunks into a single ArrayBuffer
          const totalLength = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
          const combined = new Uint8Array(totalLength);
          let position = 0;
          for (const chunk of chunks) {
            combined.set(chunk, position);
            position += chunk.length;
          }
          resolve(combined.buffer);
        }
      } else {
        reject(new Error("Failed to read file chunk as ArrayBuffer."));
      }
    };

    reader.onerror = () => reject(new Error(`File reading error: ${reader.error?.message || 'Unknown error'}`));

    function readNextChunk() {
      const slice = file.slice(offset, offset + chunkSize);
      reader.readAsArrayBuffer(slice);
    }

    readNextChunk();
  });
}

/**
 * Gets an XLSX.WorkBook object from File, ArrayBuffer, or Base64 string.
 */
export async function getWorkbookFromData(data: File | ArrayBuffer | string): Promise<XLSX.WorkBook> {
  if (data instanceof File) {
    const arrayBuffer = await readFileAsArrayBuffer(data);
    return XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
  } else if (data instanceof ArrayBuffer) {
    return XLSX.read(data, { type: 'array', cellDates: true });
  } else if (typeof data === 'string') {
    return XLSX.read(data, { type: 'base64', cellDates: true });
  } else {
    throw new Error("Unsupported data type. Provide File, ArrayBuffer, or Base64 string.");
  }
}

/**
 * Efficiently retrieves the column count from the first row of an Excel sheet.
 * It processes only the first row instead of loading all rows.
 */
export async function getExcelColumnCount(
  data: File | ArrayBuffer | string | null,
  sheetName?: string
): Promise<number> {
  if (!data) {
    throw new Error("No data provided. Provide a File, ArrayBuffer, or Base64 string.");
  }

  const workbook: XLSX.WorkBook = await getWorkbookFromData(data);
  const targetSheetName: string = sheetName || workbook.SheetNames[0];

  if (!workbook.SheetNames.includes(targetSheetName)) {
    throw new Error(`Sheet '${targetSheetName}' not found in the Excel file.`);
  }

  const worksheet: XLSX.WorkSheet = workbook.Sheets[targetSheetName];
  if (!worksheet || !worksheet['!ref']) return 0;

  // Extract only the first row for performance
  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, range: 0 }) as any[][];
  const headerRow: any[] = rows[0] || [];
  return headerRow.length;
}
