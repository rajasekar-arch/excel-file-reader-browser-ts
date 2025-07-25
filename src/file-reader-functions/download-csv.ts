import * as XLSX from "xlsx";
import { getWorkbookFromData } from "./excel-file-functions";

/**
 * Converts an Excel file to a CSV string.
 * @param data The File, ArrayBuffer, or Base64 string of the Excel file.
 * @param sheetName The name of the sheet to convert (default is the first sheet).
 * @returns A Promise that resolves with the CSV string.
 */
export async function convertExcelToCSV(
    data: File | ArrayBuffer | string,
    sheetName?: string
): Promise<string> {
    if (!data) {
        throw new Error("No data provided. Provide a File, ArrayBuffer, or Base64 string.");
    }

    const workbook: XLSX.WorkBook = await getWorkbookFromData(data);
    const targetSheetName: string = sheetName || workbook.SheetNames[0];

    if (!workbook.SheetNames.includes(targetSheetName)) {
        throw new Error(`Sheet '${targetSheetName}' not found in the Excel file.`);
    }

    const worksheet: XLSX.WorkSheet = workbook.Sheets[targetSheetName];
    if (!worksheet || !worksheet['!ref']) return '';

    // Convert the worksheet to CSV format
    return XLSX.utils.sheet_to_csv(worksheet);
}
/**
 * Downloads the Excel file as a CSV file.
 * @param data The File, ArrayBuffer, or Base64 string of the Excel file.
 * @param sheetName The name of the sheet to convert (default is the first sheet).
 * @param fileName The name of the file to download (default is 'download.csv').
 * @return A Promise that resolves when the file is downloaded.
 * */
export async function downloadExcelAsCSV(
    data: File | ArrayBuffer | string,
    sheetName?: string,
    fileName: string = "download.csv"
): Promise<void> {
    const csvContent = await convertExcelToCSV(data, sheetName);
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });

    // Create a link element to trigger the download
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);
    link.setAttribute("href", url);
    link.setAttribute("download", fileName);

    // Append to the body and trigger click
    document.body.appendChild(link);
    link.click();

    // Clean up
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}
