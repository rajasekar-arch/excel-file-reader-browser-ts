export interface ExcelMetadata {
    fileName: string;
    sheetNames: string[];
    sheetCount: number;
    fileType: string;
    fileSize: number;
    fileCreatedAt?: Date;
    fileModifiedAt?: Date;
    sheetDetails: {
        name: string;
        totalRows: number;
        totalColumns: number;
        hasData: boolean;
    }[];
}
/**
 * Extracts metadata from an Excel file (File, ArrayBuffer, or Base64 string).
 *
 * @param data - Excel file as File | ArrayBuffer | Base64
 * @param fileName - Optional original file name for reference
 * @returns Excel metadata including sheet names, row/column counts
 */
export declare function getExcelMetadata(data: File, fileName?: string): Promise<ExcelMetadata>;
