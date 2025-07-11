import * as XLSX from 'xlsx';

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
export async function getExcelMetadata(
    data: File,
    fileName?: string
): Promise<ExcelMetadata> {
    if (!data) {
        throw new Error("No data provided.");
    }

    const workbook: XLSX.WorkBook =
        data instanceof File
            ? XLSX.read(await data.arrayBuffer(), { type: 'array' })
            : XLSX.read(data, { type: typeof data === 'string' ? 'base64' : 'array' });

    const sheetDetails = workbook.SheetNames.map((sheetName: string) => {
        const sheet:XLSX.WorkSheet = workbook.Sheets[sheetName];
        const ref: string | undefined = sheet['!ref'];
        let totalRows: number = 0;
        let totalColumns: number = 0;
        if (ref) {
            const range:XLSX.Range = XLSX.utils.decode_range(ref);
            totalRows = range.e.r - range.s.r + 1;
            totalColumns = range.e.c - range.s.c + 1;
        }

        return {
            name: sheetName,
            totalRows,
            totalColumns,
            hasData: !!ref,
        };
    });

    return {
        fileName: fileName ?? (data instanceof File ? data.name : ''),
        sheetNames: workbook.SheetNames,
        sheetCount: workbook.SheetNames.length,
        fileType: data instanceof File && data.type ? data.type : '',
        fileSize: data instanceof File && typeof data.size === 'number' ? data.size : 0,
        fileModifiedAt: data instanceof File && data.lastModified ? new Date(data.lastModified) : undefined,
        sheetDetails,
    };
}
