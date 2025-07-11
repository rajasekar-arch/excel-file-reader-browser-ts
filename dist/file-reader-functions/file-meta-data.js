"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
exports.getExcelMetadata = getExcelMetadata;
const XLSX = __importStar(require("xlsx"));
/**
 * Extracts metadata from an Excel file (File, ArrayBuffer, or Base64 string).
 *
 * @param data - Excel file as File | ArrayBuffer | Base64
 * @param fileName - Optional original file name for reference
 * @returns Excel metadata including sheet names, row/column counts
 */
async function getExcelMetadata(data, fileName) {
    if (!data) {
        throw new Error("No data provided.");
    }
    const workbook = data instanceof File
        ? XLSX.read(await data.arrayBuffer(), { type: 'array' })
        : XLSX.read(data, { type: typeof data === 'string' ? 'base64' : 'array' });
    const sheetDetails = workbook.SheetNames.map((sheetName) => {
        const sheet = workbook.Sheets[sheetName];
        const ref = sheet['!ref'];
        let totalRows = 0;
        let totalColumns = 0;
        if (ref) {
            const range = XLSX.utils.decode_range(ref);
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
        fileName: fileName !== null && fileName !== void 0 ? fileName : (data instanceof File ? data.name : ''),
        sheetNames: workbook.SheetNames,
        sheetCount: workbook.SheetNames.length,
        fileType: data instanceof File && data.type ? data.type : '',
        fileSize: data instanceof File && typeof data.size === 'number' ? data.size : 0,
        fileModifiedAt: data instanceof File && data.lastModified ? new Date(data.lastModified) : undefined,
        sheetDetails,
    };
}
