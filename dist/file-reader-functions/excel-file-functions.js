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
exports.readFileAsArrayBuffer = readFileAsArrayBuffer;
exports.getWorkbookFromData = getWorkbookFromData;
exports.getExcelColumnCount = getExcelColumnCount;
const XLSX = __importStar(require("xlsx"));
/**
 * Reads a File as ArrayBuffer in chunks (to handle huge files).
 * @param file The File object from input.
 * @param chunkSize The chunk size (default 4MB).
 * @returns A Promise that resolves with the ArrayBuffer of the file.
 */
async function readFileAsArrayBuffer(file, chunkSize = 4 * 1024 * 1024) {
    return new Promise((resolve, reject) => {
        const fileSize = file.size;
        let offset = 0;
        const chunks = [];
        const reader = new FileReader();
        reader.onload = (e) => {
            var _a;
            if (((_a = e.target) === null || _a === void 0 ? void 0 : _a.result) instanceof ArrayBuffer) {
                chunks.push(new Uint8Array(e.target.result));
                offset += chunkSize;
                if (offset < fileSize) {
                    readNextChunk();
                }
                else {
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
            }
            else {
                reject(new Error("Failed to read file chunk as ArrayBuffer."));
            }
        };
        reader.onerror = () => { var _a; return reject(new Error(`File reading error: ${((_a = reader.error) === null || _a === void 0 ? void 0 : _a.message) || 'Unknown error'}`)); };
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
async function getWorkbookFromData(data) {
    if (data instanceof File) {
        const arrayBuffer = await readFileAsArrayBuffer(data);
        return XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
    }
    else if (data instanceof ArrayBuffer) {
        return XLSX.read(data, { type: 'array', cellDates: true });
    }
    else if (typeof data === 'string') {
        return XLSX.read(data, { type: 'base64', cellDates: true });
    }
    else {
        throw new Error("Unsupported data type. Provide File, ArrayBuffer, or Base64 string.");
    }
}
/**
 * Efficiently retrieves the column count from the first row of an Excel sheet.
 * It processes only the first row instead of loading all rows.
 */
async function getExcelColumnCount(data, sheetName) {
    if (!data) {
        throw new Error("No data provided. Provide a File, ArrayBuffer, or Base64 string.");
    }
    const workbook = await getWorkbookFromData(data);
    const targetSheetName = sheetName || workbook.SheetNames[0];
    if (!workbook.SheetNames.includes(targetSheetName)) {
        throw new Error(`Sheet '${targetSheetName}' not found in the Excel file.`);
    }
    const worksheet = workbook.Sheets[targetSheetName];
    if (!worksheet || !worksheet['!ref'])
        return 0;
    // Extract only the first row for performance
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, range: 0 });
    const headerRow = rows[0] || [];
    return headerRow.length;
}
