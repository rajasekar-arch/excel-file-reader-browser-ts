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
 * Helper function to read a File object as an ArrayBuffer.
 * This is necessary because XLSX.read expects an ArrayBuffer for browser environments.
 * @param file The File object from a user input.
 * @returns A Promise that resolves with the ArrayBuffer of the file.
 */
async function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            var _a;
            if (((_a = e.target) === null || _a === void 0 ? void 0 : _a.result) instanceof ArrayBuffer) {
                resolve(e.target.result);
            }
            else {
                reject(new Error("Failed to read file as ArrayBuffer."));
            }
        };
        reader.onerror = (err) => { var _a; return reject(new Error(`File reading error: ${((_a = reader.error) === null || _a === void 0 ? void 0 : _a.message) || 'Unknown error'}`)); };
        reader.readAsArrayBuffer(file);
    });
}
/**
 * Processes the input data (File, ArrayBuffer, or Base64 string) and returns an XLSX.WorkBook object.
 * This internal helper centralizes the data parsing logic.
 * @param data The input data which can be a File, ArrayBuffer, or Base64 string.
 * @returns A Promise that resolves with the XLSX.WorkBook object.
 */
async function getWorkbookFromData(data) {
    if (data instanceof File) {
        const arrayBuffer = await readFileAsArrayBuffer(data);
        return XLSX.read(arrayBuffer, { type: 'array' });
    }
    else if (data instanceof ArrayBuffer) {
        return XLSX.read(data, { type: 'array' });
    }
    else if (typeof data === 'string') {
        // Assume string is Base64 encoded Excel data
        return XLSX.read(data, { type: 'base64' });
    }
    else {
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
async function getExcelColumnCount(data, sheetName) {
    var _a;
    if (!data) {
        throw new Error("No data provided. Please provide a File, ArrayBuffer, or Base64 string.");
    }
    const workbook = await getWorkbookFromData(data);
    const targetSheetName = sheetName || workbook.SheetNames[0];
    if (!workbook.SheetNames.includes(targetSheetName)) {
        throw new Error(`Sheet '${targetSheetName}' not found in the Excel file.`);
    }
    const worksheet = workbook.Sheets[targetSheetName];
    if (!worksheet || !worksheet['!ref'])
        return 0;
    const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });
    return ((_a = sheetData[0]) === null || _a === void 0 ? void 0 : _a.length) || 0;
}
