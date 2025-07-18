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
var __exportStar = (this && this.__exportStar) || function(m, exports) {
    for (var p in m) if (p !== "default" && !Object.prototype.hasOwnProperty.call(exports, p)) __createBinding(exports, m, p);
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.changelog = void 0;
__exportStar(require("./file-reader-functions/excel-file-functions"), exports);
__exportStar(require("./file-reader-functions/excel-headers"), exports);
__exportStar(require("./file-reader-functions/excel-raw-data"), exports);
__exportStar(require("./file-reader-functions/excel-row-count"), exports);
__exportStar(require("./file-reader-functions/file-meta-data"), exports);
__exportStar(require("./file-reader-functions/find-duplicate-headers"), exports);
__exportStar(require("./file-reader-functions/find-special-character-cells"), exports);
__exportStar(require("./file-reader-functions/is-column-data-available"), exports);
const changelog_json_1 = __importDefault(require("./changelog.json"));
exports.changelog = changelog_json_1.default;
