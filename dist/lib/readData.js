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
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
const XLSX = __importStar(require("xlsx"));
/**
 * Creates workbook data from file and returns rows per sheet
 *
 * @param {XLSX.WorkBook} wb
 * @return {*}  {WorkbookData}
 */
function destructData(wb) {
    const sheets = wb.SheetNames;
    const sheetsData = {};
    const sheetsMaxWidth = {};
    if (!sheets) {
        throw new Error("Unable to read sheets");
    }
    sheets.forEach((sheet) => {
        var _a, _b;
        sheetsData[sheet] = XLSX.utils.sheet_to_json(wb.Sheets[sheet], {
            header: 1,
            raw: false
        });
        sheetsMaxWidth[sheet] = ((_b = (_a = XLSX.utils.decode_range(wb.Sheets[sheet]['!ref'] || "")) === null || _a === void 0 ? void 0 : _a.e) === null || _b === void 0 ? void 0 : _b.c) || 50;
    });
    return { sheetsData, sheetsMaxWidth };
}
exports.default = (fileName) => __awaiter(void 0, void 0, void 0, function* () {
    const wb = yield XLSX.readFile(fileName, {});
    return destructData(wb);
});
