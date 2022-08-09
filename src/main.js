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
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const path = __importStar(require("path"));
const exceljs_1 = __importDefault(require("exceljs"));
const filePathDataExcel = path.resolve(__dirname + '../data/excel', 'iso_2digit_alpha_country_codes.xls');
const getCellValue = (row, cellIndex) => {
    const cell = row.getCell(cellIndex);
    return cell.value ? cell.value.toString() : '';
};
const main = async () => {
    const workbook = new exceljs_1.default.Workbook();
    const content = await workbook.xlsx.readFile(filePathDataExcel);
    const worksheet = content.worksheets[1];
    const rowStartIndex = 2;
    const numberOfRows = worksheet.rowCount - 3;
    const rows = worksheet.getRows(rowStartIndex, numberOfRows) ?? [];
    const countries = rows.map((row) => {
        return {
            // @ts-ignore
            isoCode: getCellValue(row, 1),
            // @ts-ignore
            countryName: getCellValue(row, 2),
        };
    });
    console.log(countries);
};
main().then();
