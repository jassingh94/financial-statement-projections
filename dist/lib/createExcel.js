"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.default = default_1;
const exceljs_1 = __importDefault(require("exceljs"));
const font = { name: 'Arial', size: 11 };
/**
 *
 * Created data for sub headers (periods) for excel
 * @param {((number | string)[][])} rows
 * @param {number} indexStartx
 * @param {number} indexStartY
 * @param {excelDataObject} data
 * @param {string} type
 * @param {(number)[]} historicalStartingCol
 * @param {(number)[]} projectionsStartingCol
 * @return {*}
 */
let createTypePeriodsSubHeaders = (rows, indexStartx, indexStartY, data, type, historicalStartingCol, projectionsStartingCol) => {
    if (!rows[indexStartx])
        rows[indexStartx] = new Array(Number(data.sheet.max.width.value)).fill("‎");
    let timeStamps = Object.keys(data[type]);
    let interval = 0;
    if (type === "projection")
        interval = projectionsStartingCol[1] - indexStartx - 1;
    for (let y = indexStartY + interval, xx = 0; y < indexStartY + interval + timeStamps.length; y++, xx++) {
        rows[indexStartx][y] = `${new Date(Number(timeStamps[xx])).toLocaleString('default', { month: 'long' })} ${new Date(Number(timeStamps[xx])).getFullYear()}`;
    }
    return rows;
};
/**
 *
 * Creates data for excel for month periods
 * @param {((number | string)[][])} rows
 * @param {number} indexStartx
 * @param {number} indexStartY
 * @param {excelDataObject} data
 * @param {string} type
 * @param {(number)[]} historicalStartingCol
 * @param {(number)[]} projectionsStartingCol
 * @param {{ [key: string]: any }[]} leftHeaders
 * @return {*}
 */
let createTypePeriodsData = (rows, indexStartx, indexStartY, data, type, historicalStartingCol, projectionsStartingCol, leftHeaders) => {
    if (!rows[indexStartx])
        rows[indexStartx] = new Array(Number(data.sheet.max.width.value)).fill("‎");
    let timeStamps = Object.entries(data[type]);
    let interval = 0;
    if (type === "projection")
        interval = projectionsStartingCol[1] - indexStartx;
    for (let y = indexStartY + interval, xx = 0; y < indexStartY + interval + timeStamps.length; y++, xx++) {
        for (let x = 0; x < leftHeaders.length; x++) {
            if (!rows[indexStartx + x])
                rows[indexStartx + x] = new Array(Number(data.sheet.max.width.value)).fill("‎");
            rows[indexStartx + x][0] = leftHeaders[x].name;
            let valueOfField = timeStamps[xx][1][leftHeaders[x].name].value;
            if (leftHeaders[x].modifiedBy
                && leftHeaders[x].modifiedBy instanceof Array
                && leftHeaders[x].modifiedBy.length
                && leftHeaders[x].modificationType) {
                valueOfField = leftHeaders[x].modifiedBy.reduce((p, c) => {
                    let backIndex = leftHeaders[x].thisIndex - c;
                    let val = Number(rows[indexStartx + x - backIndex][y]);
                    if (p === null) {
                        p = val;
                        return p;
                    }
                    if (leftHeaders[x].modificationType === "sum")
                        p += val;
                    else
                        p -= val;
                    return p;
                }, null);
            }
            rows[indexStartx + x][y] = valueOfField;
        }
    }
    return rows;
};
/**
 * Creates excel data from formated data structure of the sheet
 *
 * @param {excelDataObject} dataObject
 * @param {Object[]} leftHeaders
 * @param {(number)[]} historicalStartingCol
 * @param {(number)[]} projectionsStartingCol
 * @return {*}  {Promise<excelDataObject>}
 */
function createData(dataObject, leftHeaders, historicalStartingCol, projectionsStartingCol) {
    return __awaiter(this, void 0, void 0, function* () {
        const workbook = new exceljs_1.default.Workbook();
        const sheet = workbook.addWorksheet('Parsed Sheet', { views: [] });
        let filler = "‎";
        let rows = new Array(leftHeaders.length + 2);
        rows[0] = new Array(Number(dataObject.sheet.max.width.value)).fill(filler);
        ;
        for (let x = 1; x <= 1 + projectionsStartingCol[1] - historicalStartingCol[1]; x++) {
            if (x === 1) {
                rows[0][x] = "Historical";
            }
            else if (x === (1 + projectionsStartingCol[1] - historicalStartingCol[1])) {
                rows[0][x] = "Projections";
            }
            else
                rows[0].push(filler);
        }
        rows = createTypePeriodsSubHeaders(rows, 1, 1, dataObject, "historical", historicalStartingCol, projectionsStartingCol);
        rows = createTypePeriodsSubHeaders(rows, 1, 1, dataObject, "projection", historicalStartingCol, projectionsStartingCol);
        rows = createTypePeriodsData(rows, 2, 1, dataObject, "historical", historicalStartingCol, projectionsStartingCol, leftHeaders);
        rows = createTypePeriodsData(rows, 2, 1, dataObject, "projection", historicalStartingCol, projectionsStartingCol, leftHeaders);
        sheet.addRows(rows);
        sheet.eachRow(function (row, rowNumber) {
            row.font = font;
            row.eachCell(function (cell, colNumber) {
                if (colNumber === 1 && rowNumber > 2 && cell.value) {
                    cell.border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' }
                    };
                    cell.alignment = { vertical: 'middle', horizontal: 'left' };
                    cell.font = Object.assign({ bold: true }, font);
                    sheet.getColumn(colNumber).width = 30;
                }
                if ((historicalStartingCol[0] + 1) === rowNumber) {
                    cell.font = Object.assign({ bold: true }, font);
                }
                if (colNumber >= (historicalStartingCol[1]) && colNumber < (projectionsStartingCol[1])) {
                    cell.numFmt = '$#,##0.00;[Red]-$#,##0.00';
                    cell.fill = {
                        bgColor: { argb: 'e9e9e9' },
                        fgColor: { argb: 'e9e9e9' },
                        pattern: "solid",
                        type: "pattern",
                    };
                    cell.border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' }
                    };
                    cell.alignment = { vertical: 'middle', horizontal: 'center' };
                }
                else if (colNumber >= (projectionsStartingCol[1])) {
                    cell.numFmt = '$#,##0.00;[Red]-$#,##0.00';
                    cell.fill = {
                        bgColor: { argb: 'ccdff3' },
                        fgColor: { argb: 'ccdff3' },
                        pattern: "solid",
                        type: "pattern",
                    };
                    cell.border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' }
                    };
                    cell.alignment = { vertical: 'middle', horizontal: 'center' };
                }
            });
        });
        sheet.properties.defaultColWidth = 22;
        sheet.views = [{
                zoomScale: 65
            }];
        yield workbook.xlsx.writeFile("cleanFinancialStatement.xlsx");
        return dataObject;
    });
}
/**
 * Default entry point
 *
 * @export
 * @param {...[excelDataObject, Object[], (number)[], (number)[]]} args
 * @return {*}  {Promise<excelDataObject>}
 */
function default_1(...args) {
    return __awaiter(this, void 0, void 0, function* () {
        return yield createData(...args);
    });
}
