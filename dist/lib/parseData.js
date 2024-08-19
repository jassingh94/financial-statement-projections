"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
let normalizeNumberValues = (value) => {
    return Number(value.toString().replace(/[\,$]/g, ""));
};
function parseMonthYearToTimeStamp(monthYear) {
    // Split the input string into month and year
    const [monthName, year] = monthYear.split(' ');
    const monthMap = {
        'January': 0, 'February': 1, 'March': 2, 'April': 3, 'May': 4,
        'June': 5, 'July': 6, 'August': 7, 'September': 8, 'October': 9,
        'November': 10, 'December': 11
    };
    const month = monthMap[monthName];
    return new Date(Number(year), month, 1).getTime();
}
/**
 * Fetches data basis on the given type (historical / projections)
 *
 * @param {SheetData} data
 * @param {((number | string)[])} historicalStartingCol
 * @param {((number | string)[])} projectionsStartingCol
 * @param {{
 *         [key: string] : any
 *     }[]} leftHeaders
 * @param {Object[]} subHeaders
 * @param {number} sheetMaxWidth
 * @return {*}
 */
let retrieveDataBasisOnHeaders = (data, historicalStartingCol, projectionsStartingCol, leftHeaders, subHeaders, sheetMaxWidth) => {
    let finalData = {
        "historical": {},
        "projection": {},
        sheet: {
            max: {
                width: {
                    value: sheetMaxWidth
                }
            }
        }
    };
    subHeaders.forEach((subHeader) => {
        leftHeaders.forEach((leftHeader) => {
            let valueType = "";
            if (subHeader.y >= historicalStartingCol[1] && subHeader.y < projectionsStartingCol[1]) {
                valueType = "historical";
            }
            else if (subHeader.y >= projectionsStartingCol[1]) {
                valueType = "projection";
            }
            let monthVal = parseMonthYearToTimeStamp(subHeader.name);
            if (!finalData[valueType][monthVal])
                finalData[valueType][monthVal] = {};
            finalData[valueType][monthVal][leftHeader.name] = {
                value: normalizeNumberValues(data[leftHeader.x][subHeader.y]),
                month: subHeader.name
            };
        });
    });
    return finalData;
};
exports.default = (data, sheetsMaxWidth) => {
    let firstSheetName = Object.keys(data)[0];
    let fData = data[firstSheetName];
    let sheetMaxWidth = sheetsMaxWidth[firstSheetName];
    let historicalStartingCol = [];
    let projectionsStartingCol = [];
    let freezeXForMainHeaders = fData.length;
    let freezeYForMainHeaders = ((sheetMaxWidth || 49) + 1);
    let startYForLeftHeaders = ((sheetMaxWidth || 49) + 1);
    let freezeXForLeftHeaders = fData.length;
    let leftHeaders = [];
    let subHeaders = [];
    /**
     * Identify start col for historical and projections
     *
     * @param {SheetData} data
     */
    let setHistoricalAndProjectionsCol = (data) => {
        for (let x = 0; x < data.length && x < freezeXForMainHeaders; x++) {
            for (let y = 0; y < data[x].length && y < ((sheetMaxWidth || 49) + 1); y++) {
                if (data[x][y] === "Historical" && !historicalStartingCol[0]) {
                    historicalStartingCol = [x, y];
                    freezeXForMainHeaders = x;
                    freezeYForMainHeaders = y;
                }
                if (data[x][y] === "Projections") {
                    projectionsStartingCol = [x, y];
                }
            }
        }
    };
    /**
     * Identify and extract left headers / fields
     *
     * @param {SheetData} data
     * @param {number} freezeX
     */
    let setLeftHeaders = (data, freezeX) => {
        let leftHeadersSetStarted = false;
        let modifiedBy = [];
        let modifiedByTotal = [];
        for (let x = freezeX + 1; x < data.length; x++) {
            for (let y = 0; y < freezeYForMainHeaders && y <= startYForLeftHeaders; y++) {
                if (data[x][y]) {
                    if (!leftHeadersSetStarted) {
                        startYForLeftHeaders = y;
                        freezeXForLeftHeaders = x;
                        leftHeadersSetStarted = true;
                    }
                    let addData = {
                        x,
                        y,
                        name: data[x][y]
                    };
                    leftHeaders.push(addData);
                }
            }
        }
        leftHeaders.forEach((val, index) => {
            if (val.name.toString().match(/Total /)) {
                leftHeaders[index].modifiedBy = modifiedBy;
                leftHeaders[index].modificationType = "sum";
                leftHeaders[index].thisIndex = index;
                modifiedByTotal.push(index);
                modifiedBy = [];
            }
            else if (val.name.toString().match(/Net /)) {
                leftHeaders[index].modifiedBy = modifiedByTotal;
                leftHeaders[index].modificationType = "subtract";
                leftHeaders[index].thisIndex = index;
                modifiedByTotal = [];
            }
            else {
                modifiedBy.push(index);
            }
        });
    };
    /**
     * identify and extract sub headers (months)
     *
     * @param {SheetData} data
     * @param {number} freezeX
     * @param {number} freezeXForLeftHeaders
     * @param {number} startYForLeftHeaders
     */
    let setSubHeaders = (data, freezeX, freezeXForLeftHeaders, startYForLeftHeaders) => {
        let leftHeadersSetStarted = false;
        for (let x = freezeX + 1; x < freezeXForLeftHeaders; x++) {
            for (let y = startYForLeftHeaders; y < ((sheetMaxWidth || 49) + 1); y++) {
                if (data[x][y]) {
                    subHeaders.push({
                        x,
                        y,
                        name: data[x][y]
                    });
                }
            }
        }
    };
    setHistoricalAndProjectionsCol(fData);
    setLeftHeaders(fData, freezeXForMainHeaders);
    setSubHeaders(fData, freezeXForMainHeaders, freezeXForLeftHeaders, startYForLeftHeaders);
    return Promise.resolve({
        leftHeaders,
        subHeaders,
        historicalStartingCol,
        projectionsStartingCol,
        parsedData: retrieveDataBasisOnHeaders(fData, historicalStartingCol, projectionsStartingCol, leftHeaders, subHeaders, ((sheetMaxWidth || 49) + 1))
    });
};
