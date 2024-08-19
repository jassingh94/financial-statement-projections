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
Object.defineProperty(exports, "__esModule", { value: true });
exports.default = default_1;
/**
 * Modify projections fields basis on incoming request
 *
 * @param {excelDataObject} dataObject
 * @param {Object[]} leftHeaders
 * @param {(number)[]} historicalStartingCol
 * @param {(number)[]} projectionsStartingCol
 * @param {{
 *     [key: string]: number
 * }} modifyProjectionInterests
 * @return {*}
 */
let modifyProjections = (dataObject, leftHeaders, historicalStartingCol, projectionsStartingCol, modifyProjectionInterests) => {
    var _a;
    let historicalData = dataObject.historical;
    let projectionData = dataObject.projection;
    let latestHistoricalTimestamp = Object.keys(historicalData)[Object.keys(historicalData).length - 1];
    let latestHistoricalData = historicalData[latestHistoricalTimestamp];
    let timeStampsOfProjections = Object.keys(projectionData);
    let projectionFields = Object.keys(modifyProjectionInterests);
    let lastProjection = latestHistoricalData;
    for (let x = 0; x < timeStampsOfProjections.length; x++) {
        for (let y = 0; y < projectionFields.length; y++) {
            let projectionFieldInterest = modifyProjectionInterests[projectionFields[y]];
            let prevValue = Number(((_a = lastProjection[projectionFields[y]]) === null || _a === void 0 ? void 0 : _a.value) || 0);
            dataObject.projection[timeStampsOfProjections[x]][projectionFields[y]].value = prevValue + (prevValue * (projectionFieldInterest / 100));
        }
        lastProjection = dataObject.projection[timeStampsOfProjections[x]];
    }
    return dataObject;
};
/**
 * Starting point, used to modify projections based on request
 *
 * @export
 * @param {excelDataObject} dataObject
 * @param {Object[]} leftHeaders
 * @param {(number)[]} historicalStartingCol
 * @param {(number)[]} projectionsStartingCol
 * @param {{
 *     [key: string]: number
 * }} modifyProjectionInterests
 * @return {*}  {Promise<any>}
 */
function default_1(dataObject, leftHeaders, historicalStartingCol, projectionsStartingCol, modifyProjectionInterests) {
    return __awaiter(this, void 0, void 0, function* () {
        dataObject = modifyProjections(dataObject, leftHeaders, historicalStartingCol, projectionsStartingCol, modifyProjectionInterests);
        return Promise.resolve({ parsedData: dataObject, leftHeaders, historicalStartingCol, projectionsStartingCol });
    });
}
