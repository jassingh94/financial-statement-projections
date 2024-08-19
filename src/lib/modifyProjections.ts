
type excelDataObject = {
    [key: string]: {
        [key: string | number]: {
            [key: string]: {
                [key: string]: number | string
            }
        }
    }
};

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
let modifyProjections = (dataObject: excelDataObject, leftHeaders: Object[], historicalStartingCol: (number)[], projectionsStartingCol: (number)[], modifyProjectionInterests: {
    [key: string]: number
}) => {
    let historicalData = dataObject.historical;
    let projectionData = dataObject.projection;
    let latestHistoricalTimestamp = Object.keys(historicalData)[Object.keys(historicalData).length - 1];
    let latestHistoricalData = historicalData[latestHistoricalTimestamp];
    let timeStampsOfProjections = Object.keys(projectionData);
    let projectionFields =  Object.keys(modifyProjectionInterests);
    let lastProjection = latestHistoricalData;

    for(let x = 0; x < timeStampsOfProjections.length; x++) {
        for(let y = 0; y < projectionFields.length; y++){
            let projectionFieldInterest = modifyProjectionInterests[projectionFields[y]];
            let prevValue = Number(lastProjection[projectionFields[y]]?.value || 0);
            dataObject.projection[timeStampsOfProjections[x]][projectionFields[y]].value = prevValue + (prevValue * (projectionFieldInterest / 100));
        }
        lastProjection = dataObject.projection[timeStampsOfProjections[x]];
    }

    return dataObject;
}

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
export default async function (dataObject: excelDataObject, leftHeaders: Object[], historicalStartingCol: (number)[], projectionsStartingCol: (number)[], modifyProjectionInterests: {
    [key: string]: number
}): Promise<any> {
    dataObject = modifyProjections(dataObject, leftHeaders, historicalStartingCol, projectionsStartingCol, modifyProjectionInterests);
    return Promise.resolve({ parsedData: dataObject, leftHeaders, historicalStartingCol, projectionsStartingCol })
}