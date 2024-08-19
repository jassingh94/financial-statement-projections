import * as XLSX from 'xlsx';

type SheetData = (string | number)[][];
type SheetsData = { [key: string]: SheetData };
type SheetsMaxWidth = { [key: string]: number };

// Create a type that combines sheets data and max width info
type WorkbookData = {
    sheetsData: SheetsData;
    sheetsMaxWidth: SheetsMaxWidth;
};

/**
 * Creates workbook data from file and returns rows per sheet
 *
 * @param {XLSX.WorkBook} wb
 * @return {*}  {WorkbookData}
 */
function destructData(wb: XLSX.WorkBook): WorkbookData {
    const sheets: string[] = wb.SheetNames;
    const sheetsData: SheetsData = {};
    const sheetsMaxWidth: SheetsMaxWidth = {};

    if (!sheets) {
        throw new Error("Unable to read sheets");
    }

    sheets.forEach((sheet: string) => {
        sheetsData[sheet] = XLSX.utils.sheet_to_json(wb.Sheets[sheet], {
            header: 1,
            raw: false
        });
        sheetsMaxWidth[sheet] = XLSX.utils.decode_range(wb.Sheets[sheet]['!ref'] || "")?.e?.c || 50;
    });

    return { sheetsData, sheetsMaxWidth };
}

export default async (fileName: string): Promise<WorkbookData> => {
    const wb = await XLSX.readFile(fileName, {});
    return destructData(wb);
}
