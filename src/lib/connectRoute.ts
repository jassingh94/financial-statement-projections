import express, { Express, Request, Response } from "express";
import readData from "./readData";
import parseData from "./parseData";
import createExcel from "./createExcel";
import modifyProjections from "./modifyProjections"
type SheetData = (string | number)[][]; // Adjust this type if you expect a different data structure
type WorkbookData = { [key: string]: SheetData };
type SheetMaxWidthData = { [key: string]: number };

const app: Express = express();
const PORT: number = process.env.PORT ? parseInt(process.env.PORT) : 8080;

/**
 * Entry point, creates express engine
 *
 * @export
 */
export default function () {
    app.use(express.json());
    app
        .listen(PORT, "localhost", function () {
            console.log(`Server is running on port ${PORT}.`);
        })
        .on("error", (err: any) => {
            if (err.code === "EADDRINUSE") {
                console.log("Error: address already in use");
            } else {
                console.log(err);
            }
        });

    app.post('/processNClean', (req: Request, res: Response) => {
        let projectionFieldInterests = req?.body?.projectionFieldInterests || {};
        readData('Financial Projections.xlsx')
            .then(({sheetsData, sheetsMaxWidth}) => parseData(sheetsData, sheetsMaxWidth))
            .then(({ parsedData, leftHeaders, historicalStartingCol, projectionsStartingCol }) => modifyProjections(parsedData, leftHeaders, historicalStartingCol, projectionsStartingCol,projectionFieldInterests))
            .then(({ parsedData, leftHeaders, historicalStartingCol, projectionsStartingCol }) => createExcel(parsedData, leftHeaders, historicalStartingCol, projectionsStartingCol))
            .then((parsedData) => {
                res.json({
                    message: "New clean excel process and created 'cleanFinancialStatement.xlsx'",
                    data: parsedData
                })
            }).catch(err => {
                res.sendStatus(500)
                res.send(err);
            })
    });

    app.get('/', (req: Request, res: Response) => {
        res.send('Express + TypeScript Server');
    });
}
