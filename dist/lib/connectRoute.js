"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.default = default_1;
const express_1 = __importDefault(require("express"));
const readData_1 = __importDefault(require("./readData"));
const parseData_1 = __importDefault(require("./parseData"));
const createExcel_1 = __importDefault(require("./createExcel"));
const modifyProjections_1 = __importDefault(require("./modifyProjections"));
const app = (0, express_1.default)();
const PORT = process.env.PORT ? parseInt(process.env.PORT) : 8080;
/**
 * Entry point, creates express engine
 *
 * @export
 */
function default_1() {
    app.use(express_1.default.json());
    app
        .listen(PORT, "localhost", function () {
        console.log(`Server is running on port ${PORT}.`);
    })
        .on("error", (err) => {
        if (err.code === "EADDRINUSE") {
            console.log("Error: address already in use");
        }
        else {
            console.log(err);
        }
    });
    app.post('/processNClean', (req, res) => {
        var _a;
        let projectionFieldInterests = ((_a = req === null || req === void 0 ? void 0 : req.body) === null || _a === void 0 ? void 0 : _a.projectionFieldInterests) || {};
        (0, readData_1.default)('Financial Projections.xlsx')
            .then(({ sheetsData, sheetsMaxWidth }) => (0, parseData_1.default)(sheetsData, sheetsMaxWidth))
            .then(({ parsedData, leftHeaders, historicalStartingCol, projectionsStartingCol }) => (0, modifyProjections_1.default)(parsedData, leftHeaders, historicalStartingCol, projectionsStartingCol, projectionFieldInterests))
            .then(({ parsedData, leftHeaders, historicalStartingCol, projectionsStartingCol }) => (0, createExcel_1.default)(parsedData, leftHeaders, historicalStartingCol, projectionsStartingCol))
            .then((parsedData) => {
            res.json({
                message: "New clean excel process and created 'cleanFinancialStatement.xlsx'",
                data: parsedData
            });
        }).catch(err => {
            res.sendStatus(500);
            res.send(err);
        });
    });
    app.get('/', (req, res) => {
        res.send('Express + TypeScript Server');
    });
}
