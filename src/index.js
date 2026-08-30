import { XlsxClient } from "./client.js";

export { Cell } from "./cell.js";
export { Range } from "./range.js";
export { Worksheet } from "./worksheet.js";
export { Workbook } from "./workbook.js";
export { ChartCollection } from "./charts.js";
export { PivotCollection } from "./pivots.js";
export { XlsxClient } from "./client.js";
export * from "./errors.js";
export { StyleCollection, composeStyles, normalizeColor, normalizeStyle } from "./style.js";
export { decryptWorkbookBuffer, encryptWorkbookBuffer, isCompoundFile } from "./encryption.js";

export const xlsx = new XlsxClient();
export const createWorkbook = xlsx.create.bind(xlsx);
export const openWorkbook = xlsx.open.bind(xlsx);
export const openWorkbookSync = xlsx.openSync.bind(xlsx);
export const parseWorkbook = xlsx.parse.bind(xlsx);
export const version = "0.1.0";

export default xlsx;
