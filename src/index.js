import { XlsxClient } from "./client.js";

export { Cell } from "./cell.js";
export { Column } from "./column.js";
export { Columns } from "./columns.js";
export { Range } from "./range.js";
export { Row, Rows } from "./row.js";
export { Worksheet } from "./worksheet.js";
export { Workbook } from "./workbook.js";
export { ChartCollection } from "./charts.js";
export { PivotCollection } from "./pivots.js";
export * from "./errors.js";
export { StyleCollection, composeStyles, normalizeColor, normalizeStyle } from "./style.js";
export { decryptWorkbookBuffer, encryptWorkbookBuffer, isCompoundFile } from "./encryption.js";

const client = new XlsxClient();

export const createWorkbook = client.create.bind(client);
export const openWorkbook = client.open.bind(client);
export const openWorkbookSync = client.openSync.bind(client);
export const parseWorkbook = client.parse.bind(client);
export const version = "0.4.0";
