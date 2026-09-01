import assert from "node:assert/strict";
import { parseWorkbook } from "../../src/index.js";
import { executePlayground } from "./runtime.js";

const source = `
const workbook = createWorkbook("Inventory");
const sheet = workbook.sheet();
sheet.setData([
  ["Item", "Stock", "Status"],
  ["Classic Burger", 18, "Low"],
  ["Fries", 42, "Ready"]
]);
sheet.row(1).style({ bold: true, fill: "#2457C5", color: "#FFFFFF" }).height(26);
sheet.rows("2:3").height(22);
sheet.column("B").style({ numberFormat: "$#,##0.00" }).width(16);
sheet.columns("A:C").forEach((cell) => console.log(cell.address));
sheet.cell("D2").formula("B2*2", 36);
await workbook.save("inventory.xlsx");
`;

const result = await executePlayground(source);
const workbook = parseWorkbook(result.bytes);
const sheet = workbook.sheet();

assert.equal(result.workbook.name, "inventory.xlsx");
assert.equal(sheet.get("A2"), "Classic Burger");
assert.equal(sheet.get("D2"), 36);
assert.equal(sheet.cell("A1").getStyle().font.bold, true);
assert.equal(sheet.unsafeRaw["!rows"][0].height, 26);
assert.equal(sheet.unsafeRaw["!cols"][1].width, 16);
assert.ok(result.logs.length >= 6);

console.log(`Verified browser playground output (${result.bytes.length} bytes).`);
