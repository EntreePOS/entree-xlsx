import assert from "node:assert/strict";
import test from "node:test";
import { createWorkbook, parseWorkbook } from "../src/index.js";
import { extractZip } from "../src/zip.js";

test("finds cells across worksheets, ranges, and columns", () => {
  const sheet = createWorkbook("Menu").sheet();
  sheet.setData([
    ["Item", "Price"],
    ["Classic Burger", 12.5],
    ["Fries", 4],
    ["Classic Burger", 13]
  ]);

  assert.equal(sheet.find("Fries").address, "A3");
  assert.equal(sheet.find((cell) => cell.value === 12.5).address, "B2");
  assert.deepEqual(sheet.column("a").findAll("Classic Burger").map((cell) => cell.address), ["A2", "A4"]);
  assert.equal(sheet.range("B2:B4").find((cell) => cell.value > 10).address, "B2");
  assert.equal(sheet.column(0).letter, "A");
});

test("styles a whole column without materializing empty cells", () => {
  const workbook = createWorkbook("Menu");
  const sheet = workbook.sheet();
  sheet.setData([["Item", "Price"], ["Classic Burger", 12.5]]);
  sheet.column("B").style({ numberFormat: "$#,##0.00", fill: "#EAF0FF" }).width(18);

  assert.equal(sheet.cell("B2").getStyle().numberFormat, "$#,##0.00");
  assert.equal(sheet.cell("B100").unsafeRaw, undefined);

  const files = extractZip(workbook.toBuffer());
  const worksheetXml = files.get("xl/worksheets/sheet1.xml").toString("utf8");
  assert.match(worksheetXml, /<col min="2" max="2" width="18" customWidth="1" style="\d+"\/>/);

  const reopened = parseWorkbook(workbook.toBuffer());
  assert.equal(reopened.sheet().unsafeRaw["!cols"][1].style.numberFormat, "$#,##0.00");
  assert.equal(reopened.sheet().cell("B2").getStyle().numberFormat, "$#,##0.00");
});

test("preserves column styles when editing an existing workbook", () => {
  const source = createWorkbook("Menu");
  source.sheet().setData([["Item", "Price"], ["Classic Burger", 12.5]]);
  const workbook = parseWorkbook(source.toBuffer());
  workbook.sheet().column("B").style({ numberFormat: "$#,##0.00" });
  const reopened = parseWorkbook(workbook.toBuffer());
  assert.equal(reopened.sheet().unsafeRaw["!cols"][1].style.numberFormat, "$#,##0.00");
  assert.equal(reopened.sheet().cell("B2").getStyle().numberFormat, "$#,##0.00");
});
