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

test("selects single and multiple rows and columns", () => {
  const sheet = createWorkbook("Menu").sheet();
  sheet.setData([
    ["Item", "Price", "Class"],
    ["Classic Burger", 12.5, "Burger"],
    ["Fries", 4, "Side"],
    ["Veggie Burger", 11, "Burger"]
  ]);

  assert.equal(sheet.row(2).find("Classic Burger").address, "A2");
  assert.deepEqual(sheet.rows("2:4").findAll((cell) => cell.value === "Burger").map((cell) => cell.address), ["C2", "C4"]);
  assert.deepEqual(sheet.rows([2, 4]).numbers, [2, 4]);
  assert.deepEqual(sheet.columns("A:C").letters, ["A", "B", "C"]);
  assert.deepEqual(sheet.columns(["A", "C"]).indexes, [0, 2]);
  assert.deepEqual(sheet.columns("A:B").findAll((cell) => typeof cell.value === "number").map((cell) => cell.address), ["B2", "B3", "B4"]);
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

test("styles rows and groups without materializing empty cells", () => {
  const workbook = createWorkbook("Menu");
  const sheet = workbook.sheet();
  sheet.setData([
    ["Item", "Price", "Class"],
    ["Classic Burger", 12.5, "Burger"],
    ["Fries", 4, "Side"]
  ]);
  sheet.row(1).style({ font: { bold: true } }).height(26);
  sheet.rows("2:3").style({ fill: "#EAF0FF" }).height(22);
  sheet.columns("A:C").style({ alignment: { horizontal: "center" } }).width(18);

  assert.equal(sheet.cell("A1").getStyle().font.bold, true);
  assert.equal(sheet.cell("B2").getStyle().fill.foreground.rgb, "EAF0FF");
  assert.equal(sheet.cell("C3").getStyle().alignment.horizontal, "center");
  assert.equal(sheet.cell("A100").unsafeRaw, undefined);

  const buffer = workbook.toBuffer();
  const files = extractZip(buffer);
  const worksheetXml = files.get("xl/worksheets/sheet1.xml").toString("utf8");
  assert.match(worksheetXml, /<row r="1"[^>]*ht="26"[^>]*customHeight="1"[^>]*s="\d+"[^>]*customFormat="1"/);
  assert.match(worksheetXml, /<col min="1" max="1" width="18" customWidth="1" style="\d+"\/>/);

  const reopened = parseWorkbook(buffer).sheet();
  assert.equal(reopened.unsafeRaw["!rows"][0].style.font.bold, true);
  assert.equal(reopened.unsafeRaw["!rows"][1].style.fill.foreground.rgb, "EAF0FF");
  assert.equal(reopened.unsafeRaw["!rows"][0].height, 26);
  assert.equal(reopened.unsafeRaw["!cols"][2].style.alignment.horizontal, "center");
});

test("preserves row styles when editing an existing workbook", () => {
  const source = createWorkbook("Menu");
  source.sheet().setData([["Item", "Price"], ["Classic Burger", 12.5]]);
  const workbook = parseWorkbook(source.toBuffer());
  workbook.sheet().row(2).style({ font: { italic: true } }).height(24);
  const reopened = parseWorkbook(workbook.toBuffer()).sheet();
  assert.equal(reopened.unsafeRaw["!rows"][1].style.font.italic, true);
  assert.equal(reopened.unsafeRaw["!rows"][1].height, 24);
  assert.equal(reopened.cell("A2").getStyle().font.italic, true);
});
