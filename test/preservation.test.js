import assert from "node:assert/strict";
import test from "node:test";
import { createWorkbook, parseWorkbook } from "../src/index.js";
import { createZip, extractZip } from "../src/zip.js";

const customPart = Buffer.from('<futureData xmlns="urn:entree:test"><value>keep exactly</value></futureData>');

function templateFixture() {
  const workbook = createWorkbook("Report");
  workbook.sheet().replaceData([["Name", "Total"], ["Original", 12.5]]);
  const files = extractZip(workbook.toBuffer());
  files.set("customXml/item1.xml", customPart);
  files.set(
    "xl/worksheets/sheet1.xml",
    Buffer.from(files.get("xl/worksheets/sheet1.xml").toString("utf8").replace(
      "</worksheet>",
      '<extLst><ext uri="urn:entree:test"><future:payload xmlns:future="urn:entree:future">preserve me</future:payload></ext></extLst></worksheet>'
    ))
  );
  return createZip(Object.fromEntries(files));
}

test("preserves unknown package parts and worksheet extensions without edits", () => {
  const original = extractZip(templateFixture());
  const saved = extractZip(parseWorkbook(templateFixture()).toBuffer());
  assert.deepEqual(saved.get("customXml/item1.xml"), customPart);
  assert.deepEqual(saved.get("xl/worksheets/sheet1.xml"), original.get("xl/worksheets/sheet1.xml"));
});

test("patches a cell while preserving unknown package content", () => {
  const workbook = parseWorkbook(templateFixture());
  workbook.sheet("Report").set("A2", "Updated");
  workbook.sheet("Report").cell("B2").formula("SUM(5, 7)", 12);

  const output = workbook.toBuffer();
  const files = extractZip(output);
  const worksheetXml = files.get("xl/worksheets/sheet1.xml").toString("utf8");
  assert.deepEqual(files.get("customXml/item1.xml"), customPart);
  assert.match(worksheetXml, /future:payload[^>]*>preserve me<\/future:payload>/);

  const reopened = parseWorkbook(output);
  assert.equal(reopened.sheet("Report").get("A2"), "Updated");
  assert.equal(reopened.sheet("Report").get("B2"), 12);
  assert.equal(reopened.sheet("Report").cell("B2").raw.formula, "SUM(5, 7)");
});

test("adds a new style without rebuilding the template style table", () => {
  const original = createWorkbook("Styled");
  original.sheet().set("A1", "Heading");
  original.sheet().cell("A1").style({ bold: true, fill: "#17324D", color: "#FFFFFF" });
  const parsed = parseWorkbook(original.toBuffer());
  parsed.sheet().cell("A1").style({ fontSize: 18, horizontal: "center" });
  const reopened = parseWorkbook(parsed.toBuffer());
  assert.equal(reopened.sheet().cell("A1").raw.style.font.bold, true);
  assert.equal(reopened.sheet().cell("A1").raw.style.font.size, 18);
  assert.equal(reopened.sheet().cell("A1").raw.style.alignment.horizontal, "center");
});

test("patches template layout and protection without dropping extensions", () => {
  const workbook = parseWorkbook(templateFixture());
  workbook.sheet().setColumnWidth("A", 24).setRowHeight(2, 30).merge("A4:B4").autoFilter("A1:B2").protect("sheet pass");
  workbook.protect({ password: "book pass", structure: true });
  const output = workbook.toBuffer();
  const files = extractZip(output);
  assert.match(files.get("xl/worksheets/sheet1.xml").toString("utf8"), /future:payload[^>]*>preserve me/);
  const reopened = parseWorkbook(output);
  assert.equal(reopened.sheet().raw["!cols"][0].width, 24);
  assert.equal(reopened.sheet().raw["!rows"][1].height, 30);
  assert.equal(reopened.sheet().raw["!protection"].sheet, true);
  assert.equal(reopened.source.WorkbookProtection.lockStructure, true);
});

test("adds and replaces hyperlinks in a preserved template", () => {
  const workbook = parseWorkbook(templateFixture());
  workbook.sheet().cell("A2").hyperlink("https://entreepos.com/report", "Open report");
  const first = parseWorkbook(workbook.toBuffer());
  assert.equal(first.sheet().cell("A2").raw.hyperlink.target, "https://entreepos.com/report");
  first.sheet().cell("A2").hyperlink("#Report!A1", "Jump");
  const second = parseWorkbook(first.toBuffer());
  assert.equal(second.sheet().cell("A2").raw.hyperlink.target, "#Report!A1");
});

test("inserts, copies, and deletes rows in a preserved template", () => {
  const source = createWorkbook("Rows");
  source.sheet().replaceData([["Item", "Value"], ["A", 1], ["B", 2], ["C", 3]]);
  source.sheet().cell("C2").formula("B2*2", 2);
  const workbook = parseWorkbook(source.toBuffer());
  workbook.sheet().insertRows(3, 1).copyRow(2, 3).deleteRows(5, 1);
  const reopened = parseWorkbook(workbook.toBuffer());
  assert.deepEqual(reopened.sheet().range("A1:B4").getValues(), [["Item", "Value"], ["A", 1], ["A", 1], ["B", 2]]);
  assert.equal(reopened.sheet().cell("C3").raw.formula, "B3*2");
});

test("adds, renames, removes, and reorders preserved workbook sheets", () => {
  const source = createWorkbook("First");
  source.sheet().set("A1", "keep");
  source.addSheet("Remove", [["old"]]);
  const workbook = parseWorkbook(source.toBuffer());
  workbook.renameSheet("First", "Renamed");
  workbook.addSheet("Added", [["new"]]);
  workbook.removeSheet("Remove");
  const reopened = parseWorkbook(workbook.toBuffer());
  assert.deepEqual(reopened.sheetNames, ["Renamed", "Added"]);
  assert.equal(reopened.sheet("Renamed").get("A1"), "keep");
  assert.equal(reopened.sheet("Added").get("A1"), "new");
});
