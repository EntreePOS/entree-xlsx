import assert from "node:assert/strict";
import test from "node:test";
import { createWorkbook, parseWorkbook } from "../src/index.js";
import { extractZip } from "../src/zip.js";
import { allTags, tagContent } from "../src/xml.js";

test("creates, discovers, updates, and removes a pivot table with cache records", () => {
  const workbook = createWorkbook("Orders");
  workbook.sheet().replaceData([
    ["Region", "Month", "Sales"],
    ["East", "Jan", 10],
    ["East", "Feb", 20],
    ["West", "Jan", 7],
    ["West", "Feb", 13]
  ]);
  workbook.addSheet("Summary");

  const pivot = workbook.pivots.add({
    name: "SalesPivot",
    source: { sheet: "Orders", range: "A1:C5" },
    target: { sheet: "Summary", cell: "A3" },
    rows: ["Region"],
    columns: ["Month"],
    values: [{ field: "Sales", summarize: "sum", name: "Total Sales" }]
  });
  assert.equal(pivot.name, "SalesPivot");
  assert.deepEqual(pivot.rows, ["Region"]);
  assert.equal(workbook.sheet("Summary").get("A4"), "East");
  assert.equal(workbook.sheet("Summary").get("B4"), 10);

  const output = workbook.toBuffer();
  const files = extractZip(output);
  assert.ok(files.has(pivot.id));
  assert.ok([...files.keys()].some((path) => path.startsWith("xl/pivotCache/pivotCacheDefinition")));
  assert.ok([...files.keys()].some((path) => path.startsWith("xl/pivotCache/pivotCacheRecords")));

  const reopened = parseWorkbook(output);
  assert.equal(reopened.pivots.list("Summary")[0].source.range, "A1:C5");
  const updated = reopened.pivots.update("SalesPivot", {
    rows: ["Month"],
    columns: ["Region"]
  });
  assert.deepEqual(updated.rows, ["Month"]);
  assert.deepEqual(updated.columns, ["Region"]);
  assert.equal(reopened.pivots.remove(updated.id), true);
  assert.equal(reopened.pivots.list().length, 0);
});

test("keeps worksheet rows ordered when cells are added above a generated pivot", () => {
  const workbook = createWorkbook("Orders");
  workbook.sheet().replaceData([
    ["Region", "Sales"],
    ["East", 10],
    ["West", 7]
  ]);
  workbook.addSheet("Summary");
  workbook.pivots.add({
    source: { sheet: "Orders", range: "A1:B3" },
    target: { sheet: "Summary", cell: "A5" },
    rows: ["Region"],
    values: [{ field: "Sales", summarize: "sum" }]
  });
  workbook.sheet("Summary").cell("A4").set("Pivot output");

  const files = extractZip(workbook.toBuffer());
  const worksheetXml = files.get("xl/worksheets/sheet2.xml").toString("utf8");
  const rowNumbers = allTags(tagContent(worksheetXml, "sheetData"), "row").map((row) => Number(row.attributes.r));
  assert.deepEqual(rowNumbers, [...rowNumbers].sort((left, right) => left - right));
});
