import assert from "node:assert/strict";
import test from "node:test";
import { createWorkbook, parseWorkbook } from "../src/index.js";
import { extractZip } from "../src/zip.js";

test("creates, discovers, updates, and removes a chart", () => {
  const workbook = createWorkbook("Sales");
  workbook.sheet().setData([
    ["Month", "Revenue", "Orders"],
    ["Jan", 1200, 24],
    ["Feb", 1800, 31],
    ["Mar", 1600, 28]
  ]);

  const created = workbook.charts.add({
    sheet: "Sales",
    type: "column",
    title: "Monthly sales",
    range: "A1:C4",
    position: { from: "E2", to: "M18" }
  });
  assert.equal(created.type, "column");
  assert.equal(created.title, "Monthly sales");

  const reopened = parseWorkbook(workbook.toBuffer());
  assert.equal(reopened.charts.list("Sales").length, 1);
  const updated = reopened.charts.update(created.id, {
    type: "line",
    title: "Sales trend",
    range: "A1:C4",
    position: { from: "F3", to: "N20" }
  });
  assert.equal(updated.type, "line");
  assert.equal(updated.title, "Sales trend");
  assert.deepEqual(updated.position, { from: "F3", to: "N20" });

  const updatedFiles = extractZip(reopened.toBuffer());
  assert.match(updatedFiles.get(created.id).toString("utf8"), /<c:lineChart>/);
  assert.equal(reopened.charts.remove(created.id), true);
  assert.equal(reopened.charts.list("Sales").length, 0);
  assert.equal(extractZip(reopened.toBuffer()).has(created.id), false);
});
