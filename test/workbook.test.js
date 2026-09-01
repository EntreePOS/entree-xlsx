import assert from "node:assert/strict";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { createWorkbook, openWorkbook, parseWorkbook, SheetNotFoundError, version } from "../src/index.js";

test("creates, styles, serializes, and reopens an XLSX workbook", () => {
  const workbook = createWorkbook("Orders");
  const sheet = workbook.sheet("Orders");
  sheet.setData([
    ["Order", "Total", "Created"],
    [1001, 12.5, new Date("2026-08-30T12:00:00.000Z")],
    [1002, 8.25, new Date("2026-08-31T13:30:00.000Z")]
  ]);
  sheet.range("A1:C1").style({ bold: true, fontSize: 14, color: "#FFFFFF", fill: "#17324D" });
  sheet.range("B2:B3").style({ numberFormat: "$0.00" });
  sheet.cell("D2").formula("B2*2", 25).numberFormat("$0.00");
  sheet.cell("A2").hyperlink("https://example.com/orders/1001", "Open order");
  sheet.autoFilter().autoFit();
  sheet.merge("A5:C5").set("A5", "End");

  const buffer = workbook.toBuffer();
  assert.equal(buffer.readUInt32LE(0), 0x04034B50);
  const reopened = parseWorkbook(buffer);
  assert.deepEqual(reopened.sheetNames, ["Orders"]);
  assert.equal(reopened.sheet("Orders").get("A2"), 1001);
  assert.equal(reopened.sheet("Orders").get("B3"), 8.25);
  assert.equal(reopened.sheet("Orders").get("D2"), 25);
  assert.equal(reopened.sheet("Orders").cell("D2").unsafeRaw.formula, "B2*2");
  assert.equal(reopened.sheet("Orders").cell("A2").unsafeRaw.hyperlink.target, "https://example.com/orders/1001");
  assert.equal(reopened.sheet("Orders").cell("A1").unsafeRaw.style.font.size, 14);
  assert.equal(reopened.sheet("Orders").get("C2").toISOString(), "2026-08-30T12:00:00.000Z");
  assert.equal(reopened.sheet("Orders").unsafeRaw["!merges"].length, 1);
  assert.equal(reopened.sheet("Orders").unsafeRaw["!autofilter"], "A1:D3");
});

test("supports object rows and contextual sheet errors", () => {
  const workbook = createWorkbook("People");
  workbook.sheet().setData([{ name: "Ada", active: true }, { name: "Linus", active: false }]);
  workbook.sheet().appendData([{ active: true, name: "Grace" }]);
  assert.deepEqual(workbook.sheet().toRecords(), [
    { name: "Ada", active: true },
    { name: "Linus", active: false },
    { name: "Grace", active: true }
  ]);
  assert.equal(workbook.findSheet("People")?.name, "People");
  assert.equal(workbook.findSheet("Missing"), undefined);
  assert.throws(() => workbook.sheet("Missing"), SheetNotFoundError);
  assert.throws(() => workbook.sheet().setData([["valid"], { invalid: true }]), TypeError);
  assert.equal(workbook.sheet().get("A2"), "Ada");
  assert.equal(version, "0.4.1");
});

test("saves and opens a workbook from disk", async () => {
  const directory = await mkdtemp(join(tmpdir(), "entree-xlsx-"));
  const path = join(directory, "test.xlsx");
  try {
    const workbook = createWorkbook("Data");
    workbook.properties = { title: "Round trip", author: "Entree POS" };
    workbook.sheet().setData([["value"], [42]]);
    await workbook.save(path);
    assert.ok((await readFile(path)).length > 1000);
    const reopened = await openWorkbook(path);
    assert.equal(reopened.sheet().get("A2"), 42);
    assert.equal(reopened.properties.title, "Round trip");
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});
