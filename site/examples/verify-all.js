import assert from "node:assert/strict";
import { access } from "node:fs/promises";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { openWorkbook } from "@entree_pos/xlsx";

const examplesDirectory = dirname(fileURLToPath(import.meta.url));
const outputDirectory = join(examplesDirectory, "..", "assets", "examples");

const files = [
  "01-add-data.xlsx",
  "02-export-records.xlsx",
  "03-change-cells.xlsx",
  "04-multiple-sheets.xlsx",
  "05-first-style.xlsx",
  "06-reusable-styles.xlsx",
  "07-formulas-and-formats.xlsx",
  "08-layout-and-filters.xlsx",
  "09-edit-a-template.xlsx",
  "10-format-dates-and-percentages.xlsx",
  "11-create-a-chart.xlsx",
  "12-edit-a-chart.xlsx",
  "13-create-a-pivot-table.xlsx",
  "14-protect-a-sheet.xlsx",
  "15-encrypt-a-workbook.xlsx"
];

for (const file of files) {
  const path = join(outputDirectory, file);
  await access(path);
  await openWorkbook(path, file.startsWith("15-") ? { password: "demo" } : undefined);
}

const first = await openWorkbook(join(outputDirectory, files[0]));
assert.deepEqual(first.sheet().toRows(), [
  ["Name", "Age"],
  ["Mina", 28],
  ["Noah", 34]
]);

const records = await openWorkbook(join(outputDirectory, files[1]));
assert.deepEqual(records.sheet().toRecords(), [
  { sku: "BK-101", item: "Classic Burger", stock: 34 },
  { sku: "FR-204", item: "Seasoned Fries", stock: 18 },
  { sku: "DR-305", item: "Cold Brew", stock: 27 }
]);

const changed = await openWorkbook(join(outputDirectory, files[2]));
assert.deepEqual(changed.sheet().toRows(), [
  ["Item", "Stock"],
  ["Classic Burger", 34],
  ["Seasoned Fries", 24],
  ["Cold Brew", 27]
]);

const sheets = await openWorkbook(join(outputDirectory, files[3]));
assert.deepEqual(sheets.sheetNames, ["Products", "Categories"]);

console.log(`Verified ${files.length} tutorial workbooks.`);
