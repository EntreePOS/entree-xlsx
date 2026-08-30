import { mkdir, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import { createWorkbook, parseWorkbook } from "../src/index.js";

const outputDirectory = resolve("test-output");
await mkdir(outputDirectory, { recursive: true });

function baseWorkbook() {
  const workbook = createWorkbook("Orders");
  workbook.sheet().replaceData([
    ["Region", "Month", "Sales"],
    ["East", "Jan", 10],
    ["East", "Feb", 20],
    ["West", "Jan", 7],
    ["West", "Feb", 13]
  ]);
  workbook.addSheet("Summary");
  return workbook;
}

const paths = {};
const base = baseWorkbook();
paths.base = resolve(outputDirectory, "excel-base.xlsx");
paths.encryptedBase = resolve(outputDirectory, "excel-base-encrypted.xlsx");
await writeFile(paths.base, base.toBuffer());
await writeFile(paths.encryptedBase, base.toBuffer({ password: "Entree123!" }));

const chart = baseWorkbook();
chart.charts.add("Orders", { name: "SalesChart", type: "column", title: "Sales", range: "A1:C5" });
paths.chart = resolve(outputDirectory, "excel-chart.xlsx");
await writeFile(paths.chart, chart.toBuffer());

const pivot = baseWorkbook();
pivot.pivots.add({
  name: "SalesPivot",
  source: { sheet: "Orders", range: "A1:C5" },
  target: { sheet: "Summary", cell: "A3" },
  rows: ["Region"],
  columns: ["Month"],
  values: [{ field: "Sales", summarize: "sum", name: "Total Sales" }]
});
paths.pivot = resolve(outputDirectory, "excel-pivot.xlsx");
await writeFile(paths.pivot, pivot.toBuffer());

const styled = createWorkbook("Style Showcase");
styled.styles
  .define("reportBase", { fontName: "Aptos", fontSize: 11, vertical: "center" })
  .define("reportHeader", {
    bold: true,
    color: "#FFFFFF",
    fill: "#17324D",
    horizontal: "center",
    wrapText: true,
    border: { bottom: { style: "medium", color: "#D6A85F" } }
  }, { extends: "reportBase" })
  .define("money", { numberFormat: "$#,##0.00;[Red]-$#,##0.00" });
const showcase = styled.sheet();
showcase.range("A1:D3").setValues([
  ["Item", "Description", "Price", "Margin"],
  ["Burger", "Wrapped text that demonstrates a reusable report style", 12.5, 0.325],
  ["Fries", "Gradient and grid border examples", -4.25, 0.18]
]);
showcase.range("A1:D1").style("reportHeader");
showcase.range("C2:C3").style("money");
showcase.range("D2:D3").style({ numberFormat: "0.0%" });
showcase.cell("B2").style({ wrapText: true, rotation: -30, font: { italic: true, color: { theme: 4, tint: -0.2 } } });
showcase.cell("B3").style({ fill: { type: "gradient", degree: 90, stops: [{ position: 0, color: "#FFF2CC" }, { position: 1, color: "#D6A85F" }] } });
showcase.range("A1:D3").style({ border: { outline: { style: "medium", color: "#17324D" }, inside: { style: "thin", color: "#B7C9D6" } } });
showcase.cell("A3").style({ superscript: true, underline: "doubleAccounting" });
showcase.cell("D3").style({ protection: { locked: false, hidden: false }, shrinkToFit: true });
showcase.setColumnWidth("A", 16).setColumnWidth("B", 42).setColumnWidth("C", 16).setColumnWidth("D", 14);
paths.styles = resolve(outputDirectory, "excel-styles.xlsx");
await writeFile(paths.styles, styled.toBuffer());
const preservedStyles = parseWorkbook(styled.toBuffer());
preservedStyles.styles.define("note", { italic: true, fill: "#E2F0D9", color: "#375623" });
preservedStyles.sheet().range("A5:D6").setValues([["Reusable", "styles", "also", "work"], ["when", "editing", "a", "template"]]).style("note");
paths.preservedStyles = resolve(outputDirectory, "excel-styles-preserved.xlsx");
await writeFile(paths.preservedStyles, preservedStyles.toBuffer());

const complete = baseWorkbook();
complete.charts.add("Orders", { name: "SalesChart", type: "column", title: "Sales", range: "A1:C5" });
complete.pivots.add({
  name: "SalesPivot",
  source: { sheet: "Orders", range: "A1:C5" },
  target: { sheet: "Summary", cell: "A3" },
  rows: ["Region"],
  columns: ["Month"],
  values: [{ field: "Sales", summarize: "sum", name: "Total Sales" }]
});
complete.protect({ password: "Book123!", structure: true });
complete.sheet("Summary").protect({ password: "Sheet123!" });
paths.complete = resolve(outputDirectory, "excel-complete-encrypted.xlsx");
await writeFile(paths.complete, complete.toBuffer({ password: "Entree123!" }));

process.stdout.write(JSON.stringify(paths));
