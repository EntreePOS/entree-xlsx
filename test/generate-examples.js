import { mkdir, rm } from "node:fs/promises";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { createWorkbook } from "@entree_pos/xlsx";

const outputDirectory = join(dirname(fileURLToPath(import.meta.url)), "showcase");
await mkdir(outputDirectory, { recursive: true });

const previousExamples = [
  "01-styled-sales-dashboard.xlsx",
  "02-pivot-table-and-chart.xlsx",
  "03-protected-budget-template.xlsx",
  "04-password-encrypted-report.xlsx"
];
await Promise.all(previousExamples.map((name) => rm(join(outputDirectory, name), { force: true })));

function defineStyles(workbook) {
  workbook.styles
    .define("base", {
      fontName: "Aptos",
      fontSize: 11,
      vertical: "center"
    })
    .define("title", {
      fontName: "Aptos Display",
      fontSize: 20,
      bold: true,
      color: "#FFFFFF",
      fill: "#17324D",
      horizontal: "left"
    }, { extends: "base" })
    .define("subtitle", {
      italic: true,
      color: "#44546A",
      wrapText: true
    }, { extends: "base" })
    .define("header", {
      bold: true,
      color: "#FFFFFF",
      fill: "#2F75B5",
      horizontal: "center",
      wrapText: true
    }, { extends: "base" })
    .define("money", {
      numberFormat: "$#,##0.00;[Red]-$#,##0.00"
    }, { extends: "base" })
    .define("percent", {
      numberFormat: "0.0%"
    }, { extends: "base" })
    .define("date", {
      numberFormat: "yyyy-mm-dd"
    }, { extends: "base" })
    .define("input", {
      fill: "#FFF2CC",
      editable: true
    }, { extends: "base" })
    .define("output", {
      fill: "#D9EAF7",
      bold: true
    }, { extends: "base" })
    .define("note", {
      italic: true,
      color: "#666666",
      wrapText: true
    }, { extends: "base" });
}

function addTitle(sheet, title, subtitle, endColumn) {
  sheet.merge(`A1:${endColumn}1`).set("A1", title);
  sheet.range(`A1:${endColumn}1`).style("title");
  sheet.merge(`A2:${endColumn}2`).set("A2", subtitle);
  sheet.range(`A2:${endColumn}2`).style("subtitle");
  sheet.setRowHeight(1, 34).setRowHeight(2, 28);
}

async function createStylesExample() {
  const workbook = createWorkbook("Styles");
  defineStyles(workbook);
  workbook.properties = {
    title: "@entree_pos/xlsx Styles Example",
    subject: "Simple examples of reusable named styles and direct style objects",
    author: "Entree POS"
  };
  const sheet = workbook.sheet();
  addTitle(sheet, "1. STYLES", "Reusable named styles keep workbook formatting clear and consistent.", "D");

  sheet.range("A4:D12").setValues([
    ["Feature", "Result", "API idea", "What it shows"],
    ["Named header", "Header text", 'style("header")', "Reusable font, fill, and alignment"],
    ["Currency", 12345.67, 'style("money")', "Real number with an Excel number format"],
    ["Percentage", 0.275, 'style("percent")', "A numeric percentage, not text"],
    ["Date", new Date("2026-08-30T12:00:00Z"), 'style("date")', "A typed date value"],
    ["Wrapped text", "Long text can wrap cleanly inside a cell.", "{ wrapText: true }", "Alignment and wrapping"],
    ["Fill + border", "Custom style", "{ fill, border }", "Direct style objects"],
    ["Unlocked input", "Editable", 'style("input")', "Cell protection metadata"],
    ["Copied style", "Copied", "copyStyleFrom()", "Reuse an existing cell style"]
  ]);
  sheet.range("A4:D4").style("header");
  sheet.cell("B5").style("header");
  sheet.cell("B6").style("money");
  sheet.cell("B7").style("percent");
  sheet.cell("B8").style("date");
  sheet.cell("B9").style({ wrapText: true, fill: "#F2F2F2" });
  sheet.cell("B10").style({
    bold: true,
    color: "#17324D",
    fill: "#D9EAF7",
    horizontal: "center",
    border: { outline: { style: "medium", color: "#2F75B5" } }
  });
  sheet.cell("B11").style("input");
  sheet.cell("B12").copyStyleFrom("B10");
  sheet.range("A4:D12").style({
    border: { insideHorizontal: { style: "thin", color: "#D9E2F3" } }
  });
  sheet.setColumnWidth("A", 20).setColumnWidth("B", 26).setColumnWidth("C", 24).setColumnWidth("D", 38);
  sheet.setRowHeight(4, 28).setRowHeight(9, 34);

  await workbook.save(join(outputDirectory, "01-styles.xlsx"));
}

async function createChartExample() {
  const workbook = createWorkbook("Sales");
  defineStyles(workbook);
  workbook.properties = {
    title: "@entree_pos/xlsx Chart Example",
    subject: "A small data table and a native Excel chart",
    author: "Entree POS"
  };
  const sheet = workbook.sheet();
  addTitle(sheet, "2. CHART", "Create a native Excel chart from a normal worksheet range.", "J");

  sheet.range("A4:B9").setValues([
    ["Month", "Sales"],
    ["Jan", 12000],
    ["Feb", 14500],
    ["Mar", 13800],
    ["Apr", 17200],
    ["Total", null]
  ]);
  sheet.range("A4:B4").style("header");
  sheet.range("B5:B9").style("money");
  sheet.cell("B9").formula("SUM(B5:B8)", 57500).style(["money", "output"]);
  sheet.cell("A9").style("output");
  sheet.merge("A11:C13").set("A11", "This file demonstrates: values, a SUM formula, number formatting, and a native line chart.");
  sheet.range("A11:C13").style("note");
  sheet.setColumnWidth("A", 16).setColumnWidth("B", 16);
  for (const column of ["C", "D", "E", "F", "G", "H", "I", "J"]) sheet.setColumnWidth(column, 11);
  workbook.charts.add("Sales", {
    type: "line",
    title: "Monthly Sales ($)",
    legend: false,
    series: [{ name: "Sales", categories: "A5:A8", values: "B5:B8" }],
    position: { from: "D4", to: "J14" }
  });

  await workbook.save(join(outputDirectory, "02-chart.xlsx"));
}

async function createPivotExample() {
  const workbook = createWorkbook("Orders");
  workbook.addSheet("Summary");
  defineStyles(workbook);
  workbook.properties = {
    title: "@entree_pos/xlsx Pivot Table Example",
    subject: "A small source table and a native PivotTable",
    author: "Entree POS"
  };

  const orders = workbook.sheet("Orders");
  addTitle(orders, "3. PIVOT TABLE — SOURCE", "Six rows are enough to demonstrate a real PivotTable and pivot cache.", "C");
  orders.range("A4:C10").setValues([
    ["Region", "Category", "Sales"],
    ["East", "Food", 100],
    ["East", "Drinks", 50],
    ["Central", "Food", 90],
    ["Central", "Drinks", 45],
    ["West", "Food", 80],
    ["West", "Drinks", 40]
  ]);
  orders.range("A4:C4").style("header");
  orders.range("C5:C10").style("money");
  orders.autoFilter("A4:C10");
  orders.setColumnWidth("A", 17).setColumnWidth("B", 17).setColumnWidth("C", 16);

  const summary = workbook.sheet("Summary");
  addTitle(summary, "3. PIVOT TABLE — RESULT", "The library writes the PivotTable definition, cache definition, and cache records.", "D");
  summary.range("A4:D4").setValues([["Pivot output", "Food", "Drinks", "Grand Total"]]).style("header");
  workbook.pivots.add({
    name: "SalesByRegion",
    source: { sheet: "Orders", range: "A4:C10" },
    target: { sheet: "Summary", cell: "A5" },
    rows: ["Region"],
    columns: ["Category"],
    values: [{ field: "Sales", summarize: "sum", name: "Sales" }]
  });
  const pivotSheet = workbook.sheet("Summary");
  pivotSheet.range("A5:D5").style({ bold: true, color: "#17324D", fill: "#D9EAF7" });
  pivotSheet.range("B6:D9").style("money");
  pivotSheet.range("A9:D9").style("output");
  pivotSheet.range("A5:D9").style({
    border: { insideHorizontal: { style: "thin", color: "#D9E2F3" } }
  });
  pivotSheet.setColumnWidth("A", 18).setColumnWidth("B", 17).setColumnWidth("C", 17).setColumnWidth("D", 18);

  await workbook.save(join(outputDirectory, "03-pivot-table.xlsx"));
}

async function createProtectionAndEncryptionExample() {
  const workbook = createWorkbook("Invoice");
  defineStyles(workbook);
  workbook.properties = {
    title: "@entree_pos/xlsx Protection and Encryption Example",
    subject: "Unlocked inputs, protected formulas, and AES-256 Office encryption",
    author: "Entree POS"
  };
  const sheet = workbook.sheet();
  addTitle(sheet, "4. PROTECTION + ENCRYPTION", "Password for opening and editing this demo: demo", "D");
  sheet.range("A4:D7").setValues([
    ["Item", "Quantity", "Price", "Total"],
    ["Coffee", 2, 3.5, null],
    ["Lunch", 1, 12, null],
    ["Grand Total", null, null, null]
  ]);
  sheet.range("A4:D4").style("header");
  sheet.range("A5:C6").style("input");
  sheet.range("C5:C6").style(["money", "input"]);
  sheet.cell("D5").formula("B5*C5", 7).style(["money", "output"]);
  sheet.cell("D6").formula("B6*C6", 12).style(["money", "output"]);
  sheet.cell("D7").formula("SUM(D5:D6)", 19).style(["money", "output"]);
  sheet.range("A7:C7").style("output");
  sheet.merge("A9:D10").set("A9", "Yellow cells are editable. Blue cells contain protected formulas. The whole XLSX file is encrypted with AES-256 Agile Office encryption.");
  sheet.range("A9:D10").style("note");
  sheet.setColumnWidth("A", 22).setColumnWidth("B", 14).setColumnWidth("C", 14).setColumnWidth("D", 16);
  sheet.protect({ password: "demo", selectUnlockedCells: true, formatCells: false });
  workbook.protect({ password: "demo", structure: true });

  await workbook.save(join(outputDirectory, "04-protection-and-encryption.xlsx"), { password: "demo" });
}

await createStylesExample();
await createChartExample();
await createPivotExample();
await createProtectionAndEncryptionExample();

console.log(`Created four simple showcase workbooks in ${outputDirectory}`);
