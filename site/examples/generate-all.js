import { mkdir, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { createWorkbook, openWorkbook } from "@entree_pos/xlsx";

const siteDirectory = join(dirname(fileURLToPath(import.meta.url)), "..");
const outputDirectory = join(siteDirectory, "assets", "examples");
const previewDirectory = join(siteDirectory, "..", ".codex-tmp-tutorial");

await mkdir(outputDirectory, { recursive: true });
await mkdir(previewDirectory, { recursive: true });

const palette = {
  ink: "#17202A",
  blue: "#2457C5",
  blueSoft: "#EAF0FF",
  green: "#2E7D5B",
  greenSoft: "#E6F4ED",
  amber: "#A15C00",
  amberSoft: "#FFF2D8",
  line: "#CBD5E1",
  muted: "#5F6B7A",
  white: "#FFFFFF"
};

function addReportTitle(sheet, title, subtitle, endColumn) {
  sheet.range(`A1:${endColumn}1`).merge();
  sheet.cell("A1").set(title).style({
    bold: true,
    fontSize: 18,
    color: palette.white,
    fill: palette.ink,
    vertical: "center"
  });
  sheet.setRowHeight(1, 32);
  sheet.range(`A2:${endColumn}2`).merge();
  sheet.cell("A2").set(subtitle).style({
    color: palette.muted,
    italic: true,
    vertical: "center"
  });
  sheet.setRowHeight(2, 24);
}

function addHeader(sheet, range) {
  sheet.range(range).style({
    bold: true,
    color: palette.white,
    fill: palette.blue,
    horizontal: "center",
    vertical: "center",
    border: { bottom: { style: "medium", color: palette.ink } }
  });
}

function addGrid(sheet, range) {
  sheet.range(range).style({
    border: {
      outline: { style: "thin", color: palette.line },
      inside: { style: "thin", color: palette.line }
    }
  });
}

async function createBasicWorkbook() {
  const workbook = createWorkbook("Inventory");
  const sheet = workbook.sheet();

  sheet.setData([
    { sku: "BK-101", item: "Classic Burger", stock: 34 },
    { sku: "FR-204", item: "Seasoned Fries", stock: 18 },
    { sku: "DR-305", item: "Cold Brew", stock: 27 },
    { sku: "DS-410", item: "Chocolate Cake", stock: 9 }
  ]);

  addHeader(sheet, "A1:C1");
  addGrid(sheet, "A1:C5");
  sheet.range("C2:C5").style({ numberFormat: "#,##0", horizontal: "right" });
  sheet.autoFit({ min: 12, max: 28, padding: 3 });

  await workbook.save(join(outputDirectory, "01-create-workbook.xlsx"));
}

async function createStyledWorkbook() {
  const workbook = createWorkbook("Style Guide");
  const sheet = workbook.sheet();

  workbook.styles
    .define("label", {
      bold: true,
      color: palette.ink,
      fill: palette.blueSoft,
      vertical: "center"
    })
    .define("money", { numberFormat: "$#,##0.00;[Red]-$#,##0.00" })
    .define("success", {
      bold: true,
      color: palette.green,
      fill: palette.greenSoft,
      horizontal: "center"
    })
    .define("warning", {
      bold: true,
      color: palette.amber,
      fill: palette.amberSoft,
      horizontal: "center"
    });

  addReportTitle(sheet, "Reusable workbook styles", "Define once, then apply by name to cells or ranges.", "D");
  sheet.range("A4:D8").setValues([
    ["Style", "Example", "Value", "Purpose"],
    ["label", "Product", "Cold Brew", "Section labels"],
    ["money", "Unit price", 4.5, "Currency values"],
    ["success", "Inventory", "Ready", "Positive state"],
    ["warning", "Inventory", "Low stock", "Needs attention"]
  ]);
  addHeader(sheet, "A4:D4");
  addGrid(sheet, "A4:D8");
  sheet.cell("A5").style("label");
  sheet.cell("C6").style("money");
  sheet.cell("C7").style("success");
  sheet.cell("C8").style("warning");
  sheet.setColumnWidth("A", 18).setColumnWidth("B", 18).setColumnWidth("C", 18).setColumnWidth("D", 25);

  await workbook.save(join(outputDirectory, "02-reusable-styles.xlsx"));
}

async function createFormulaWorkbook() {
  const workbook = createWorkbook("Invoice");
  const sheet = workbook.sheet();

  addReportTitle(sheet, "Invoice totals", "Formulas remain formulas when the workbook opens in Excel.", "D");
  sheet.range("A4:D8").setValues([
    ["Item", "Quantity", "Price", "Total"],
    ["Classic Burger", 2, 12.5, null],
    ["Seasoned Fries", 1, 4, null],
    ["Cold Brew", 2, 3.5, null],
    ["Grand total", null, null, null]
  ]);
  addHeader(sheet, "A4:D4");
  addGrid(sheet, "A4:D8");
  sheet.cell("D5").formula("B5*C5", 25);
  sheet.cell("D6").formula("B6*C6", 4);
  sheet.cell("D7").formula("B7*C7", 7);
  sheet.cell("D8").formula("SUM(D5:D7)", 36).style({ bold: true, fill: palette.greenSoft });
  sheet.range("C5:D8").style({ numberFormat: "$#,##0.00" });
  sheet.range("A8:C8").style({ bold: true, fill: palette.greenSoft });
  sheet.setColumnWidth("A", 24).setColumnWidth("B", 12).setColumnWidth("C", 14).setColumnWidth("D", 15);

  await workbook.save(join(outputDirectory, "03-formulas-and-formats.xlsx"));
}

async function createLayoutWorkbook() {
  const workbook = createWorkbook("Orders");
  const sheet = workbook.sheet();

  addReportTitle(sheet, "Open orders", "A readable report with merged titles, filters, widths, and links.", "E");
  sheet.range("A4:E9").setValues([
    ["Order", "Customer", "Status", "Total", "Details"],
    [1041, "Ada Rivera", "Ready", 42.5, null],
    [1042, "Noah Patel", "Preparing", 31, null],
    [1043, "Mina Park", "Ready", 26.75, null],
    [1044, "Owen Brooks", "New", 18.5, null],
    [1045, "Lena Ortiz", "Preparing", 54, null]
  ]);
  addHeader(sheet, "A4:E4");
  addGrid(sheet, "A4:E9");
  sheet.range("D5:D9").style({ numberFormat: "$#,##0.00" });
  for (let row = 5; row <= 9; row += 1) {
    sheet.cell(`E${row}`).set("Open order").hyperlink(`https://example.com/orders/${sheet.cell(`A${row}`).value}`, "View order details");
  }
  sheet.autoFilter("A4:E9");
  sheet.setColumnWidth("A", 12).setColumnWidth("B", 22).setColumnWidth("C", 16).setColumnWidth("D", 14).setColumnWidth("E", 18);

  await workbook.save(join(outputDirectory, "04-layout-and-filters.xlsx"));
}

async function createTemplateWorkbook() {
  const templatePath = join(outputDirectory, "05-invoice-template.xlsx");
  const template = createWorkbook("Invoice");
  const templateSheet = template.sheet();

  addReportTitle(templateSheet, "INVOICE", "Fill the highlighted cells and keep the original design.", "D");
  templateSheet.range("A4:D8").setValues([
    ["Item", "Quantity", "Price", "Total"],
    ["", null, null, null],
    ["", null, null, null],
    ["", null, null, null],
    ["Grand total", null, null, null]
  ]);
  addHeader(templateSheet, "A4:D4");
  addGrid(templateSheet, "A4:D8");
  templateSheet.range("A5:C7").style({ fill: palette.amberSoft });
  templateSheet.range("C5:D8").style({ numberFormat: "$#,##0.00" });
  templateSheet.range("A8:D8").style({ bold: true, fill: palette.blueSoft });
  templateSheet.setColumnWidth("A", 24).setColumnWidth("B", 12).setColumnWidth("C", 14).setColumnWidth("D", 15);
  await template.save(templatePath);

  const workbook = await openWorkbook(templatePath);
  const sheet = workbook.sheet("Invoice");
  sheet.range("A5:C7").setValues([
    ["Lunch catering", 12, 18.5],
    ["Coffee service", 12, 3.25],
    ["Delivery", 1, 25]
  ]);
  sheet.cell("D5").formula("B5*C5", 222);
  sheet.cell("D6").formula("B6*C6", 39);
  sheet.cell("D7").formula("B7*C7", 25);
  sheet.cell("D8").formula("SUM(D5:D7)", 286);

  await workbook.save(join(outputDirectory, "05-edit-a-template.xlsx"));
}

async function createChartWorkbook() {
  const workbook = createWorkbook("Sales");
  const sheet = workbook.sheet();

  addReportTitle(sheet, "Monthly revenue", "The chart references worksheet cells, so Excel can refresh it after edits.", "H");
  sheet.range("A4:B9").setValues([
    ["Month", "Revenue"],
    ["January", 18400],
    ["February", 21350],
    ["March", 20100],
    ["April", 24750],
    ["May", 26800]
  ]);
  addHeader(sheet, "A4:B4");
  addGrid(sheet, "A4:B9");
  sheet.range("B5:B9").style({ numberFormat: "$#,##0" });
  sheet.setColumnWidth("A", 15).setColumnWidth("B", 15);

  workbook.charts.add({
    sheet: "Sales",
    name: "RevenueTrend",
    type: "line",
    title: "Revenue trend",
    range: "A4:B9",
    position: { from: "D4", to: "K18" },
    legend: false
  });

  await workbook.save(join(outputDirectory, "06-create-a-chart.xlsx"));
}

async function createPivotWorkbook() {
  const workbook = createWorkbook("Orders");
  const source = workbook.sheet();
  const summary = workbook.addSheet("Summary");

  source.setData([
    { region: "North", category: "Food", sales: 420 },
    { region: "North", category: "Drinks", sales: 180 },
    { region: "South", category: "Food", sales: 510 },
    { region: "South", category: "Drinks", sales: 220 },
    { region: "West", category: "Food", sales: 390 },
    { region: "West", category: "Drinks", sales: 260 }
  ]);
  addHeader(source, "A1:C1");
  addGrid(source, "A1:C7");
  source.range("C2:C7").style({ numberFormat: "$#,##0" });
  source.autoFit({ min: 12, max: 22, padding: 3 });

  addReportTitle(summary, "Sales by region", "A native PivotTable with cached source records and refresh-on-open support.", "F");
  workbook.pivotTables.add({
    name: "SalesByRegion",
    source: { sheet: "Orders", range: "A1:C7" },
    target: { sheet: "Summary", cell: "A4" },
    rows: ["region"],
    columns: ["category"],
    values: [{ field: "sales", summarize: "sum", name: "Total sales" }]
  });
  summary.cell("B4").set("Drinks");
  summary.cell("C4").set("Food");
  summary.cell("A8").set("Total");
  summary.setColumnWidth("A", 24).setColumnWidth("B", 16).setColumnWidth("C", 16).setColumnWidth("D", 20);

  await workbook.save(join(outputDirectory, "07-create-a-pivot-table.xlsx"));

  // Native PivotTables are recalculated by Excel. Keep a small cached preview
  // workbook for the static documentation image renderer.
  const preview = createWorkbook("Summary");
  const previewSheet = preview.sheet();
  addReportTitle(previewSheet, "Sales by region", "The same cached summary Excel displays when the PivotTable opens.", "D");
  previewSheet.range("A4:D8").setValues([
    ["Region", "Drinks", "Food", "Grand total"],
    ["North", 180, 420, 600],
    ["South", 220, 510, 730],
    ["West", 260, 390, 650],
    ["Grand total", 660, 1320, 1980]
  ]);
  addHeader(previewSheet, "A4:D4");
  addGrid(previewSheet, "A4:D8");
  previewSheet.range("B5:D8").style({ numberFormat: "$#,##0" });
  previewSheet.range("A8:D8").style({ bold: true, fill: palette.blueSoft });
  previewSheet.setColumnWidth("A", 24).setColumnWidth("B", 16).setColumnWidth("C", 16).setColumnWidth("D", 20);
  await preview.save(join(previewDirectory, "07-pivot-preview.xlsx"));
}

async function createSecurityWorkbook() {
  const workbook = createWorkbook("Private Report");
  const sheet = workbook.sheet();

  addReportTitle(sheet, "Private sales report", "Yellow cells are editable. Blue cells are protected formulas.", "D");
  sheet.range("A4:D7").setValues([
    ["Item", "Quantity", "Price", "Total"],
    ["Lunch", 2, 12, null],
    ["Coffee", 3, 3.5, null],
    ["Grand total", null, null, null]
  ]);
  addHeader(sheet, "A4:D4");
  addGrid(sheet, "A4:D7");
  sheet.range("A5:C6").style({ fill: palette.amberSoft, protection: { locked: false } });
  sheet.cell("D5").formula("B5*C5", 24).style({ fill: palette.blueSoft, numberFormat: "$#,##0.00" });
  sheet.cell("D6").formula("B6*C6", 10.5).style({ fill: palette.blueSoft, numberFormat: "$#,##0.00" });
  sheet.cell("D7").formula("SUM(D5:D6)", 34.5).style({ bold: true, fill: palette.blueSoft, numberFormat: "$#,##0.00" });
  sheet.range("A7:C7").style({ bold: true, fill: palette.blueSoft });
  sheet.protectSheet({ password: "demo", selectUnlockedCells: true, formatCells: false });
  workbook.protectStructure({ password: "demo", structure: true });
  sheet.setColumnWidth("A", 22).setColumnWidth("B", 14).setColumnWidth("C", 14).setColumnWidth("D", 16);

  await writeFile(join(previewDirectory, "08-security-preview.xlsx"), workbook.toBuffer());
  await workbook.save(join(outputDirectory, "08-protection-and-encryption.xlsx"), { password: "demo" });
}

await createBasicWorkbook();
await createStyledWorkbook();
await createFormulaWorkbook();
await createLayoutWorkbook();
await createTemplateWorkbook();
await createChartWorkbook();
await createPivotWorkbook();
await createSecurityWorkbook();

console.log(`Created tutorial workbooks in ${outputDirectory}`);
