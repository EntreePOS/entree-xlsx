# @entree_pos/xlsx

Create polished Excel workbooks with a JavaScript API that stays close to how
developers think about spreadsheets: workbooks, sheets, cells, ranges, styles,
charts, and pivots.

`@entree_pos/xlsx` is a modern ESM library for Node.js 18+. It can create new
`.xlsx` files or safely update existing `.xlsx` and `.xlsm` templates while
preserving workbook parts it does not modify. There are no runtime dependencies.

Current version: **0.4.0**

```bash
npm install @entree_pos/xlsx
```

New to the library? Follow the [visual tutorial](https://huangxuewu.github.io/entree-xlsx/)
from a basic data export through styling, formulas, templates, charts,
PivotTables, and password encryption.

```js
import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("Orders");
const sheet = workbook.sheet("Orders");

sheet.setData([
  { order: 1001, customer: "Ada", total: 18.5 },
  { order: 1002, customer: "Linus", total: 24 },
  { order: 1003, customer: "Grace", total: 31.25 }
]);

sheet.range("A1:C1").style({
  bold: true,
  color: "#FFFFFF",
  fill: "#17324D",
  horizontal: "center"
});

sheet.range("C2:C100").style({ numberFormat: "$#,##0.00" });
sheet.autoFilter().autoFit({ max: 40 });

await workbook.save("orders.xlsx");
```

That is a complete workbook: structured data, a styled header, currency
formatting, sensible column widths, and an Excel filter.

## Why use it?

| Capability | What you can do |
| --- | --- |
| Direct worksheet helpers | Select ranges, rows, or columns and search, update, style, and size them with chainable helpers |
| Rich styling | Fonts, colors, fills, gradients, borders, alignment, number formats, and cell protection |
| Reusable design system | Define named styles once, inherit from base styles, and compose them across a workbook |
| Template editing | Open an existing workbook, change only what you need, and preserve unknown OOXML parts |
| Charts | Create, inspect, update, and remove column, bar, line, pie, and scatter charts |
| PivotTables | Build and update native PivotTables with row, column, filter, and value fields |
| Password encryption | Save and open files using Excel-compatible AES-256 Agile encryption |
| Simple deployment | JavaScript-only ESM, Node.js 18+, and zero runtime dependencies |

For the complete method-by-method documentation, see the
[API reference](./docs/API.md).

## Styling is a first-class API

Spreadsheet styling quickly becomes repetitive when every cell carries its own
large configuration object. This library lets you define a small workbook design
system and reuse it by name.

```js
workbook.styles
  .define("base", {
    fontName: "Aptos",
    fontSize: 11,
    vertical: "center"
  })
  .define("header", {
    bold: true,
    color: "#FFFFFF",
    fill: "#17324D",
    horizontal: "center",
    wrapText: true
  }, { extends: "base" })
  .define("money", {
    numberFormat: "$#,##0.00;[Red]-$#,##0.00"
  });

sheet.range("A1:D1").style("header");
sheet.range("D2:D100").style("money");

// Named styles and one-off adjustments can be composed left to right.
sheet.cell("D2").style(["money", { bold: true }]);
```

Named styles are written as real Excel cell styles and survive save/reopen
cycles. Direct styles can use friendly shortcuts or detailed nested objects:

```js
sheet.cell("B3").style({
  font: {
    name: "Aptos Display",
    size: 14,
    italic: true,
    color: { theme: 4, tint: -0.2 }
  },
  fill: {
    type: "gradient",
    degree: 90,
    stops: [
      { position: 0, color: "#FFF2CC" },
      { position: 1, color: "#D6A85F" }
    ]
  },
  horizontal: "center",
  vertical: "center",
  rotation: -30,
  shrinkToFit: true,
  protection: { locked: false }
});
```

Range borders understand position, so an outline is not repeatedly applied to
every cell:

```js
sheet.range("A1:D20").style({
  border: {
    outline: { style: "medium", color: "#17324D" },
    inside: { style: "thin", color: "#B7C9D6" }
  }
});
```

You can also copy or selectively remove formatting without touching a cell's
value:

```js
sheet.cell("D2").copyStyleFrom("C2");
sheet.cell("D2").clearStyle(["fill", "font.bold"]);
sheet.range("A2:D20").clearStyle("border");
```

## Edit an existing workbook

Use an existing spreadsheet as a template instead of rebuilding its layout in
code.

```js
import { openWorkbook } from "@entree_pos/xlsx";

const workbook = await openWorkbook("invoice-template.xlsx");
const invoice = workbook.sheet("Invoice");

invoice.cell("B3").set("Entree POS").style({ bold: true });
invoice.cell("F18").formula("SUM(F5:F17)").numberFormat("$#,##0.00");
invoice.range("A20:F20").merge().style({ fill: "#EFE8D8" });

await workbook.save("invoice-ready.xlsx");
```

`openWorkbook()` accepts a local path, HTTP(S) URL, `Buffer`, `Uint8Array`, or
`ArrayBuffer`. `openWorkbookSync(path)` is available for synchronous file access,
and `parseWorkbook(bytes)` reads in-memory XLSX data.

When an existing `.xlsx` or `.xlsm` file is opened, the original OOXML package is
retained and only changed parts are patched. Unknown extensions, VBA, images,
form controls, and embedded parts are preserved.

## Work with cells, ranges, and sheets

The everyday API is intentionally small and chainable:

```js
const sheet = workbook.sheet("Orders");

sheet.set("A1", "Order Number");
sheet.cell("A2").set(1001);
sheet.cell("B2").formula("SUM(B3:B20)", 125.5);
sheet.cell("C2").hyperlink("https://example.com/orders/1001", "Open order");

sheet.range("A1:C3").setValues([
  ["Item", "Qty", "Price"],
  ["Burger", 2, 12.5],
  ["Fries", 1, 4]
]);

sheet.appendRows([["Drink", 2, 3.5]]);
sheet.appendData([{ item: "Shake", qty: 1, price: 5.25 }]);
sheet.insertRows(4, 2);
sheet.copyRow(2, 4);
sheet.deleteRows(10, 1);

sheet.column("B").width(24);
sheet.row(1).height(28);
sheet.autoFit({ min: 8, max: 48, padding: 2 });
```

Select exactly the worksheet area you need. Every selection supports `find()`,
`findAll()`, and `forEach()`:

```js
const table = sheet.range("A1:C20");
const header = sheet.row(1);
const body = sheet.rows("2:20");
const prices = sheet.column("C");
const reportColumns = sheet.columns("A:C");

const burger = table.find("Classic Burger");
console.log(burger?.address); // A2

header.style({ bold: true }).height(24);
body.style({ vertical: "center" });
prices.style({ numberFormat: "$#,##0.00" }).width(14);
reportColumns.forEach((cell) => console.log(cell.address, cell.value));
```

Rows use one-based Excel row numbers. Numeric column references remain
zero-based; column letters are usually easier to read.

Read worksheet data in the form your application needs:

```js
const rows = sheet.toRows();
const records = sheet.toRecords();
const csv = sheet.toCsv();
const html = sheet.toHtml();
```

Manage workbook sheets with similarly direct methods:

```js
workbook.addSheet("Summary", [["Status", "Total"]]);
workbook.renameSheet("Summary", "Dashboard");
workbook.removeSheet("Dashboard");

console.log(workbook.sheetNames);
console.log(workbook.sheetCount);
```

Use `setData()` when the new dataset should replace the current cell values.
Use `appendData()` for object records that should follow the sheet's existing
headers, and `appendRows()` for positional array rows.

## Add a chart

Charts use worksheet ranges as their data source and cell addresses for
placement.

```js
const chart = workbook.charts.add({
  sheet: "Orders",
  name: "RevenueChart",
  type: "column", // column, bar, line, pie, or scatter
  title: "Monthly revenue",
  range: "A1:B13",
  position: { from: "D2", to: "L18" }
});

workbook.charts.update(chart.id, {
  type: "line",
  title: "Revenue trend",
  range: "A1:B13"
});

console.log(workbook.charts.list("Orders"));
// workbook.charts.remove(chart.id);
```

## Build a PivotTable

PivotTables are native Excel objects, not a static table made to look like a
pivot.

```js
workbook.addSheet("Summary");

const pivot = workbook.pivotTables.add({
  name: "SalesPivot",
  source: { sheet: "Orders", range: "A1:D500" },
  target: { sheet: "Summary", cell: "A3" },
  rows: ["Region"],
  columns: ["Month"],
  filters: ["Store"],
  values: [
    { field: "Sales", summarize: "sum", name: "Total Sales" }
  ]
});

workbook.pivotTables.update(pivot.id, {
  rows: ["Store"],
  columns: ["Region"]
});
```

Supported summaries are `sum`, `count`, `average`, `min`, and `max`. The library
writes the PivotTable definition, cache definition, cache records, and a cached
worksheet rendering.

## Protect or encrypt a workbook

Password encryption protects the entire file:

```js
await workbook.save("private.xlsx", {
  password: process.env.REPORT_PASSWORD
});

const opened = await openWorkbook("private.xlsx", {
  password: process.env.REPORT_PASSWORD
});
```

Encrypted files use Excel-compatible AES-256 Agile Office encryption.

Workbook and worksheet protection control editing and structure. They are
separate from file encryption:

```js
workbook.protectStructure({
  password: process.env.STRUCTURE_PASSWORD,
  structure: true
});

workbook.sheet("Orders").protectSheet({
  password: process.env.SHEET_PASSWORD
});
```

## Save or export

```js
workbook.properties = {
  title: "Daily orders",
  author: "Entree POS"
};

await workbook.save("daily-orders.xlsx");

const bytes = workbook.toBuffer();
const base64 = workbook.toBase64();
```

## Command line

The package also includes a small command-line utility for inspection and data
conversion:

```bash
entree-xlsx inspect orders.xlsx
entree-xlsx convert orders.xlsx orders.json --sheet Orders
entree-xlsx convert orders.xlsx orders.csv --sheet Orders
```

## Important behavior

- The library reads and writes `.xlsx` and `.xlsm`, not legacy `.xls` or `.xlsb`
  files.
- Formulas are stored but not calculated by the library. Excel recalculates them
  when the workbook opens.
- The optional second argument to `formula()` is a cached value that spreadsheet
  applications can display before recalculation.
- Workbook and sheet protection discourage editing; use file encryption when the
  workbook contents must remain confidential.

## Requirements

- Node.js 18 or newer
- Native ECMAScript modules (`import` / `export`)
- No runtime dependencies

## License

MIT
