# @entree_pos/xlsx API Reference

Version: `0.3.0`
Runtime: Node.js 18 or newer  
Module format: native ECMAScript modules (ESM)  
Runtime dependencies: none

`@entree_pos/xlsx` creates, reads, edits, and writes `.xlsx` and `.xlsm` workbooks. Its API is
organized around four objects:

- `Workbook` represents an Excel file.
- `Worksheet` represents one sheet in a workbook.
- `Range` represents a rectangular group of cells.
- `Cell` represents one cell.

## Installation

```bash
npm install @entree_pos/xlsx
```

The package is ESM-only:

```js
import {
  createWorkbook,
  openWorkbook,
  openWorkbookSync,
  parseWorkbook
} from "@entree_pos/xlsx";
```

CommonJS `require()` is not supported. A CommonJS application can load the
package with dynamic import:

```js
const { createWorkbook } = await import("@entree_pos/xlsx");
```

## Quick start

```js
import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("Orders");
const orders = workbook.sheet("Orders");

orders.setData([
  { order: 1001, customer: "Ada", total: 18.5 },
  { order: 1002, customer: "Linus", total: 24 }
]);

orders.range("A1:C1").style({
  bold: true,
  color: "#FFFFFF",
  fill: "#17324D"
});
orders.range("C2:C3").style({ numberFormat: "$#,##0.00" });
orders.autoFilter().autoFit();

await workbook.save("orders.xlsx");
```

## Public exports

| Export | Description |
| --- | --- |
| `createWorkbook(name?)` | Creates a workbook containing one worksheet. |
| `openWorkbook(source, options?)` | Opens an XLSX/XLSM path, URL, or binary value; `options.password` opens Agile-encrypted files. |
| `openWorkbookSync(path, options?)` | Synchronously opens a local workbook. |
| `parseWorkbook(data, options?)` | Parses in-memory workbook bytes, including password-encrypted Office files. |
| `version` | The package version string. |
| `Workbook` | Workbook class. |
| `Worksheet` | Worksheet class. |
| `Range` | Range class. |
| `Cell` | Cell class. |
| `ChartCollection` | Chart create/list/update/remove API exposed as `workbook.charts`. |
| `PivotCollection` | PivotTable create/list/update/remove API exposed as `workbook.pivotTables`. |
| `StyleCollection` | Reusable named-style API exposed as `workbook.styles`. |
| `composeStyles(...styles)` | Normalizes and merges style objects from left to right. |
| `encryptWorkbookBuffer` | Applies AES-256 Agile Office password encryption to XLSX bytes. |
| `decryptWorkbookBuffer` | Decrypts Agile-encrypted Office bytes. |
| `isCompoundFile` | Detects the OLE compound container used by encrypted Office files. |
| `normalizeColor` | Converts a friendly color value to the internal color object. |
| `normalizeStyle` | Converts friendly style shortcuts to the internal style object. |
| `XlsxError` | Base library error class. |
| `SheetNotFoundError` | Thrown when a requested worksheet does not exist. |
| `DuplicateSheetError` | Thrown when a worksheet name is already in use. |
| `InvalidSheetNameError` | Thrown when a name violates Excel's sheet-name rules. |
| `InvalidSourceError` | Thrown when XLSX input cannot be loaded or parsed. |

## Accepted values

Cells accept these JavaScript values:

| JavaScript value | Excel representation |
| --- | --- |
| `string` | Text |
| `number` | Number |
| `boolean` | Boolean |
| `Date` | Excel date/time value with a date number format |
| `null` | Blank cell |
| `undefined` | Blank cell |

Excel dates do not store a timezone. A JavaScript `Date` is serialized from its
UTC timestamp and read back as a `Date`.

## Cell and range addresses

String addresses use standard Excel A1 notation:

```js
sheet.cell("A1");
sheet.range("A1:D20");
```

A cell can also be addressed with zero-based coordinates:

```js
sheet.cell({ r: 0, c: 0 }); // A1
sheet.cell({ r: 4, c: 2 }); // C5
```

Rows in A1 strings are one-based. Coordinate objects use zero-based `r` and `c`
values.

---

# Entry-point functions

## `createWorkbook(firstSheetName?)`

Creates a new workbook with one empty worksheet.

Parameters:

| Parameter | Type | Default | Description |
| --- | --- | --- | --- |
| `firstSheetName` | `string` | `"Sheet1"` | Name of the initial worksheet. |

Returns: `Workbook`

```js
const workbook = createWorkbook();
const namedWorkbook = createWorkbook("Orders");
```

The initial name follows Excel's worksheet-name rules: it cannot be blank,
cannot exceed 31 characters, cannot start or end with an apostrophe, and cannot
contain `\ / ? * : [ ]`.

## `openWorkbook(source, options?)`

Asynchronously loads an XLSX or XLSM workbook. Pass `{ password }` for an
Agile-encrypted Office file.

Accepted sources:

- Local file path as a string
- HTTP or HTTPS URL as a string
- `URL` object
- `Buffer`
- `Uint8Array`
- `ArrayBuffer`

Returns: `Promise<Workbook>`

```js
const fromFile = await openWorkbook("reports/orders.xlsx");
const fromUrl = await openWorkbook("https://example.com/orders.xlsx");
const fromBuffer = await openWorkbook(buffer);
const encrypted = await openWorkbook("private.xlsx", { password: process.env.REPORT_PASSWORD });
```

Network requests use the Node.js 18 global `fetch()` implementation. A failed
download or invalid workbook throws `InvalidSourceError`.

## `openWorkbookSync(path, options?)`

Synchronously loads a local XLSX file.

Parameters:

| Parameter | Type | Description |
| --- | --- | --- |
| `path` | `string` | Local `.xlsx` or `.xlsm` file path. |
| `options.password` | `string` | Password for an encrypted Office file. |

Returns: `Workbook`

```js
const workbook = openWorkbookSync("orders.xlsx");
```

Prefer `openWorkbook()` in servers and applications where blocking the event
loop is undesirable.

## `parseWorkbook(data, options?)`

Parses XLSX data already loaded in memory.

Accepted values: `Buffer`, `Uint8Array`, or `ArrayBuffer`.

Returns: `Workbook`

```js
const response = await fetch("https://example.com/orders.xlsx");
const workbook = parseWorkbook(await response.arrayBuffer());
const privateWorkbook = parseWorkbook(encryptedBytes, { password: process.env.REPORT_PASSWORD });
```

---

# `Workbook`

## `workbook.sheetNames`

Read-only getter returning a new array of worksheet names in workbook order.
Changing the returned array does not change the workbook.

```js
console.log(workbook.sheetNames); // ["Orders", "Summary"]
```

## `workbook.sheetCount`

Read-only getter returning the number of worksheets.

```js
console.log(workbook.sheetCount); // 2
```

## `workbook.properties`

Gets or merges document properties.

Recognized properties include:

| Property | Value | Description |
| --- | --- | --- |
| `title` | `string` | Workbook title. |
| `author` | `string` | Workbook creator. |
| `lastAuthor` | `string` | Last editor name. |
| `createdAt` | `Date` or ISO string | Creation timestamp. |

```js
workbook.properties = {
  title: "August orders",
  author: "Entree POS"
};

console.log(workbook.properties.title);
```

Assigning properties merges them with existing properties.

## `workbook.sheet(reference?)`

Returns a worksheet by name or zero-based index.

| Parameter | Type | Default | Description |
| --- | --- | --- | --- |
| `reference` | `string` or `number` | `0` | Sheet name or zero-based index. |

Returns: `Worksheet`

Throws: `SheetNotFoundError`

```js
const first = workbook.sheet();
const byIndex = workbook.sheet(1);
const orders = workbook.sheet("Orders");
```

## `workbook.findSheet(reference?)`

Returns a worksheet when found, otherwise returns `undefined`. Other unexpected
errors are still thrown.

```js
const optional = workbook.findSheet("Optional");
if (optional) console.log(optional.usedRange);
```

## `workbook.addSheet(name, data?)`

Adds a worksheet and optionally fills it with array rows or object rows.

Returns: the new `Worksheet`

Throws: `DuplicateSheetError` or `InvalidSheetNameError`

```js
const notes = workbook.addSheet("Notes");

const data = workbook.addSheet("Data", [
  ["Item", "Quantity"],
  ["Burger", 2]
]);

const people = workbook.addSheet("People", [
  { name: "Ada", active: true },
  { name: "Linus", active: false }
]);
```

Object data automatically creates a header row from its property names.

## `workbook.removeSheet(reference)`

Removes a sheet by name or zero-based index.

Returns: the same `Workbook` for chaining.

The final remaining worksheet cannot be removed.

```js
workbook.removeSheet("Temporary");
workbook.removeSheet(1);
```

## `workbook.renameSheet(reference, newName)`

Renames a sheet and returns the renamed `Worksheet`.

```js
const renamed = workbook.renameSheet("Sheet1", "Orders");
console.log(renamed.name); // Orders
```

This operation changes the worksheet name. It does not rewrite formula text that
contains the old worksheet name.

## `workbook.toBuffer(options?)`

Serializes the workbook to a Node.js `Buffer` containing an XLSX file.

```js
const buffer = workbook.toBuffer();
const encrypted = workbook.toBuffer({ password: process.env.REPORT_PASSWORD });
```

## `workbook.toUint8Array(options?)`

Serializes the workbook to `Uint8Array`.

```js
const bytes = workbook.toUint8Array();
```

## `workbook.toBase64(options?)`

Serializes the workbook to a base64 string.

```js
const base64 = workbook.toBase64();
```

## `await workbook.save(path, options?)`

Asynchronously writes the workbook to disk and returns the same `Workbook`.
The path must end in `.xlsx` or `.xlsm`. Pass `options.password` for real
AES-256 Agile Office encryption.

```js
await workbook.save("output/orders.xlsx");
await workbook.save("output/private.xlsx", { password: process.env.REPORT_PASSWORD });
```

## `workbook.saveSync(path, options?)`

Synchronously writes the workbook to disk and returns the same `Workbook`.
The path must end in `.xlsx` or `.xlsm`.

```js
workbook.saveSync("output/orders.xlsx");
```

## `workbook.toJSON(options?)`

Returns an object whose keys are sheet names and values are the result of each
sheet's `toRecords(options)` call.

```js
const data = workbook.toJSON();

// {
//   Orders: [{ order: 1001, total: 18.5 }],
//   Customers: [{ name: "Ada" }]
// }
```

## `workbook.styles`

Returns the workbook's `StyleCollection`. Use it to define reusable styles,
inherit from base styles, inspect styles read from an existing workbook, and
apply a style by name.

```js
workbook.styles
  .define("base", { fontName: "Aptos", fontSize: 11 })
  .define("header", {
    bold: true,
    color: "#FFFFFF",
    fill: "#17324D"
  }, { extends: "base" });

workbook.sheet().range("A1:D1").style("header");
```

Named styles are written to `cellStyles` and `cellStyleXfs` in the XLSX style
part, so they appear as real named cell styles in Excel and survive reopening.

## `workbook.charts`

Returns the workbook's `ChartCollection`.

```js
const chart = workbook.charts.add({
  sheet: "Sales",
  name: "RevenueChart",
  type: "column",
  title: "Monthly revenue",
  range: "A1:C13",
  position: { from: "E2", to: "M18" }
});

workbook.charts.list("Sales");
workbook.charts.update(chart.id, {
  type: "line",
  title: "Revenue trend",
  range: "A1:C13"
});
workbook.charts.remove(chart.id);
```

Supported chart types are `column`, `bar`, `line`, `pie`, and `scatter`.
Instead of `range`, advanced callers can pass `series` entries containing
`name`, `nameCell`, `categories`, `xValues`, and `values` ranges. Chart updates
require `range` or `series` so the data mapping remains explicit.

## `workbook.pivotTables`

Returns the workbook's `PivotCollection`.

```js
const pivot = workbook.pivotTables.add({
  name: "SalesPivot",
  source: { sheet: "Orders", range: "A1:D500" },
  target: { sheet: "Summary", cell: "A3" },
  rows: ["Region"],
  columns: ["Month"],
  filters: ["Store"],
  values: [{ field: "Sales", summarize: "sum", name: "Total Sales" }],
  refreshOnLoad: true,
  style: "PivotStyleMedium9"
});

workbook.pivotTables.list("Summary");
workbook.pivotTables.update(pivot.id, { rows: ["Store"], columns: ["Region"] });
workbook.pivotTables.remove(pivot.id);
```

Summaries can use `sum`, `count`, `average`, `min`, or `max`. Creation writes a
PivotTable definition, cache definition, cache records, workbook and worksheet
relationships, and a cached summary in the target cells.

## `workbook.protectStructure(options?)` and `workbook.unprotectStructure()`

Protects workbook structure from ordinary Excel UI edits.

```js
workbook.protectStructure({ password: "structure-password", structure: true });
workbook.unprotectStructure();
```

This is an editing control, not encryption. Use a save password to protect the
file contents.

---

# `StyleCollection`

`workbook.styles` owns reusable styles for the workbook. A named style can be
applied anywhere a style input is accepted.

## `styles.define(name, style, options?)`

Defines or replaces a named style and returns the same collection for chaining.

| Parameter | Type | Description |
| --- | --- | --- |
| `name` | non-empty string | Case-sensitive style name. |
| `style` | style object | Any style described in the styling reference. |
| `options.extends` | string or string array | Parent style or styles inherited from left to right. |

```js
workbook.styles
  .define("reportBase", {
    font: { name: "Aptos", size: 11 },
    vertical: "center"
  })
  .define("reportHeader", {
    bold: true,
    fill: "#17324D",
    color: "#FFFFFF",
    horizontal: "center"
  }, { extends: "reportBase" });
```

Child properties override inherited properties. Missing parents and circular
inheritance throw immediately, and the invalid definition is not retained.
Replacing a definition affects future applications of the style; it does not
automatically restyle cells that were already formatted.

## `styles.defineMany(definitions)`

Defines several styles atomically and returns the collection. This form permits
forward references because all definitions are installed before inheritance is
validated.

```js
workbook.styles.defineMany({
  base: { fontName: "Aptos", fontSize: 11 },
  header: {
    extends: "base",
    style: { bold: true, fill: "#17324D", color: "#FFFFFF" }
  },
  money: { numberFormat: "$#,##0.00;[Red]-$#,##0.00" }
});
```

Use the `{ extends, style }` wrapper only when inheritance is needed. A plain
value is treated as the style object.

## `styles.names`

Read-only array of style names.

```js
console.log(workbook.styles.names);
```

## `styles.has(name)`

Returns `true` when a style exists.

## `styles.get(name)`

Returns a deep copy of the fully resolved style, including inherited values.

```js
const resolvedHeader = workbook.styles.get("reportHeader");
```

Changing the returned object does not change the definition.

## `styles.getDefinition(name)`

Returns the stored definition without resolving inheritance:

```js
{
  name: "reportHeader",
  extends: ["reportBase"],
  style: { font: { bold: true }, fill: { /* ... */ } }
}
```

## `styles.list()`

Returns every stored definition in workbook order.

## `styles.resolve(input)`

Resolves a style object, a style name, or an array containing both. Array items
are merged from left to right.

```js
const style = workbook.styles.resolve([
  "reportBase",
  "money",
  { bold: true }
]);
```

Most applications call `cell.style()` or `range.style()` directly instead.

## `styles.remove(name)`

Removes a named style and returns `true`, or returns `false` when it did not
exist. A style cannot be removed while another style extends it. Removal is
persisted for both new workbooks and opened templates.

---

# `Worksheet`

## `worksheet.name`

Read-only worksheet name.

## `worksheet.usedRange`

Read-only A1 range covering every populated cell, such as `"A1:D25"`. It is
`undefined` for a completely empty sheet.

## `worksheet.unsafeRaw`

Returns the internal worksheet object. This is an advanced escape hatch and is
not a stable interchange format. Prefer the documented methods.

## `worksheet.cell(address)`

Returns a `Cell` object.

```js
const cell = worksheet.cell("B2");
const sameCell = worksheet.cell({ r: 1, c: 1 });
```

## `worksheet.get(address)`

Returns a cell's JavaScript value, or `undefined` when the cell is empty.

```js
const total = worksheet.get("C2");
```

## `worksheet.set(address, value)`

Sets a value and returns the same `Worksheet` for chaining.

```js
worksheet
  .set("A1", "Order")
  .set("B1", "Total");
```

## `worksheet.range(address)`

Returns a `Range` object.

```js
const header = worksheet.range("A1:D1");
```

## `worksheet.column(column)`

Returns a `Column` helper. A column can be selected by letter or by a zero-based
index.

```js
const items = worksheet.column("A");
const prices = worksheet.column(1); // column B
```

## `worksheet.find(valueOrPredicate)` and `worksheet.findAll(valueOrPredicate)`

Search populated cells in row order. `find()` returns the first matching
`Cell`, or `undefined`. `findAll()` returns every matching `Cell`.

```js
const item = worksheet.find("Classic Burger");
console.log(item?.address);

const expensive = worksheet.findAll((cell) =>
  typeof cell.value === "number" && cell.value > 100
);
```

## `worksheet.appendRows(data, options?)`

Appends or inserts array rows or object rows.

Options:

| Option | Type | Default | Description |
| --- | --- | --- | --- |
| `origin` | A1 string, zero-based row number, `{r,c}`, or `-1` | `-1` | Starting cell. `-1` appends below the used range. A numeric origin is a zero-based row with column 0. |
| `header` | `string[]` | inferred | Property order for object rows. |
| `skipHeader` | `boolean` | `false` | Prevents automatic headers for object rows. |

Array rows:

```js
worksheet.appendRows([
  [1001, "Ada", 18.5],
  [1002, "Linus", 24]
]);
```

Object rows:

```js
worksheet.appendRows([
  { order: 1001, customer: "Ada", total: 18.5 },
  { order: 1002, customer: "Linus", total: 24 }
], {
  header: ["order", "customer", "total"],
  skipHeader: false
});
```

Append object rows to a sheet that already has headers:

```js
worksheet.appendRows([
  { order: 1003, customer: "Grace", total: 31 }
], {
  header: ["order", "customer", "total"],
  skipHeader: true
});
```

For record objects, prefer `appendData()`. It reuses the current header row so
you do not need to repeat the header order.

## `worksheet.appendData(records, options?)`

Appends object records below existing data. When the worksheet is empty, it
creates a header row. When data already exists, it reads the first used row as
the header order and maps each record to those columns.

```js
worksheet.setData([
  { order: 1001, customer: "Ada", total: 18.5 }
]);

worksheet.appendData([
  { customer: "Linus", total: 24, order: 1002 }
]);
```

Use `options.header` to provide an explicit property order. Use
`options.origin` only when the records should be written somewhere other than
the next row after the used range.

## `worksheet.insertRows(beforeRow, count?, options?)`

Inserts rows before a one-based row number and shifts cells, row properties,
merges, and straightforward A1 formula references.

```js
worksheet.insertRows(5, 2);
worksheet.insertRows(10, 1, { copyFrom: "above" });
```

`options.copyFrom` can be `"above"`, `"below"`, or a one-based source row.
Set `options.values` to `false` to copy formatting without values.

## `worksheet.deleteRows(startRow, count?)`

Deletes rows and shifts following rows upward.

```js
worksheet.deleteRows(20, 3);
```

## `worksheet.copyRow(sourceRow, targetRow, options?)`

Copies a row using one-based row numbers. Relative formula row references are
adjusted for the target row.

```js
worksheet.copyRow(5, 12);
worksheet.copyRow(5, 13, { values: false });
```

## `worksheet.setData(data, options?)`

Removes existing cells, resets the used range, and writes new data starting at
`A1`. Existing layout metadata such as merges, widths, heights, and filters is
not automatically cleared.

```js
worksheet.setData([
  { item: "Burger", quantity: 2 },
  { item: "Fries", quantity: 1 }
]);
```

## `worksheet.toRows(options?)`

Returns the selected region as an array of arrays.

Options:

| Option | Type | Description |
| --- | --- | --- |
| `range` | A1 range string | Limits the output region. Defaults to `usedRange`. |
| `defaultValue` | any | Replaces `undefined` empty values. |

```js
const rows = worksheet.toRows({ range: "A1:C10", defaultValue: null });
```

## `worksheet.toRecords(options?)`

Returns rows as objects.

By default, the first row supplies property names and is excluded from the
result. Provide `headers` to supply property names and include every selected
row as data.

Options:

| Option | Type | Description |
| --- | --- | --- |
| `headers` | `string[]` | Explicit property names. |
| `range` | A1 range string | Limits the source region. |
| `defaultValue` | any | Value used for missing cells. Defaults to `null`. |

```js
const records = worksheet.toRecords();

const recordsWithoutAHeaderRow = worksheet.toRecords({
  headers: ["order", "customer", "total"],
  range: "A1:C20"
});
```

## `worksheet.toCsv(options?)`

Returns CSV or another delimiter-separated text format.

Options:

| Option | Default | Description |
| --- | --- | --- |
| `delimiter` | `","` | Field separator. Use `"\t"` for TSV. |
| `newline` | `"\n"` | Row separator. |
| `range` | used range | Optional A1 output range. |
| `defaultValue` | `undefined` | Replacement for empty cells. |

```js
const csv = worksheet.toCsv();
const tsv = worksheet.toCsv({ delimiter: "\t", newline: "\r\n" });
```

Text containing the delimiter, quotes, or line breaks is automatically quoted.

## `worksheet.toHtml(options?)`

Returns an HTML `<table>` string. Values are XML/HTML escaped.

Options include the `toRows()` options plus:

| Option | Default | Description |
| --- | --- | --- |
| `header` | `true` | When not `false`, emits `<th>` elements for the first row. |

```js
const table = worksheet.toHtml({ header: true });
```

## `worksheet.merge(range)`

Merges a range and returns the worksheet.

```js
worksheet.merge("A1:D1").set("A1", "Daily Orders");
```

Only the top-left cell's value is normally displayed by Excel.

## `worksheet.unmerge(range)`

Removes the exact matching merged range and returns the worksheet.

```js
worksheet.unmerge("A1:D1");
```

## `worksheet.setColumnWidth(column, width)`

Sets an Excel character-based column width.

`column` can be a letter or zero-based number.

```js
worksheet.setColumnWidth("A", 14);
worksheet.setColumnWidth(1, 28); // column B
```

## `worksheet.setRowHeight(row, height)`

Sets a row height in points. The row number is one-based.

```js
worksheet.setRowHeight(1, 28);
```

## `worksheet.autoFilter(range?)`

Applies an Excel auto-filter. The current used range is used when no range is
provided.

```js
worksheet.autoFilter();
worksheet.autoFilter("A1:D100");
```

## `worksheet.autoFit(options?)`

Calculates approximate column widths from the JavaScript string lengths of cell
values.

Options:

| Option | Default | Description |
| --- | --- | --- |
| `min` | `8` | Minimum width. |
| `max` | `60` | Maximum width. |
| `padding` | `2` | Extra characters added to the measured width. |
| `includeHeader` | `true` | Set to `false` to ignore the first used row. |

```js
worksheet.autoFit({ min: 10, max: 40, padding: 3 });
```

Auto-fit is an estimate; Node.js does not have Excel's font-rendering engine.

## `worksheet.protectSheet(options?)` and `worksheet.unprotectSheet()`

Adds or removes Excel worksheet editing protection.

```js
worksheet.protectSheet({
  password: "sheet-password",
  objects: true,
  scenarios: true
});

worksheet.unprotectSheet();
```

Worksheet protection is not file encryption. Use `save(path, { password })`
when the workbook contents must be encrypted.

---

# `Column`

## `column.index` and `column.letter`

Read-only zero-based index and normalized Excel column letter.

## `column.find(valueOrPredicate)` and `column.findAll(valueOrPredicate)`

Search populated cells in the column and return `Cell` objects.

```js
const cell = worksheet.column("A").find("Classic Burger");
console.log(cell?.address); // A2
cell?.set("Deluxe Burger");
```

## `column.forEach(callback)`

Calls the callback for each populated cell in row order. Empty cells are not
materialized or visited.

## `column.style(style, mode?)`

Styles the entire Excel column without creating a cell for every possible row.
The style is also applied to existing cells in the column. As with cell and
range styles, `mode` can be `"merge"` or `"replace"`.

```js
worksheet.column("B").style({ numberFormat: "$#,##0.00" });
```

## `column.width(width)`

Sets the Excel character-based width and returns the same column helper.

```js
worksheet.column("A").width(24).style({ bold: true });
```

---

# `Cell`

## `cell.address`

Normalized A1 address, such as `"B3"`.

## `cell.coordinates`

Zero-based coordinate object, such as `{ r: 2, c: 1 }` for `B3`.

## `cell.value`

Gets or sets the JavaScript value directly.

```js
console.log(cell.value);
cell.value = 42;
```

## `cell.unsafeRaw`

Returns the internal cell record, or `undefined` when the cell has not been
created. This is an advanced escape hatch and not a stable interchange format.

## `cell.set(value)`

Sets the value and returns the same cell for chaining.

```js
worksheet.cell("A1").set("Orders").style({ bold: true });
```

## `cell.formula(formula, cachedResult?)`

Sets an Excel formula. A leading `=` is optional and is removed automatically.

```js
worksheet.cell("D2").formula("B2*C2", 25);
worksheet.cell("D3").formula("=SUM(D4:D20)", 125.5);
```

The library writes formulas but does not calculate them. `cachedResult` is the
value displayed by readers before Excel recalculates the workbook. It can be a
string, number, boolean, or date.

## `cell.style(style, mode?)`

Applies a style and returns the cell.

| Parameter | Default | Description |
| --- | --- | --- |
| `style` | required | Style object, named style, or array of objects and names. |
| `mode` | `"merge"` | `"merge"` keeps unspecified existing style properties; `"replace"` replaces the complete style. |

```js
worksheet.cell("A1").style({ bold: true, fill: "#17324D" });
worksheet.cell("A1").style({ italic: true }, "merge");
worksheet.cell("B2").style("money");
worksheet.cell("C2").style(["money", { bold: true }]);
```

Array styles are composed from left to right. When a named style is used, the
cell retains the Excel named-style association in addition to the resolved
formatting.

## `cell.getStyle()`

Returns a deep copy of the cell's direct style object. Editing the returned
object does not edit the cell.

```js
const style = worksheet.cell("A1").getStyle();
```

## `cell.copyStyleFrom(source, mode?)`

Copies formatting from another `Cell` or an address in the same worksheet.
The default mode is `"replace"`; pass `"merge"` to preserve unrelated target
formatting.

```js
worksheet.cell("D2").copyStyleFrom("C2");
worksheet.cell("D3").copyStyleFrom(anotherSheet.cell("C3"), "merge");
```

Values and formulas are not copied.

## `cell.clearStyle(parts?)`

With no argument, removes all direct formatting and the named-style
association. Pass one path or an array of paths to remove only selected parts.
Friendly aliases such as `bold`, `fontSize`, `color`, `horizontal`, and
`wrapText` are accepted.

```js
worksheet.cell("A1").clearStyle();
worksheet.cell("B2").clearStyle("fill");
worksheet.cell("C2").clearStyle(["font.bold", "alignment.wrapText"]);
```

## `cell.numberFormat(format)`

Sets an Excel number-format string and returns the cell.

```js
worksheet.cell("B2").numberFormat("$#,##0.00");
worksheet.cell("C2").numberFormat("0.00%");
worksheet.cell("D2").numberFormat("yyyy-mm-dd");
```

## `cell.hyperlink(target, tooltip?)`

Adds an external or internal hyperlink and returns the cell.

```js
worksheet.cell("A2").hyperlink(
  "https://example.com/orders/1001",
  "Open order"
);

worksheet.cell("B2").hyperlink("#Summary!A1", "Go to summary");
```

## `cell.clear(options?)`

Clears a cell and returns it.

| Option | Default | Description |
| --- | --- | --- |
| `keepStyle` | `false` | Retains formatting on a blank cell. |

```js
worksheet.cell("A2").clear();
worksheet.cell("B2").clear({ keepStyle: true });
```

---

# `Range`

## `range.address`

Normalized A1 range address.

```js
console.log(worksheet.range("a1:c3").address); // A1:C3
```

## `range.coordinates`

Zero-based range coordinates:

```js
{
  s: { r: 0, c: 0 },
  e: { r: 2, c: 2 }
}
```

## `range.getValues()`

Returns a rectangular array of arrays. Empty cells are `undefined`.

```js
const values = worksheet.range("A1:C3").getValues();
```

## `range.find(valueOrPredicate)` and `range.findAll(valueOrPredicate)`

Search populated cells inside the range and return `Cell` objects.

```js
const cell = worksheet.range("A2:A100").find("Classic Burger");
```

## `range.setValues(rows)`

Writes a rectangular array and returns the range. The supplied rows cannot be
larger than the selected range.

```js
worksheet.range("A1:C3").setValues([
  ["Item", "Qty", "Price"],
  ["Burger", 2, 12.5],
  ["Fries", 1, 4]
]);
```

## `range.style(style, mode?)`

Applies a style object, named style, or style array to every cell and returns
the range. Empty cells inside the range are created so they can carry
formatting. Range-aware `outline`, `inside`, `insideHorizontal`, and
`insideVertical` borders are positioned for each cell automatically.

```js
worksheet.range("A1:C1").style({
  bold: true,
  fill: "#17324D",
  color: "#FFFFFF"
});

worksheet.range("A2:C20").style("reportBody");
worksheet.range("A1:C20").style({ border: {
  outline: { style: "medium", color: "#17324D" },
  inside: { style: "thin", color: "#B7C9D6" }
} });
```

## `range.copyStyleFrom(source, options?)`

Copies only styles from another `Range` or an address in the same worksheet.
Source and target dimensions must match unless `repeat: true` is supplied,
which tiles the source styles across the target.

```js
worksheet.range("A10:D12").copyStyleFrom("A1:D3");
worksheet.range("F2:F100").copyStyleFrom("E2", { repeat: true });
worksheet.range("A1:D3").copyStyleFrom(otherSheet.range("A1:D3"));
```

Options:

| Option | Default | Description |
| --- | --- | --- |
| `mode` | `"replace"` | `"replace"` or `"merge"`. |
| `repeat` | `false` | Tiles source styles when dimensions differ. |

## `range.clearStyle(parts?)`

Calls `cell.clearStyle(parts)` for every cell in the range.

```js
worksheet.range("A1:D20").clearStyle("fill");
```

## `range.clear(options?)`

Clears every cell in the range and returns the range.

```js
worksheet.range("A2:C20").clear({ keepStyle: true });
```

## `range.merge()`

Adds the exact range to the worksheet's merge list and returns the range.
Calling it repeatedly for the same range does not duplicate the merge.

## `range.unmerge()`

Removes the exact matching merged range and returns the range.

## `range.forEach(callback)`

Calls the callback once for every cell, from left to right and top to bottom,
and returns the range.

```js
worksheet.range("A2:A20").forEach((cell) => {
  if (typeof cell.value === "number") {
    cell.set(cell.value * 1.1);
  }
});
```

---

# Styling reference

Styles can be applied as an object, a named style, or an array composed from
left to right. Friendly top-level shortcuts and full nested objects can be mixed.

```js
const style = {
  bold: true,
  italic: false,
  underline: "single",
  strike: false,
  fontName: "Arial",
  fontSize: 12,
  color: "#FFFFFF",
  fill: "#17324D",
  horizontal: "center",
  vertical: "center",
  wrapText: true,
  textRotation: 0,
  numberFormat: "$#,##0.00",
  border: {
    top: { style: "thin", color: "#999999" },
    right: { style: "thin", color: "#999999" },
    bottom: { style: "thin", color: "#999999" },
    left: { style: "thin", color: "#999999" }
  },
  protection: {
    locked: true,
    hidden: false
  }
};

worksheet.range("A1:D1").style(style);
```

## Friendly style shortcuts

| Property | Accepted values | Description |
| --- | --- | --- |
| `bold` | boolean | Bold font. |
| `italic` | boolean | Italic font. |
| `underline` | boolean, `"single"`, `"double"`, `"singleAccounting"`, `"doubleAccounting"` | Underline style. |
| `strike` | boolean | Strikethrough font. |
| `fontName` | string | Font family name. |
| `fontSize` | positive number | Font size in points. |
| `color` | color | Font color. |
| `subscript` | boolean | Enables subscript text. |
| `superscript` | boolean | Enables superscript text. |
| `fill` | color or fill object | Solid, patterned, linear-gradient, or path-gradient background. |
| `horizontal` | `"left"`, `"center"`, `"right"`, `"fill"`, `"justify"`, `"centerContinuous"`, `"distributed"` | Horizontal alignment. |
| `vertical` | `"top"`, `"center"`, `"bottom"`, `"justify"`, `"distributed"` | Vertical alignment. |
| `wrapText` | boolean | Wrap text within the cell. |
| `shrinkToFit` | boolean | Shrinks text to fit the cell. |
| `rotation` | -90 through 90 | Friendly clockwise/counterclockwise text angle. |
| `textRotation` | 0 through 180, or 255 | Raw OOXML text rotation. |
| `verticalText` | boolean | Stacks characters vertically (`textRotation: 255`). |
| `numberFormat` | string | Excel number-format code. |
| `border` | object | Cell sides and range-aware outline/interior definitions. |
| `protection` | object | `locked` and `hidden` flags. |
| `editable` | boolean | Friendly inverse of `protection.locked`. |
| `raw` | object | Advanced internal style fields merged before friendly fields. |

## Font object

Use `font` when the top-level shortcuts are not enough:

```js
{
  font: {
    name: "Aptos Display",
    size: 16,
    bold: true,
    italic: false,
    underline: "doubleAccounting",
    strike: false,
    outline: false,
    shadow: false,
    color: { theme: 4, tint: -0.2 },
    verticalAlign: "superscript", // baseline, subscript, superscript
    family: 2,
    charset: 1,
    scheme: "major" // major, minor, none
  }
}
```

The top-level shortcuts override the equivalent nested property when both are
provided.

## Pattern and gradient fills

A color value is shorthand for a solid fill:

```js
{ fill: "#17324D" }
```

Pattern fills support a foreground and background color:

```js
{
  fill: {
    type: "pattern",
    pattern: "lightTrellis",
    foreground: "#FFF2CC",
    background: "#D6A85F"
  }
}
```

`patternType`, `fgColor`, `bgColor`, and `color` are accepted aliases for
`pattern`, `foreground`, and `background`. Common patterns include `solid`,
`darkGray`, `mediumGray`, `lightGray`, `gray125`, `darkHorizontal`,
`darkVertical`, `darkDown`, `darkUp`, `darkGrid`, `darkTrellis`,
`lightHorizontal`, `lightVertical`, `lightDown`, `lightUp`, `lightGrid`, and
`lightTrellis`.

Linear gradient:

```js
{
  fill: {
    type: "gradient",
    degree: 90,
    stops: [
      { position: 0, color: "#FFFFFF" },
      { position: 0.5, color: "#FFF2CC" },
      { position: 1, color: "#D6A85F" }
    ]
  }
}
```

Each stop position must be between 0 and 1. For a path gradient, use
`gradientType: "path"` with optional `left`, `right`, `top`, and `bottom`
convergence values.

## Alignment object

```js
{
  alignment: {
    horizontal: "center",
    vertical: "center",
    wrapText: true,
    shrinkToFit: false,
    textRotation: 45,
    indent: 1,
    readingOrder: 0,
    justifyLastLine: false
  }
}
```

Use the top-level `rotation` shortcut for the familiar -90 through 90 degree
range shown in Excel's Format Cells dialog.

## Borders

Each border side accepts a style string or `{ style, color }`:

```js
worksheet.cell("A1").style({ border: {
  bottom: { style: "double", color: "#17324D" },
  diagonal: { style: "thin", color: "#D6A85F" },
  diagonalUp: true
} });
```

Common border styles include `thin`, `medium`, `thick`, `dotted`, `dashed`,
`double`, `hair`, `mediumDashed`, `dashDot`, `mediumDashDot`, `dashDotDot`, and
`mediumDashDotDot`.

Ranges additionally understand these semantic sides:

| Border key | Effect |
| --- | --- |
| `all` | Applies the same border to all four sides of every cell. |
| `outline` | Applies only to the outside edges of the range. |
| `inside` | Applies to every internal horizontal and vertical edge. |
| `insideHorizontal` | Applies only between rows. |
| `insideVertical` | Applies only between columns. |

```js
worksheet.range("A1:D20").style({ border: {
  outline: { style: "medium", color: "#17324D" },
  inside: { style: "thin", color: "#B7C9D6" }
} });
```

## Cell protection style

Cell protection becomes active when the worksheet is protected:

```js
worksheet.range("A2:C20").style({ protection: {
  locked: false,
  hidden: false
} });
worksheet.protectSheet({ password: process.env.SHEET_PASSWORD });
```

`editable: true` is shorthand for `protection.locked: false`.

## Colors

The simplest color input is a six- or eight-digit hexadecimal value:

```js
{ color: "#FFFFFF", fill: "#17324D" }
```

Color objects are also accepted:

```js
{ color: { rgb: "FF0000" } }
{ color: { theme: 1, tint: 0.25 } }
{ color: { indexed: 10 } }
{ color: { auto: true } }
```

RGB colors can also be numeric values such as `0xFF0000`. Invalid colors and
style values throw before the workbook is saved.

## Style composition helpers

`composeStyles(...objects)` normalizes and merges plain style objects:

```js
import { composeStyles } from "@entree_pos/xlsx";

const total = composeStyles(
  { fontName: "Aptos", fontSize: 11 },
  { bold: true, numberFormat: "$#,##0.00" }
);
```

To compose named styles, use `workbook.styles.resolve()` or pass an array to
`cell.style()` / `range.style()`.

## Number-format examples

| Format | Example use |
| --- | --- |
| `0` | Integer |
| `0.00` | Two decimals |
| `#,##0` | Thousands separators |
| `$#,##0.00` | Currency |
| `0.00%` | Percentage |
| `yyyy-mm-dd` | Date |
| `yyyy-mm-dd hh:mm:ss` | Date and time |
| `@` | Text |

---

# Errors

All library-specific errors extend `XlsxError` and include a stable `code`.

| Error | Code | Typical cause |
| --- | --- | --- |
| `XlsxError` | `XLSX_ERROR` | Base error. |
| `SheetNotFoundError` | `SHEET_NOT_FOUND` | Invalid sheet name or index. |
| `DuplicateSheetError` | `DUPLICATE_SHEET` | Adding or renaming to an existing name. |
| `InvalidSheetNameError` | `INVALID_SHEET_NAME` | Invalid Excel worksheet name. |
| `InvalidSourceError` | `INVALID_SOURCE` | Missing file, download failure, unsupported binary value, or invalid XLSX data. |

```js
import {
  openWorkbook,
  InvalidSourceError,
  SheetNotFoundError
} from "@entree_pos/xlsx";

try {
  const workbook = await openWorkbook("orders.xlsx");
  const orders = workbook.sheet("Orders");
} catch (error) {
  if (error instanceof InvalidSourceError) {
    console.error("Could not open the workbook", error.cause);
  } else if (error instanceof SheetNotFoundError) {
    console.error(error.message);
  } else {
    throw error;
  }
}
```

Standard JavaScript `TypeError` and `RangeError` are used for invalid method
arguments such as malformed addresses, negative widths, or oversized range data.

---

# Practical examples

## Build a styled sales report

```js
import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("Sales");
workbook.properties = {
  title: "Daily Sales",
  author: "Entree POS"
};

const sales = workbook.sheet("Sales");
sales.setData([
  { item: "Burger", quantity: 12, revenue: 150 },
  { item: "Fries", quantity: 8, revenue: 32 },
  { item: "Drink", quantity: 15, revenue: 37.5 }
]);

sales.range("A1:C1").style({
  bold: true,
  color: "#FFFFFF",
  fill: "#17324D",
  horizontal: "center"
});
sales.range("C2:C4").style({ numberFormat: "$#,##0.00" });
sales.cell("B6").set("Total").style({ bold: true });
sales.cell("C6").formula("SUM(C2:C4)", 219.5)
  .numberFormat("$#,##0.00")
  .style({ bold: true });
sales.autoFilter("A1:C4").autoFit({ max: 30 });

await workbook.save("daily-sales.xlsx");
```

## Modify an existing workbook safely

```js
import { openWorkbook } from "@entree_pos/xlsx";

const workbook = await openWorkbook("orders.xlsx");
const orders = workbook.findSheet("Orders");

if (!orders) {
  throw new Error("The Orders sheet is required");
}

orders.appendRows([
  { order: 1003, customer: "Grace", total: 31 }
], {
  header: ["order", "customer", "total"],
  skipHeader: true
});

await workbook.save("orders-updated.xlsx");
```

## Send an XLSX file from a Node HTTP server

```js
import { createServer } from "node:http";
import { createWorkbook } from "@entree_pos/xlsx";

createServer((request, response) => {
  const workbook = createWorkbook("Orders");
  workbook.sheet().setData([
    { order: 1001, total: 18.5 },
    { order: 1002, total: 24 }
  ]);

  const body = workbook.toBuffer();
  response.writeHead(200, {
    "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "Content-Disposition": 'attachment; filename="orders.xlsx"',
    "Content-Length": body.length
  });
  response.end(body);
}).listen(3000);
```

## Export JSON and CSV

```js
import { openWorkbook } from "@entree_pos/xlsx";

const workbook = await openWorkbook("orders.xlsx");
const orders = workbook.sheet("Orders");

console.log(orders.toRecords());
console.log(orders.toCsv({ newline: "\r\n" }));
console.log(workbook.toJSON());
```

## Read from a URL and save locally

```js
import { openWorkbook } from "@entree_pos/xlsx";

const workbook = await openWorkbook(
  "https://example.com/templates/report.xlsx"
);

workbook.sheet(0).cell("B2").set("Updated by Entree POS");
await workbook.save("report-updated.xlsx");
```

## Create multiple sheets

```js
import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("Orders");
workbook.sheet("Orders").setData([
  { order: 1001, customerId: 1, total: 18.5 }
]);

workbook.addSheet("Customers", [
  { id: 1, name: "Ada" }
]);

workbook.addSheet("Summary")
  .cell("A1")
  .set("Workbook Summary")
  .style({ bold: true, fontSize: 16 });

await workbook.save("business-data.xlsx");
```

---

# Command-line API

The package installs the `entree-xlsx` command.

## Inspect a workbook

```bash
entree-xlsx inspect orders.xlsx
```

Output:

```json
{
  "file": "C:\\reports\\orders.xlsx",
  "sheets": [
    { "name": "Orders", "range": "A1:C20" }
  ]
}
```

## Convert a worksheet to JSON

```bash
entree-xlsx convert orders.xlsx orders.json --sheet Orders
```

Without `--sheet`, JSON output contains an object with every worksheet.

## Convert a worksheet to CSV

```bash
entree-xlsx convert orders.xlsx orders.csv --sheet Orders
```

Without `--sheet`, the first worksheet is used for CSV.

## Copy or normalize an XLSX workbook

```bash
entree-xlsx convert input.xlsx output.xlsx
```

The workbook is parsed and written using the supported feature set.

---

# Supported scope and limitations

The library supports modern XLSX/XLSM data, layout, reporting, and security
features:

- Strings, numbers, booleans, blanks, and dates
- Formulas with optional cached values
- Shared strings and inline strings when reading
- Fonts, fills, borders, alignment, protection, and number formats
- Merged cells
- External and internal hyperlinks
- Auto-filters
- Column widths and row heights
- Row insert, copy, and delete operations
- Worksheet add, rename, and remove operations
- Core document properties
- Column, bar, line, pie, and scatter charts
- PivotTables with worksheet sources, cache definitions, cache records, row,
  column, filter, and value fields
- Workbook and worksheet editing protection
- AES-256 Agile Office password encryption and decryption
- OLE compound-file reading and writing for encrypted Office documents
- ZIP entries using stored or deflated compression
- Lossless copying of unknown OOXML package parts and relationships

Existing package parts are preserved even when no high-level editing API exists,
including VBA projects, images, embedded objects, form controls, comments,
conditional formatting, data validation, unknown extensions, and chart or
PivotTable features outside the supported creation API.

Current limitations:

- No legacy `.xls`, `.xlsb`, `.ods`, or Numbers reader
- No formula calculation engine
- No high-level editor yet for images, comments, validation, conditional
  formatting, form controls, or VBA code
- New ZIP output does not use ZIP64
- Complex formula rewrites after row operations may need application-specific
  review

When opening an existing `.xlsm`, use an `.xlsm` output filename so users and
spreadsheet applications continue to recognize it as macro-enabled. Excel or
another spreadsheet application recalculates formulas when the file opens.
