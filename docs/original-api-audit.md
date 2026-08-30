# Original API audit and rewrite decisions

The source package `@sheet/edit` is a bundled SheetJS Pro edit build. The audit
identified three major areas: multi-format file I/O, raw workbook utilities,
and licensed template-preservation operations.

The old surface is broad but difficult to discover. It exposes abbreviated
internal fields, mixes low-level conversion with file operations, and returns
errors with limited workbook context. Its single bundled implementation is also
hard to maintain or safely republish under another package name.

## Rewrite direction

The new package intentionally does not preserve the old API. It uses modern ESM
JavaScript and requires Node.js 18 or newer. The public model is organized around
four objects:

- `Workbook` manages sheets, properties, and XLSX output.
- `Worksheet` manages tabular data, layout, and conversions.
- `Range` manages a rectangular cell selection.
- `Cell` manages values, formulas, styles, and hyperlinks.

Entry points have explicit intent:

- `createWorkbook(name)` starts a new XLSX workbook.
- `openWorkbook(source)` reads a path, URL, or binary value asynchronously.
- `openWorkbookSync(path)` reads a local file synchronously.
- `parseWorkbook(bytes)` reads in-memory XLSX bytes.

## Dependency decision

The rewrite uses no third-party runtime libraries. ZIP container handling uses
Node's built-in `zlib`; XML handling is implemented for the OOXML structures the
package supports. This removes the transitive dependency tree and gives the
package a zero-vulnerability npm audit baseline.

## Supported migration concepts

| Old concept | New API |
| --- | --- |
| create a workbook | `createWorkbook()` |
| select a sheet | `workbook.sheet(nameOrIndex)` |
| append object rows | `sheet.addRows(objects)` |
| read object rows | `sheet.toRecords()` |
| access a cell | `sheet.cell("A1")` |
| style a range | `sheet.range("A1:C3").style({...})` |
| write a file | `await workbook.save("file.xlsx")` |

This is a conceptual migration guide, not a compatibility promise.

## Intentional boundaries

The rewrite focuses on common XLSX workbook data. It does not claim preservation
of macros, charts, pivot internals, form controls, embedded binaries, or unknown
OOXML extensions. It also drops XLS, XLSB, ODS, DBF, SYLK, and other legacy
formats. These boundaries are explicit so the API remains dependable within its
documented scope.
