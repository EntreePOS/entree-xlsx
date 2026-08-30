# Changelog

## 0.2.1

- Add repository, homepage, and issue-tracker links for the public GitHub
  project.
- Rewrite the README as a programmer-focused introduction with concise examples
  for styling, template editing, charts, PivotTables, and encryption.

## 0.1.0

- Add a native ESM JavaScript API for Node.js 18 and newer.
- Implement XLSX ZIP and OOXML reading/writing with Node built-ins only.
- Ship with zero runtime dependencies and no TypeScript layer.
- Introduce friendly `Workbook`, `Worksheet`, `Cell`, and `Range` objects.
- Support values, dates, formulas, styles, merges, hyperlinks, filters, widths,
  heights, metadata, object rows, array rows, JSON, CSV, and HTML conversion.
- Add file, URL, Buffer, Uint8Array, and ArrayBuffer inputs.
- Add the `entree-xlsx` inspect and convert command-line tool.
- Preserve untouched XLSX/XLSM OOXML parts, relationships, macros, drawings,
  extensions, and embedded content when editing templates.
- Add chart create/list/update/remove APIs for column, bar, line, pie, and
  scatter charts.
- Add PivotTable create/list/update/remove APIs with worksheet caches and
  summary rendering.
- Add AES-256 Agile Office password encryption and decryption.
- Add workbook and worksheet protection, template style and hyperlink edits,
  and row insert/copy/delete operations.
- Add first-class Excel named styles with inheritance, atomic bulk definition,
  composition, inspection, removal, and cell/range copy helpers.
- Expand styles with nested fonts, pattern and gradient fills, alignment,
  number formats, cell protection, diagonal borders, and range-aware outline
  and interior borders.
- Deduplicate cell formats and style resources when formatting opened templates
  so large styled ranges do not inflate the workbook style table.
