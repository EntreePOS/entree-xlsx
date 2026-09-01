const openOptions = {
  title: "Open options",
  properties: [
    ["password?", "string", "", "Password for an encrypted Office file."],
    ["encryption.verifyIntegrity?", "boolean", "true", "Verify the encrypted file HMAC before opening."]
  ],
  example: '{ password: "demo" }'
};

const saveOptions = {
  title: "Save options",
  properties: [
    ["password?", "string", "", "Encrypt the XLSX with this open password."],
    ["encryption.spinCount?", "number", "100000", "Password-key derivation work factor."]
  ],
  example: '{ password: "demo" }'
};

const recordOptions = {
  title: "Record options",
  properties: [
    ["headers?", "string[]", "", "Use these keys and treat every selected row as data."],
    ["range?", "string", "usedRange", "Read only this A1 range."],
    ["defaultValue?", "any", "null", "Value returned for missing cells."]
  ],
  example: '{ range: "A1:C20", defaultValue: null }'
};

const styleObject = {
  title: "Style object",
  properties: [
    ["bold?, italic?, strike?", "boolean", "", "Font emphasis."],
    ["fontName?, fontSize?", "string, number", "", "Font family and point size."],
    ["color?, fill?", "color | object", "", "Font color and cell background."],
    ["horizontal?, vertical?", "string", "", "Cell alignment."],
    ["wrapText?, shrinkToFit?", "boolean", "", "Control long cell text."],
    ["numberFormat?", "string", "", "Excel format such as $#,##0.00."],
    ["border?", "object", "", "Top, right, bottom, left, outline, or inside borders."],
    ["protection?", "object", "", "{ locked?, hidden? }; active after sheet protection."]
  ],
  example: '{ bold: true, fill: "#17324D", color: "#FFFFFF" }'
};

const chartOptions = {
  title: "Chart options",
  properties: [
    ["sheet", "string | number", "", "Worksheet containing the chart data."],
    ["range?", "string", "", "Header-first range such as A1:C13. Use this or series."],
    ["series?", "object[]", "", "{ name?, nameCell?, categories?, xValues?, values }."],
    ["type?", "string", "column", "column, bar, line, pie, or scatter."],
    ["name?", "string", "", "Internal chart name."],
    ["title?", "string", "", "Visible chart title."],
    ["position?", "object", "E2:M18", "{ from: \"E2\", to: \"M18\" }."],
    ["legend?", "boolean", "true", "Set false to hide the legend."],
    ["legendPosition?", "string", "r", "Excel legend position code."]
  ],
  example: '{ sheet: "Sales", type: "column", range: "A1:C13" }'
};

const pivotOptions = {
  title: "PivotTable configuration",
  properties: [
    ["source", "object", "", "Required: { sheet, range? }. Range defaults to used data."],
    ["target", "object", "", "Required: { sheet, cell }, such as Summary!A3."],
    ["name?", "string", "", "Unique PivotTable name."],
    ["rows?, columns?, filters?", "string[]", "", "Source headers assigned to each area."],
    ["values", "array", "", "Required: strings or { field, summarize?, name? }."],
    ["refreshOnLoad?", "boolean", "true", "Ask Excel to refresh when opened."],
    ["showGrandTotals?", "boolean", "true", "Show row and column grand totals."],
    ["style?", "string", "PivotStyleMedium9", "Built-in Excel PivotTable style name."]
  ],
  example: '{ source: { sheet: "Orders" }, target: { sheet: "Summary", cell: "A3" }, rows: ["Region"], values: ["Sales"] }'
};

export const parameterSchemas = {
  "openWorkbook(source, options?)": { options: openOptions },
  "openWorkbookSync(path, options?)": { options: openOptions },
  "parseWorkbook(bytes, options?)": { options: openOptions },
  "workbook.save(path, options?)": { options: saveOptions },
  "workbook.saveSync(path, options?)": { options: saveOptions },
  "workbook.toBuffer(options?)": { options: saveOptions },
  "workbook.toUint8Array(options?)": { options: saveOptions },
  "workbook.toBase64(options?)": { options: saveOptions },
  "workbook.toJSON(options?)": { options: recordOptions },
  "sheet.setData(data, options?)": {
    options: {
      title: "Data options",
      properties: [
        ["header?", "string[]", "inferred", "Property order for object records."],
        ["skipHeader?", "boolean", "false", "Do not write the object-key header row."]
      ],
      example: '{ header: ["order", "customer", "total"] }'
    }
  },
  "sheet.appendData(records, options?)": {
    options: {
      title: "Append options",
      properties: [
        ["header?", "string[]", "first used row", "Explicit object property order."],
        ["origin?", "A1 | number | {r,c}", "next used row", "Override the first destination cell."]
      ],
      example: '{ header: ["order", "customer", "total"] }'
    }
  },
  "sheet.appendRows(rows, options?)": {
    options: {
      title: "Row options",
      properties: [
        ["origin?", "A1 | number | {r,c} | -1", "-1", "Destination; -1 appends below used data."],
        ["header?", "string[]", "inferred", "Property order for object rows."],
        ["skipHeader?", "boolean", "false", "Do not create a header for object rows."]
      ],
      example: '{ origin: "A5", skipHeader: true }'
    }
  },
  "sheet.toRows(options?)": {
    options: {
      title: "Read options",
      properties: [
        ["range?", "string", "usedRange", "Read only this A1 range."],
        ["defaultValue?", "any", "undefined", "Replace empty cells with this value."]
      ],
      example: '{ range: "A1:C10", defaultValue: null }'
    }
  },
  "sheet.toRecords(options?)": { options: recordOptions },
  "sheet.toCsv(options?)": {
    options: {
      title: "CSV options",
      properties: [
        ["delimiter?", "string", ",", "Field separator; use \\t for TSV."],
        ["newline?", "string", "\\n", "Row separator."],
        ["range?", "string", "usedRange", "Export only this A1 range."],
        ["defaultValue?", "any", "undefined", "Replace empty cells."]
      ],
      example: '{ delimiter: "\\t", newline: "\\r\\n" }'
    }
  },
  "sheet.toHtml(options?)": {
    options: {
      title: "HTML options",
      properties: [
        ["header?", "boolean", "true", "Use th elements for the first row."],
        ["range?", "string", "usedRange", "Export only this A1 range."],
        ["defaultValue?", "any", "undefined", "Replace empty cells."]
      ],
      example: '{ header: true, range: "A1:D20" }'
    }
  },
  "sheet.insertRows(before, count?, options?)": {
    options: {
      title: "Insert options",
      properties: [
        ["copyFrom?", '"above" | "below" | number', "", "Copy a neighboring or one-based source row."],
        ["values?", "boolean", "true", "Set false to copy formatting without values."]
      ],
      example: '{ copyFrom: "above", values: false }'
    }
  },
  "sheet.copyRow(source, target, options?)": {
    options: {
      title: "Copy options",
      properties: [["values?", "boolean", "true", "Set false to copy only row formatting."]],
      example: '{ values: false }'
    }
  },
  "cell.style(style, mode?)": { style: styleObject },
  "range.style(style, mode?)": { style: styleObject },
  "column.style(style, mode?)": { style: styleObject },
  "columns.style(style, mode?)": { style: styleObject },
  "row.style(style, mode?)": { style: styleObject },
  "rows.style(style, mode?)": { style: styleObject },
  "styles.define(name, style, options?)": {
    style: styleObject,
    options: {
      title: "Named style options",
      properties: [["extends?", "string | string[]", "", "Parent style names, inherited from left to right."]],
      example: '{ extends: ["base", "money"] }'
    }
  },
  "styles.defineMany(definitions)": {
    definitions: {
      title: "Definitions object",
      properties: [
        ["[styleName]", "style object", "", "A direct style definition."],
        ["[styleName].style", "style object", "", "Style body when inheritance is used."],
        ["[styleName].extends?", "string | string[]", "", "Parent named styles."]
      ],
      example: '{ base: { fontName: "Aptos" }, header: { extends: "base", style: { bold: true } } }'
    }
  },
  "range.copyStyleFrom(source, options?)": {
    options: {
      title: "Style copy options",
      properties: [
        ["mode?", '"replace" | "merge"', "replace", "Replace all formatting or merge selected parts."],
        ["repeat?", "boolean", "false", "Tile source styles when range sizes differ."]
      ],
      example: '{ mode: "merge", repeat: true }'
    }
  },
  "sheet.autoFit(options?)": {
    options: {
      title: "Auto-fit options",
      properties: [
        ["min?", "number", "8", "Minimum column width."],
        ["max?", "number", "60", "Maximum column width."],
        ["padding?", "number", "2", "Extra characters added to measured text."],
        ["includeHeader?", "boolean", "true", "Set false to ignore the first used row."]
      ],
      example: '{ min: 10, max: 40, padding: 3 }'
    }
  },
  "charts.add(options)": { options: chartOptions },
  "charts.update(reference, changes)": { changes: { ...chartOptions, title: "Chart changes" } },
  "pivotTables.add(config)": { config: pivotOptions },
  "pivotTables.update(reference, changes)": { changes: { ...pivotOptions, title: "PivotTable changes" } },
  "cell.clear(options?)": {
    options: {
      title: "Clear options",
      properties: [["keepStyle?", "boolean", "false", "Keep formatting on the empty cell."]],
      example: '{ keepStyle: true }'
    }
  },
  "sheet.protectSheet(options?)": {
    options: {
      title: "Worksheet protection options",
      properties: [
        ["password?", "string", "", "Password needed to remove protection."],
        ["objects?, scenarios?", "boolean", "true", "Protect drawings and scenarios."],
        ["selectLockedCells?", "boolean", "", "Set false to block selecting locked cells."],
        ["selectUnlockedCells?", "boolean", "", "Set false to block selecting unlocked cells."],
        ["formatCells?, formatColumns?, formatRows?", "boolean", "", "Set false to block formatting."],
        ["insertColumns?, insertRows?", "boolean", "", "Set false to block insertion."],
        ["deleteColumns?, deleteRows?", "boolean", "", "Set false to block deletion."],
        ["sort?, autoFilter?, pivotTables?", "boolean", "", "Set false to block these data actions."]
      ],
      example: '{ password: "demo", formatCells: false, deleteRows: false }'
    }
  },
  "workbook.protectStructure(options?)": {
    options: {
      title: "Workbook protection options",
      properties: [
        ["password?", "string", "", "Password needed to remove protection."],
        ["structure?", "boolean", "true", "Lock adding, deleting, renaming, and moving sheets."],
        ["windows?", "boolean", "false", "Lock workbook window arrangement."]
      ],
      example: '{ password: "demo", structure: true }'
    }
  }
};
