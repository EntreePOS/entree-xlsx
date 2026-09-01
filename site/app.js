const root = document.documentElement;
const themeButton = document.querySelector(".theme-button");
const savedTheme = localStorage.getItem("entree-xlsx-theme");

if (savedTheme === "light" || savedTheme === "dark") {
  root.dataset.theme = savedTheme;
} else if (window.matchMedia("(prefers-color-scheme: dark)").matches) {
  root.dataset.theme = "dark";
}

function currentTheme() {
  return root.dataset.theme === "dark" ? "dark" : "light";
}

function updateThemeLabel() {
  themeButton.textContent = currentTheme() === "dark" ? "Light" : "Dark";
  themeButton.setAttribute("aria-label", `Use ${themeButton.textContent.toLowerCase()} theme`);
}

updateThemeLabel();

themeButton.addEventListener("click", () => {
  const next = currentTheme() === "dark" ? "light" : "dark";
  root.dataset.theme = next;
  localStorage.setItem("entree-xlsx-theme", next);
  updateThemeLabel();
});

async function copyText(value) {
  try {
    await navigator.clipboard.writeText(value);
    return;
  } catch {
    const textarea = document.createElement("textarea");
    textarea.value = value;
    textarea.setAttribute("readonly", "");
    textarea.style.position = "fixed";
    textarea.style.opacity = "0";
    document.body.append(textarea);
    textarea.select();
    const copied = document.execCommand("copy");
    textarea.remove();
    if (!copied) throw new Error("The browser blocked clipboard access.");
  }
}

document.querySelectorAll(".copy-button").forEach((button) => {
  button.addEventListener("click", async () => {
    const explicitText = button.dataset.copy;
    const code = button.closest(".code-panel")?.querySelector("code")?.textContent;
    const value = explicitText ?? code;
    if (!value) return;

    const original = button.textContent;
    try {
      await copyText(value);
      button.textContent = "Copied";
    } catch {
      button.textContent = "Copy failed";
    }
    window.setTimeout(() => { button.textContent = original; }, 1400);
  });
});

const links = [...document.querySelectorAll(".contents a[href^='#']")];
const lessons = [...document.querySelectorAll(".lesson, .helper-guide")];

const visibleLessons = new Map();
const observer = new IntersectionObserver((entries) => {
  entries.forEach((entry) => {
    if (entry.isIntersecting) visibleLessons.set(entry.target.id, entry);
    else visibleLessons.delete(entry.target.id);
  });
  const visible = [...visibleLessons.values()]
    .sort((a, b) => Math.abs(a.boundingClientRect.top) - Math.abs(b.boundingClientRect.top))[0];
  if (!visible) return;

  links.forEach((link) => {
    const active = link.hash === `#${visible.target.id}`;
    link.classList.toggle("active", active);
    if (active) link.setAttribute("aria-current", "step");
    else link.removeAttribute("aria-current");
  });
}, { rootMargin: "-18% 0px -64%", threshold: [0, .25, .5] });

lessons.forEach((lesson) => observer.observe(lesson));

const parameterHelp = {
  address: "An Excel cell address such as A1 or D12.",
  before: "The one-based row number where new rows will be inserted.",
  bytes: "XLSX file data as a Buffer, Uint8Array, or ArrayBuffer.",
  callback: "A function called once for every cell in the range.",
  changes: "An object containing only the properties you want to update.",
  column: "A column letter, such as B, or a supported column reference.",
  config: "The complete configuration for the new PivotTable.",
  count: "How many rows the operation should insert, delete, or copy.",
  data: "A two-dimensional row array or an array of JavaScript objects.",
  definitions: "An object containing multiple named style definitions.",
  format: "An Excel number-format code such as $#,##0.00 or 0.0%.",
  formula: "An Excel formula expression stored in the cell.",
  height: "The row height in points.",
  mode: "How the new style should combine with existing formatting.",
  matcher: "A cell value or callback that identifies matching cells.",
  name: "The name used for the workbook item being created or requested.",
  options: "Optional settings that change how this method runs.",
  parts: "The style properties to clear while leaving other formatting intact.",
  password: "The password required to open the encrypted workbook.",
  path: "The local file path to read from or save to.",
  records: "An array of objects where object keys map to column headers.",
  range: "An Excel cell range such as A1:D20.",
  reference: "A name, zero-based index, or object ID identifying an existing item.",
  result: "An optional cached value shown before Excel recalculates the formula.",
  row: "The one-based worksheet row number.",
  rows: "A two-dimensional array where each inner array is one worksheet row.",
  sheet: "A worksheet name, zero-based index, or Worksheet object.",
  source: "The workbook input, cell, range, or style being read or copied.",
  start: "The first one-based row affected by the operation.",
  style: "A named style, style object, or list of styles to apply.",
  target: "The destination cell, row, or hyperlink URL used by this method.",
  tooltip: "Optional hover text displayed for the hyperlink in Excel.",
  value: "The JavaScript value to store in the cell.",
  width: "The Excel column width to apply."
};

const parameterTypes = {
  address: "string · object",
  before: "number",
  bytes: "binary",
  callback: "function",
  changes: "object",
  column: "string · number",
  config: "object",
  count: "number",
  data: "array",
  definitions: "object",
  format: "string",
  formula: "string",
  height: "number",
  mode: "string",
  matcher: "any · function",
  name: "string",
  options: "object",
  parts: "string · string[]",
  password: "string",
  path: "string",
  records: "object[]",
  range: "string",
  reference: "string · number",
  result: "any",
  row: "number",
  rows: "array[]",
  sheet: "string · number",
  source: "value",
  start: "number",
  style: "style input",
  target: "value",
  tooltip: "string",
  value: "any",
  width: "number"
};

const signatureHelp = {
  "createWorkbook(name?)": {
    name: "The name of the first worksheet. Excel uses Sheet1 when omitted."
  },
  "openWorkbook(source, options?)": {
    source: "A local path, URL, Buffer, Uint8Array, or ArrayBuffer containing an XLSX file."
  },
  "workbook.addSheet(name, data?)": {
    name: "The unique name shown on the new worksheet tab.",
    data: "Optional rows or object records used to populate the new sheet immediately."
  },
  "styles.define(name, style, options?)": {
    name: "The reusable style name you will pass to cell.style() or range.style()."
  },
  "cell.copyStyleFrom(source, mode?)": {
    source: "The Cell whose formatting should be copied."
  },
  "range.copyStyleFrom(source, options?)": {
    source: "The source Range whose formatting should be copied."
  },
  "cell.hyperlink(target, tooltip?)": {
    target: "The URL, email address, or workbook location opened by the link."
  },
  "charts.list(sheet?)": {
    sheet: "Optional worksheet filter. Omit it to list charts from the entire workbook."
  },
  "pivotTables.list(sheet?)": {
    sheet: "Optional worksheet filter. Omit it to list every PivotTable."
  },
  "workbook.save(path, { password })": {
    password: "The password Excel will require before opening the encrypted file."
  }
};

const signatureTypes = {
  "openWorkbook(source, options?)": { source: "path · URL · binary" },
  "workbook.addSheet(name, data?)": { data: "any[][] · object[]" },
  "sheet.setData(data, options?)": { data: "any[][] · object[]" },
  "sheet.appendRows(rows, options?)": { rows: "any[][] · object[]" },
  "sheet.get(address)": { address: "string · {r,c}" },
  "sheet.set(address, value)": { address: "string · {r,c}" },
  "sheet.cell(address)": { address: "string · {r,c}" },
  "sheet.copyRow(source, target, options?)": { source: "number", target: "number" },
  "range.setValues(rows)": { rows: "any[][]" },
  "cell.style(style, mode?)": { style: "name · object · array", mode: "merge · replace" },
  "range.style(style, mode?)": { style: "name · object · array", mode: "merge · replace" },
  "cell.copyStyleFrom(source, mode?)": { source: "Cell · string", mode: "merge · replace" },
  "range.copyStyleFrom(source, options?)": { source: "Range · string" },
  "cell.formula(formula, result?)": { result: "any" },
  "cell.hyperlink(target, tooltip?)": { target: "string" },
  "charts.list(sheet?)": { sheet: "string · number" },
  "charts.update(reference, changes)": { reference: "name · ID" },
  "charts.remove(reference)": { reference: "name · ID" },
  "pivotTables.list(sheet?)": { sheet: "string · number" },
  "pivotTables.update(reference, changes)": { reference: "name · ID" },
  "pivotTables.remove(reference)": { reference: "name · ID" },
  "range.forEach(callback)": { callback: "function" }
};

const openOptions = {
  title: "Open options",
  properties: [
    { name: "password?", type: "string", detail: "Password for an encrypted Office file." },
    { name: "encryption.verifyIntegrity?", type: "boolean", default: "true", detail: "Verify the encrypted file HMAC before opening." }
  ],
  example: '{ password: "demo" }'
};

const saveOptions = {
  title: "Save options",
  properties: [
    { name: "password?", type: "string", detail: "Encrypt the XLSX with this open password." },
    { name: "encryption.spinCount?", type: "number", default: "100000", detail: "Advanced password-key derivation work factor." }
  ],
  example: '{ password: "demo" }'
};

const recordOptions = {
  title: "Record options",
  properties: [
    { name: "headers?", type: "string[]", detail: "Use these keys and treat every selected row as data." },
    { name: "range?", type: "string", default: "usedRange", detail: "Read only this A1 range." },
    { name: "defaultValue?", type: "any", default: "null", detail: "Value returned for missing cells." }
  ],
  example: '{ range: "A1:C20", defaultValue: null }'
};

const styleObject = {
  title: "Style object",
  properties: [
    { name: "bold?, italic?, strike?", type: "boolean", detail: "Common font emphasis." },
    { name: "fontName?, fontSize?", type: "string, number", detail: "Font family and point size." },
    { name: "color?, fill?", type: "color | object", detail: "Font color and cell background." },
    { name: "horizontal?, vertical?", type: "string", detail: "Cell alignment." },
    { name: "wrapText?, shrinkToFit?", type: "boolean", detail: "Control long cell text." },
    { name: "numberFormat?", type: "string", detail: "Excel format such as $#,##0.00." },
    { name: "border?", type: "object", detail: "top, right, bottom, left, outline, or inside borders." },
    { name: "protection?", type: "object", detail: "{ locked?, hidden? }; active after sheet protection." }
  ],
  example: '{ bold: true, fill: "#17324D", color: "#FFFFFF" }',
  footnote: "Also accepts a named style string or an array of styles."
};

const chartOptions = {
  title: "Chart options",
  properties: [
    { name: "sheet", type: "string | number", detail: "Worksheet containing the chart data." },
    { name: "range?", type: "string", detail: "Header-first range, such as A1:C13. Use this or series." },
    { name: "series?", type: "object[]", detail: "{ name?, nameCell?, categories?, xValues?, values }." },
    { name: "type?", type: "string", default: "column", detail: "column, bar, line, pie, or scatter." },
    { name: "name?", type: "string", detail: "Internal chart name." },
    { name: "title?", type: "string", detail: "Visible chart title." },
    { name: "position?", type: "object", default: "E2:M18", detail: "{ from: \"E2\", to: \"M18\" }." },
    { name: "legend?", type: "boolean", default: "true", detail: "Set false to hide the legend." },
    { name: "legendPosition?", type: "string", default: "r", detail: "Excel legend position code." }
  ],
  example: '{ sheet: "Sales", type: "column", range: "A1:C13" }'
};

const chartChanges = {
  ...chartOptions,
  title: "Chart changes",
  properties: chartOptions.properties.filter((property) => property.name !== "sheet"),
  example: '{ type: "line", range: "A1:C13", title: "Revenue trend" }',
  footnote: "Pass range or series on every update so the data mapping is explicit."
};

const pivotOptions = {
  title: "PivotTable configuration",
  properties: [
    { name: "source", type: "object", detail: "Required: { sheet, range? }. Range defaults to used data." },
    { name: "target", type: "object", detail: "Required: { sheet, cell }, such as Summary!A3." },
    { name: "name?", type: "string", detail: "Unique PivotTable name." },
    { name: "rows?, columns?, filters?", type: "string[]", detail: "Source header names assigned to each area." },
    { name: "values", type: "array", detail: "Required: strings or { field, summarize?, name? }." },
    { name: "refreshOnLoad?", type: "boolean", default: "true", detail: "Ask Excel to refresh when opened." },
    { name: "showGrandTotals?", type: "boolean", default: "true", detail: "Show row and column grand totals." },
    { name: "style?", type: "string", default: "PivotStyleMedium9", detail: "Built-in Excel PivotTable style name." }
  ],
  example: '{ source: { sheet: "Orders" }, target: { sheet: "Summary", cell: "A3" }, rows: ["Region"], values: ["Sales"] }',
  footnote: "summarize accepts sum, count, average, min, or max."
};

const pivotChanges = {
  ...pivotOptions,
  title: "PivotTable changes",
  properties: pivotOptions.properties.map((property) => ({
    ...property,
    name: property.name.endsWith("?") ? property.name : `${property.name}?`
  })),
  example: '{ rows: ["Store"], columns: ["Region"] }',
  footnote: "Only include the fields you want to replace."
};

const parameterSchemas = {
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
        { name: "header?", type: "string[]", default: "inferred", detail: "Property order for object records." },
        { name: "skipHeader?", type: "boolean", default: "false", detail: "Do not write the object-key header row." }
      ],
      example: '{ header: ["order", "customer", "total"] }'
    }
  },
  "sheet.appendData(records, options?)": {
    options: {
      title: "Append options",
      properties: [
        { name: "header?", type: "string[]", default: "first used row", detail: "Explicit object property order." },
        { name: "origin?", type: "A1 | number | {r,c}", default: "next used row", detail: "Override the first destination cell." }
      ],
      example: '{ header: ["order", "customer", "total"] }'
    }
  },
  "sheet.appendRows(rows, options?)": {
    options: {
      title: "Row options",
      properties: [
        { name: "origin?", type: "A1 | number | {r,c} | -1", default: "-1", detail: "Destination; -1 appends below used data." },
        { name: "header?", type: "string[]", default: "inferred", detail: "Property order for object rows." },
        { name: "skipHeader?", type: "boolean", default: "false", detail: "Do not create a header for object rows." }
      ],
      example: '{ origin: "A5", skipHeader: true }'
    }
  },
  "sheet.toRows(options?)": {
    options: {
      title: "Read options",
      properties: [
        { name: "range?", type: "string", default: "usedRange", detail: "Read only this A1 range." },
        { name: "defaultValue?", type: "any", default: "undefined", detail: "Replace empty cells with this value." }
      ],
      example: '{ range: "A1:C10", defaultValue: null }'
    }
  },
  "sheet.toRecords(options?)": { options: recordOptions },
  "sheet.toCsv(options?)": {
    options: {
      title: "CSV options",
      properties: [
        { name: "delimiter?", type: "string", default: ",", detail: "Field separator; use \\t for TSV." },
        { name: "newline?", type: "string", default: "\\n", detail: "Row separator." },
        { name: "range?", type: "string", default: "usedRange", detail: "Export only this A1 range." },
        { name: "defaultValue?", type: "any", default: "undefined", detail: "Replace empty cells." }
      ],
      example: '{ delimiter: "\\t", newline: "\\r\\n" }'
    }
  },
  "sheet.toHtml(options?)": {
    options: {
      title: "HTML options",
      properties: [
        { name: "header?", type: "boolean", default: "true", detail: "Use th elements for the first row." },
        { name: "range?", type: "string", default: "usedRange", detail: "Export only this A1 range." },
        { name: "defaultValue?", type: "any", default: "undefined", detail: "Replace empty cells." }
      ],
      example: '{ header: true, range: "A1:D20" }'
    }
  },
  "sheet.insertRows(before, count?, options?)": {
    options: {
      title: "Insert options",
      properties: [
        { name: "copyFrom?", type: '"above" | "below" | number', detail: "Copy a neighboring or one-based source row." },
        { name: "values?", type: "boolean", default: "true", detail: "Set false to copy formatting without values." }
      ],
      example: '{ copyFrom: "above", values: false }'
    }
  },
  "sheet.copyRow(source, target, options?)": {
    options: {
      title: "Copy options",
      properties: [
        { name: "values?", type: "boolean", default: "true", detail: "Set false to copy only row formatting." }
      ],
      example: '{ values: false }'
    }
  },
  "cell.style(style, mode?)": { style: styleObject },
  "range.style(style, mode?)": { style: styleObject },
  "column.style(style, mode?)": { style: styleObject },
  "styles.define(name, style, options?)": {
    style: styleObject,
    options: {
      title: "Named style options",
      properties: [
        { name: "extends?", type: "string | string[]", detail: "Parent style names, inherited from left to right." }
      ],
      example: '{ extends: ["base", "money"] }'
    }
  },
  "styles.defineMany(definitions)": {
    definitions: {
      title: "Definitions object",
      properties: [
        { name: "[styleName]", type: "style object", detail: "A direct style definition." },
        { name: "[styleName].style", type: "style object", detail: "Style body when inheritance is used." },
        { name: "[styleName].extends?", type: "string | string[]", detail: "Parent named styles." }
      ],
      example: '{ base: { fontName: "Aptos" }, header: { extends: "base", style: { bold: true } } }'
    }
  },
  "range.copyStyleFrom(source, options?)": {
    options: {
      title: "Style copy options",
      properties: [
        { name: "mode?", type: '"replace" | "merge"', default: "replace", detail: "Replace all formatting or merge selected parts." },
        { name: "repeat?", type: "boolean", default: "false", detail: "Tile source styles when range sizes differ." }
      ],
      example: '{ mode: "merge", repeat: true }'
    }
  },
  "sheet.autoFit(options?)": {
    options: {
      title: "Auto-fit options",
      properties: [
        { name: "min?", type: "number", default: "8", detail: "Minimum column width." },
        { name: "max?", type: "number", default: "60", detail: "Maximum column width." },
        { name: "padding?", type: "number", default: "2", detail: "Extra characters added to measured text." },
        { name: "includeHeader?", type: "boolean", default: "true", detail: "Set false to ignore the first used row." }
      ],
      example: '{ min: 10, max: 40, padding: 3 }'
    }
  },
  "charts.add(options)": { options: chartOptions },
  "charts.update(reference, changes)": { changes: chartChanges },
  "pivotTables.add(config)": { config: pivotOptions },
  "pivotTables.update(reference, changes)": { changes: pivotChanges },
  "cell.clear(options?)": {
    options: {
      title: "Clear options",
      properties: [
        { name: "keepStyle?", type: "boolean", default: "false", detail: "Keep formatting on the now-empty cell." }
      ],
      example: '{ keepStyle: true }'
    }
  },
  "sheet.protectSheet(options?)": {
    options: {
      title: "Worksheet protection options",
      properties: [
        { name: "password?", type: "string", detail: "Password needed to remove protection." },
        { name: "objects?, scenarios?", type: "boolean", default: "true", detail: "Protect drawings and scenarios." },
        { name: "selectLockedCells?", type: "boolean", detail: "Set false to block selecting locked cells." },
        { name: "selectUnlockedCells?", type: "boolean", detail: "Set false to block selecting unlocked cells." },
        { name: "formatCells?, formatColumns?, formatRows?", type: "boolean", detail: "Set false to block the matching formatting action." },
        { name: "insertColumns?, insertRows?", type: "boolean", detail: "Set false to block insertion." },
        { name: "deleteColumns?, deleteRows?", type: "boolean", detail: "Set false to block deletion." },
        { name: "sort?, autoFilter?, pivotTables?", type: "boolean", detail: "Set false to block these data actions." }
      ],
      example: '{ password: "demo", formatCells: false, deleteRows: false }'
    }
  },
  "workbook.protectStructure(options?)": {
    options: {
      title: "Workbook protection options",
      properties: [
        { name: "password?", type: "string", detail: "Password needed to remove protection." },
        { name: "structure?", type: "boolean", default: "true", detail: "Lock adding, deleting, renaming, and moving sheets." },
        { name: "windows?", type: "boolean", default: "false", detail: "Lock workbook window arrangement." }
      ],
      example: '{ password: "demo", structure: true }'
    }
  }
};

const parameterTooltip = document.createElement("div");
parameterTooltip.className = "parameter-tooltip";
parameterTooltip.id = "parameter-tooltip";
parameterTooltip.setAttribute("role", "tooltip");
parameterTooltip.setAttribute("aria-hidden", "true");
parameterTooltip.hidden = true;
parameterTooltip.innerHTML = `
  <div class="parameter-tooltip-heading">
    <strong class="parameter-tooltip-name"></strong>
    <span class="parameter-tooltip-kind" hidden>object</span>
  </div>
  <span class="parameter-tooltip-copy"></span>
  <div class="parameter-tooltip-schema" hidden>
    <span class="parameter-tooltip-schema-title"></span>
    <div class="parameter-tooltip-properties"></div>
    <div class="parameter-tooltip-example"><span>Example</span><code></code></div>
    <span class="parameter-tooltip-footnote" hidden></span>
  </div>`;
document.body.append(parameterTooltip);

const tooltipName = parameterTooltip.querySelector(".parameter-tooltip-name");
const tooltipCopy = parameterTooltip.querySelector(".parameter-tooltip-copy");
const tooltipKind = parameterTooltip.querySelector(".parameter-tooltip-kind");
const tooltipSchema = parameterTooltip.querySelector(".parameter-tooltip-schema");
const tooltipSchemaTitle = parameterTooltip.querySelector(".parameter-tooltip-schema-title");
const tooltipProperties = parameterTooltip.querySelector(".parameter-tooltip-properties");
const tooltipExample = parameterTooltip.querySelector(".parameter-tooltip-example");
const tooltipExampleCode = tooltipExample.querySelector("code");
const tooltipFootnote = parameterTooltip.querySelector(".parameter-tooltip-footnote");
let activeParameter;

function positionParameterTooltip() {
  if (!activeParameter || parameterTooltip.hidden) return;
  const target = activeParameter.getBoundingClientRect();
  const tooltip = parameterTooltip.getBoundingClientRect();
  const gap = 10;
  let placement = "top";
  let top = target.top - tooltip.height - gap;
  if (top < 12) {
    placement = "bottom";
    top = target.bottom + gap;
  }
  const left = Math.min(
    window.innerWidth - tooltip.width - 12,
    Math.max(12, target.left + (target.width - tooltip.width) / 2)
  );
  top = Math.min(window.innerHeight - tooltip.height - 12, Math.max(12, top));
  const arrowLeft = Math.min(
    tooltip.width - 16,
    Math.max(16, target.left + target.width / 2 - left)
  );
  parameterTooltip.dataset.placement = placement;
  parameterTooltip.style.setProperty("--tooltip-arrow-x", `${Math.round(arrowLeft)}px`);
  parameterTooltip.style.transform = `translate3d(${Math.round(left)}px, ${Math.round(top)}px, 0)`;
}

function showParameterTooltip(parameter) {
  activeParameter = parameter;
  tooltipName.textContent = parameter.textContent;
  tooltipCopy.textContent = parameter.dataset.tooltip;
  const schema = parameter.dataset.schema ? JSON.parse(parameter.dataset.schema) : undefined;
  tooltipKind.textContent = parameter.dataset.parameterType;
  tooltipKind.hidden = false;
  tooltipSchema.hidden = !schema;
  parameterTooltip.classList.toggle("has-schema", Boolean(schema));
  tooltipProperties.replaceChildren();
  if (schema) {
    tooltipSchemaTitle.textContent = schema.title;
    for (const property of schema.properties) {
      const row = document.createElement("div");
      row.className = "parameter-tooltip-property";
      const key = document.createElement("code");
      key.textContent = property.name;
      const details = document.createElement("div");
      const meta = document.createElement("span");
      meta.className = "parameter-tooltip-property-meta";
      meta.textContent = property.default === undefined
        ? property.type
        : `${property.type} · default ${property.default}`;
      const copy = document.createElement("span");
      copy.textContent = property.detail;
      details.append(meta, copy);
      row.append(key, details);
      tooltipProperties.append(row);
    }
    tooltipExample.hidden = !schema.example;
    tooltipExampleCode.textContent = schema.example ?? "";
    tooltipFootnote.hidden = !schema.footnote;
    tooltipFootnote.textContent = schema.footnote ?? "";
  }
  parameterTooltip.hidden = false;
  parameterTooltip.setAttribute("aria-hidden", "false");
  positionParameterTooltip();
}

function hideParameterTooltip(parameter) {
  if (parameter && activeParameter !== parameter) return;
  activeParameter = undefined;
  parameterTooltip.hidden = true;
  parameterTooltip.setAttribute("aria-hidden", "true");
}

document.querySelectorAll(".api-table td:first-child code").forEach((code) => {
  const signature = code.textContent;
  const open = signature.indexOf("(");
  const close = signature.lastIndexOf(")");
  if (open < 0 || close <= open + 1) return;

  const prefix = signature.slice(0, open + 1);
  const parameters = signature.slice(open + 1, close);
  const suffix = signature.slice(close);
  const fragment = document.createDocumentFragment();
  fragment.append(document.createTextNode(prefix));

  let cursor = 0;
  for (const match of parameters.matchAll(/[A-Za-z_$][\w$]*\??/g)) {
    fragment.append(document.createTextNode(parameters.slice(cursor, match.index)));
    const label = match[0];
    const name = label.replace(/\?$/, "");
    const description = signatureHelp[signature]?.[name] ?? parameterHelp[name];
    const schema = parameterSchemas[signature]?.[name];
    const parameterType = signatureTypes[signature]?.[name] ?? parameterTypes[name];
    if (!description) {
      fragment.append(document.createTextNode(label));
    } else {
      const parameter = document.createElement("span");
      parameter.className = "api-param";
      parameter.tabIndex = 0;
      parameter.textContent = label;
      parameter.dataset.tooltip = description;
      parameter.dataset.parameterType = parameterType;
      if (schema) parameter.dataset.schema = JSON.stringify(schema);
      parameter.setAttribute("aria-describedby", parameterTooltip.id);
      parameter.setAttribute("aria-label", schema
        ? `${name} parameter, type ${parameterType}: ${description} Accepted properties: ${schema.properties.map((property) => property.name).join(", ")}.`
        : `${name} parameter, type ${parameterType}: ${description}`);
      parameter.addEventListener("pointerenter", () => showParameterTooltip(parameter));
      parameter.addEventListener("pointerleave", () => {
        if (document.activeElement !== parameter) hideParameterTooltip(parameter);
      });
      parameter.addEventListener("focus", () => showParameterTooltip(parameter));
      parameter.addEventListener("blur", () => hideParameterTooltip(parameter));
      parameter.addEventListener("keydown", (event) => {
        if (event.key === "Escape") parameter.blur();
      });
      fragment.append(parameter);
    }
    cursor = match.index + label.length;
  }

  fragment.append(document.createTextNode(parameters.slice(cursor)));
  fragment.append(document.createTextNode(suffix));
  code.replaceChildren(fragment);
});

window.addEventListener("resize", positionParameterTooltip);
window.addEventListener("scroll", positionParameterTooltip, { passive: true });
