import { parameterSchemas } from "./api-details.js";

const samples = {
  data: `import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("Menu");
const sheet = workbook.sheet();

sheet.setData([
  ["Item", "Category", "Price"],
  ["Classic Burger", "Burger", 12.5],
  ["Fries", "Side", 4],
  ["Chocolate Shake", "Drink", 5.25]
]);

await workbook.save("menu.xlsx");`,
  styles: `import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("Sales");
const sheet = workbook.sheet();

sheet.setData([
  ["Location", "Revenue"],
  ["North Market", 1840],
  ["Park Bistro", 2175],
  ["Lake Cafe", 1560]
]);

sheet.row(1)
  .style({ bold: true, color: "#FFFFFF", fill: "#2457C5" })
  .height(26);

sheet.column("B")
  .style({ numberFormat: "$#,##0.00" })
  .width(16);

sheet.column("A").width(22);
await workbook.save("sales.xlsx");`,
  helpers: `import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("Inventory");
const sheet = workbook.sheet();

sheet.setData([
  ["Item", "Stock", "Status"],
  ["Classic Burger", 18, "Low"],
  ["Fries", 42, "Ready"],
  ["Chocolate Shake", 12, "Low"]
]);

const burger = sheet.column("A").find("Classic Burger");
console.log("Found at", burger?.address);

sheet.rows("2:4").height(24);
sheet.columns("A:C").width(18);
sheet.column("C").findAll("Low").forEach((cell) => {
  cell.style({ bold: true, color: "#A23B22", fill: "#FFF0E8" });
});

await workbook.save("inventory.xlsx");`,
  formula: `import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("Order");
const sheet = workbook.sheet();

sheet.setData([
  ["Item", "Qty", "Price", "Total"],
  ["Classic Burger", 2, 12.5, null],
  ["Fries", 1, 4, null]
]);

sheet.cell("D2").formula("B2*C2", 25);
sheet.cell("D3").formula("B3*C3", 4);
sheet.cell("D4").formula("SUM(D2:D3)", 29);
sheet.set("C4", "Grand total");

sheet.range("A1:D1").style({ bold: true, fill: "#E8EFFF" });
sheet.range("C2:D4").style({ numberFormat: "$#,##0.00" });
sheet.cell("D4").style({ bold: true });
sheet.autoFit();

await workbook.save("order-total.xlsx");`
};

const sample = document.querySelector("#sample");
const editor = document.querySelector("#code");
const runButton = document.querySelector("#run");
const resetButton = document.querySelector("#reset");
const downloadButton = document.querySelector("#download");
const preview = document.querySelector("#preview");
const tabs = document.querySelector("#sheet-tabs");
const status = document.querySelector("#status");
const consoleOutput = document.querySelector("#console");
const completionMenu = document.querySelector("#editor-completions");
const apiTooltip = document.querySelector("#editor-api-tooltip");
let result;
let activeSheet = 0;
let completionMatches = [];
let completionIndex = 0;
let completionContext;
let hoveredApiName = "";

const fallbackApis = [
  ["createWorkbook(name?)", "Create a workbook with one named sheet."],
  ["openWorkbook(source, options?)", "Open an XLSX workbook."],
  ["workbook.sheet(reference?)", "Get a worksheet."],
  ["workbook.addSheet(name, data?)", "Add a worksheet."],
  ["workbook.save(path, options?)", "Save an XLSX file."],
  ["sheet.setData(data, options?)", "Write rows or object records."],
  ["sheet.get(address)", "Read a cell value."],
  ["sheet.set(address, value)", "Set a cell value."],
  ["sheet.cell(address)", "Get a Cell object."],
  ["sheet.range(address)", "Get a rectangular Range object."],
  ["sheet.row(row)", "Select one row."],
  ["sheet.rows(selector)", "Select multiple rows."],
  ["sheet.column(column)", "Select one column."],
  ["sheet.columns(selector)", "Select multiple columns."],
  ["sheet.find(matcher)", "Find the first matching cell."],
  ["sheet.autoFit(options?)", "Fit column widths to their values."],
  ["sheet.merge(range)", "Merge a cell range."],
  ["sheet.toRecords(options?)", "Read rows as JavaScript objects."],
  ["sheet.toHtml(options?)", "Create an HTML table."],
  ["cell.style(style, mode?)", "Apply formatting to a cell."],
  ["cell.formula(formula, result?)", "Store a formula and cached result."],
  ["range.style(style, mode?)", "Apply formatting to a range."]
];

const parameterInfo = {
  address: ["string · {r,c}", "Excel address such as A1 or D12."],
  callback: ["function", "Function called for each selected cell."],
  changes: ["object", "Properties to update."],
  column: ["string · number", "Column letter or index."],
  config: ["object", "Configuration for the new item."],
  count: ["number", "Number of rows or items."],
  data: ["any[][] · object[]", "Rows or JavaScript object records."],
  definitions: ["object", "Named style definitions."],
  format: ["string", "Excel number-format code."],
  formula: ["string", "Excel formula without the leading equals sign."],
  height: ["number", "Row height in points."],
  matcher: ["any · function", "Value or callback used to find cells."],
  mode: ["merge · replace", "How formatting combines with the current style."],
  name: ["string", "Name for the workbook item."],
  options: ["object", "Optional settings for the method."],
  parts: ["string · string[]", "Style properties to clear."],
  password: ["string", "Password required to open the workbook."],
  path: ["string", "Local XLSX or XLSM file path."],
  records: ["object[]", "Objects whose keys map to columns."],
  range: ["string", "Excel range such as A1:D20."],
  reference: ["string · number", "Name, index, or ID of an existing item."],
  result: ["any", "Cached value shown before Excel recalculates."],
  row: ["number", "One-based worksheet row number."],
  rows: ["any[][]", "Two-dimensional row array."],
  selector: ["string · array", "Range such as A:C or 2:5, or selected references."],
  sheet: ["string · number", "Worksheet name or zero-based index."],
  source: ["value", "Workbook, cell, range, or style to read."],
  start: ["number", "First row affected by the operation."],
  style: ["string · object · array", "Named style or style object."],
  target: ["value", "Destination cell, row, or URL."],
  tooltip: ["string", "Optional hyperlink hover text."],
  value: ["any", "JavaScript value stored in the cell."],
  width: ["number", "Excel column width."]
};

function makeApiEntry(signature, description) {
  const open = signature.indexOf("(");
  return { signature, description, name: (open === -1 ? signature : signature.slice(0, open)).trim() };
}

let apiCatalog = fallbackApis.map(([signature, description]) => makeApiEntry(signature, description));

async function loadApiCatalog() {
  try {
    const response = await fetch("../index.html");
    if (!response.ok) return;
    const documentCopy = new DOMParser().parseFromString(await response.text(), "text/html");
    const entries = [...documentCopy.querySelectorAll("#cheatsheet .api-table tbody tr")].map((row) => {
      const signature = row.querySelector("code")?.textContent.trim();
      const description = row.querySelector("td:nth-child(2)")?.textContent.trim();
      return signature && description ? makeApiEntry(signature, description) : undefined;
    }).filter(Boolean);
    if (entries.length) apiCatalog = entries;
  } catch {
    // The built-in core catalog remains available when the reference cannot load.
  }
}

function getCompletionContext() {
  if (editor.selectionStart !== editor.selectionEnd) return undefined;
  const before = editor.value.slice(0, editor.selectionStart);
  const match = before.match(/([A-Za-z_$][\w$]*(?:\.[A-Za-z_$][\w$]*)*\.?)$/);
  if (!match) return undefined;
  return { text: match[1], start: editor.selectionStart - match[1].length, end: editor.selectionStart };
}

function measureEditor() {
  const style = getComputedStyle(editor);
  const canvas = measureEditor.canvas ??= document.createElement("canvas");
  const context = canvas.getContext("2d");
  context.font = style.font;
  return {
    charWidth: context.measureText("M").width,
    lineHeight: Number.parseFloat(style.lineHeight),
    paddingLeft: Number.parseFloat(style.paddingLeft),
    paddingTop: Number.parseFloat(style.paddingTop)
  };
}

function positionCompletions() {
  if (!completionContext) return;
  const metrics = measureEditor();
  const before = editor.value.slice(0, completionContext.end);
  const lines = before.split("\n");
  const left = metrics.paddingLeft + (lines.at(-1).length * metrics.charWidth) - editor.scrollLeft;
  const top = metrics.paddingTop + (lines.length * metrics.lineHeight) - editor.scrollTop;
  completionMenu.style.left = `${Math.max(10, Math.min(left, editor.clientWidth - 250))}px`;
  completionMenu.style.top = `${Math.max(10, Math.min(top, editor.clientHeight - 160))}px`;
}

function updateCompletionSelection() {
  let selectedButton;
  completionMenu.querySelectorAll("button").forEach((button, index) => {
    const selected = index === completionIndex;
    button.setAttribute("aria-selected", String(selected));
    if (selected) {
      selectedButton = button;
      editor.setAttribute("aria-activedescendant", button.id);
      button.scrollIntoView({ block: "nearest" });
    }
  });
  if (selectedButton) {
    const bounds = selectedButton.getBoundingClientRect();
    showApiTooltip(completionMatches[completionIndex], bounds.right, bounds.top);
  }
}

function hideCompletions() {
  completionMenu.hidden = true;
  completionMenu.replaceChildren();
  completionMatches = [];
  completionContext = undefined;
  editor.removeAttribute("aria-activedescendant");
  hideApiTooltip();
}

function acceptCompletion(index = completionIndex) {
  const entry = completionMatches[index];
  if (!entry || !completionContext) return;
  const callable = entry.signature.includes("(");
  const replacement = `${entry.name}${callable ? "()" : ""}`;
  editor.setRangeText(replacement, completionContext.start, completionContext.end, "end");
  if (callable) editor.setSelectionRange(editor.selectionStart - 1, editor.selectionStart - 1);
  hideCompletions();
  editor.focus();
}

function showCompletions() {
  completionContext = getCompletionContext();
  const query = completionContext?.text.toLowerCase();
  if (!query || (!query.includes(".") && query.length < 2)) {
    hideCompletions();
    return;
  }
  const unique = new Map();
  for (const entry of apiCatalog) {
    const fullName = entry.name.toLowerCase();
    const shortName = fullName.split(".").at(-1);
    if ((fullName.startsWith(query) || (!query.includes(".") && shortName.startsWith(query))) && !unique.has(entry.name)) {
      unique.set(entry.name, entry);
    }
  }
  completionMatches = [...unique.values()].slice(0, 8);
  if (!completionMatches.length) {
    hideCompletions();
    return;
  }
  completionIndex = 0;
  completionMenu.replaceChildren(...completionMatches.map((entry, index) => {
    const button = document.createElement("button");
    button.type = "button";
    button.id = `editor-completion-${index}`;
    button.className = "editor-completion";
    button.role = "option";
    const signature = document.createElement("code");
    signature.textContent = entry.signature;
    const kind = document.createElement("small");
    kind.textContent = "API";
    const detail = document.createElement("span");
    detail.textContent = entry.description;
    button.append(signature, kind, detail);
    button.addEventListener("mousedown", (event) => event.preventDefault());
    button.addEventListener("pointerenter", () => {
      completionIndex = index;
      updateCompletionSelection();
    });
    button.addEventListener("click", () => acceptCompletion(index));
    return button;
  }));
  completionMenu.hidden = false;
  positionCompletions();
  updateCompletionSelection();
}

function parametersFor(signature) {
  const open = signature.indexOf("(");
  const close = signature.lastIndexOf(")");
  if (open === -1 || close <= open + 1) return [];
  const names = [...signature.slice(open + 1, close).matchAll(/[A-Za-z_$][\w$]*/g)]
    .map((match) => match[0])
    .filter((name, index, all) => all.indexOf(name) === index);
  return names.map((name) => ({
    name,
    info: parameterInfo[name] ?? ["value", "Method parameter."],
    schema: parameterSchemas[signature]?.[name]
  }));
}

function hideApiTooltip() {
  hoveredApiName = "";
  apiTooltip.hidden = true;
  apiTooltip.replaceChildren();
}

function showApiTooltip(entry, clientX, clientY) {
  hoveredApiName = entry.name;
  const signature = document.createElement("code");
  signature.textContent = entry.signature;
  const description = document.createElement("p");
  description.textContent = entry.description;
  apiTooltip.replaceChildren(signature, description);
  const parameters = parametersFor(entry.signature);
  if (parameters.length) {
    const list = document.createElement("dl");
    for (const parameter of parameters) {
      const term = document.createElement("dt");
      term.textContent = `${parameter.name}: ${parameter.info[0]}`;
      const detail = document.createElement("dd");
      detail.textContent = parameter.info[1];
      list.append(term, detail);
    }
    apiTooltip.append(list);
    for (const parameter of parameters.filter(({ schema }) => schema)) {
      const schema = document.createElement("section");
      schema.className = "editor-api-schema";
      const title = document.createElement("strong");
      title.textContent = parameter.schema.title;
      const properties = document.createElement("div");
      properties.className = "editor-api-properties";
      for (const [name, type, defaultValue, detail] of parameter.schema.properties) {
        const property = document.createElement("div");
        const heading = document.createElement("code");
        heading.textContent = name;
        const meta = document.createElement("small");
        meta.textContent = `${type}${defaultValue ? ` · default ${defaultValue}` : ""}`;
        const copy = document.createElement("span");
        copy.textContent = detail;
        property.append(heading, meta, copy);
        properties.append(property);
      }
      schema.append(title, properties);
      if (parameter.schema.example) {
        const example = document.createElement("p");
        example.className = "editor-api-example";
        const label = document.createElement("span");
        label.textContent = "Example";
        const code = document.createElement("code");
        code.textContent = parameter.schema.example;
        example.append(label, code);
        schema.append(example);
      }
      apiTooltip.append(schema);
    }
  }
  apiTooltip.hidden = false;
  const bounds = apiTooltip.getBoundingClientRect();
  const left = Math.min(window.innerWidth - bounds.width - 12, clientX + 14);
  const top = Math.min(window.innerHeight - bounds.height - 12, clientY + 16);
  apiTooltip.style.left = `${Math.max(12, left)}px`;
  apiTooltip.style.top = `${Math.max(12, top)}px`;
}

function apiAtPointer(event) {
  const bounds = editor.getBoundingClientRect();
  const metrics = measureEditor();
  const x = event.clientX - bounds.left - metrics.paddingLeft + editor.scrollLeft;
  const y = event.clientY - bounds.top - metrics.paddingTop + editor.scrollTop;
  if (x < 0 || y < 0) return undefined;
  const line = editor.value.split("\n")[Math.floor(y / metrics.lineHeight)];
  if (line === undefined) return undefined;
  const column = Math.floor(x / metrics.charWidth);
  const token = [...line.matchAll(/[A-Za-z_$][\w$]*(?:\.[A-Za-z_$][\w$]*)*/g)]
    .find((match) => column >= match.index && column <= match.index + match[0].length)?.[0];
  if (!token) return undefined;
  return apiCatalog.find((entry) => entry.name === token)
    ?? (token.length > 2 ? apiCatalog.find((entry) => entry.name.endsWith(`.${token}`)) : undefined);
}

function setStatus(message, kind = "") {
  status.textContent = message;
  status.className = `run-status ${kind}`.trim();
}

function formatValue(cell) {
  let value = cell.formula ? (cell.result ?? `=${cell.formula}`) : cell.value;
  const format = cell.style?.numberFormat ?? "";
  if (typeof value === "number" && format.includes("%")) return `${(value * 100).toFixed(format.includes("0.0") ? 1 : 0)}%`;
  if (typeof value === "number" && format.includes("$")) return new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" }).format(value);
  if (value instanceof Date) return value.toLocaleDateString();
  return value ?? "";
}

function applyStyle(element, style = {}) {
  const font = style.font ?? {};
  if (style.bold ?? font.bold) element.style.fontWeight = "700";
  if (style.italic ?? font.italic) element.style.fontStyle = "italic";
  if (style.fontName ?? font.name) element.style.fontFamily = style.fontName ?? font.name;
  if (style.fontSize ?? font.size) element.style.fontSize = `${style.fontSize ?? font.size}px`;
  const color = style.color ?? font.color;
  if (color) element.style.color = typeof color === "string" ? color : `#${color.rgb}`;
  if (style.fill) {
    const fill = typeof style.fill === "string" ? style.fill : style.fill.foreground?.rgb ?? style.fill.color;
    if (fill) element.style.backgroundColor = fill.startsWith?.("#") ? fill : `#${fill}`;
  }
  const alignment = style.alignment ?? {};
  if (style.horizontal ?? alignment.horizontal) element.style.textAlign = style.horizontal ?? alignment.horizontal;
  if (style.vertical ?? alignment.vertical) element.style.verticalAlign = style.vertical ?? alignment.vertical;
  if (style.wrapText ?? alignment.wrapText) element.style.whiteSpace = "normal";
}

function columnName(index) {
  let name = "";
  for (let value = index + 1; value > 0; value = Math.floor((value - 1) / 26)) {
    name = String.fromCharCode(65 + ((value - 1) % 26)) + name;
  }
  return name;
}

function renderSheet(index) {
  activeSheet = index;
  const sheet = result.workbook.sheets[index];
  tabs.querySelectorAll("button").forEach((button, buttonIndex) => button.setAttribute("aria-selected", String(buttonIndex === index)));
  preview.replaceChildren();
  const maxDataRow = sheet.cells.length ? Math.max(...sheet.cells.map(({ r }) => r)) : -1;
  const maxDataColumn = sheet.cells.length ? Math.max(...sheet.cells.map(({ c }) => c)) : -1;
  const rowCount = Math.max(maxDataRow + 1, 18);
  const columnCount = Math.max(maxDataColumn + 1, 8);
  const cells = new Map(sheet.cells.map((cell) => [`${cell.r}:${cell.c}`, cell]));
  const merged = new Map();
  for (const range of sheet.merges) {
    merged.set(`${range.s.r}:${range.s.c}`, { rowSpan: range.e.r - range.s.r + 1, colSpan: range.e.c - range.s.c + 1 });
    for (let r = range.s.r; r <= range.e.r; r += 1) for (let c = range.s.c; c <= range.e.c; c += 1) {
      if (r !== range.s.r || c !== range.s.c) merged.set(`${r}:${c}`, { covered: true });
    }
  }

  const frame = document.createElement("div");
  frame.className = "sheet-frame";
  const table = document.createElement("table");
  table.className = "sheet-table";
  const colgroup = document.createElement("colgroup");
  const rowMarkerColumn = document.createElement("col");
  rowMarkerColumn.className = "row-marker-column";
  colgroup.append(rowMarkerColumn);
  for (let c = 0; c < columnCount; c += 1) {
    const col = document.createElement("col");
    if (sheet.columnWidths[c]) col.style.width = `${Math.max(56, sheet.columnWidths[c] * 7)}px`;
    colgroup.append(col);
  }
  table.append(colgroup);
  const thead = document.createElement("thead");
  const markerRow = document.createElement("tr");
  const corner = document.createElement("th");
  corner.className = "sheet-corner";
  corner.setAttribute("aria-label", "Select all cells");
  markerRow.append(corner);
  for (let c = 0; c < columnCount; c += 1) {
    const marker = document.createElement("th");
    marker.className = "column-marker";
    marker.scope = "col";
    marker.textContent = columnName(c);
    markerRow.append(marker);
  }
  thead.append(markerRow);
  table.append(thead);
  const tbody = document.createElement("tbody");
  for (let r = 0; r < rowCount; r += 1) {
    const row = document.createElement("tr");
    if (sheet.rowHeights[r]) row.style.height = `${sheet.rowHeights[r]}px`;
    const rowMarker = document.createElement("th");
    rowMarker.className = "row-marker";
    rowMarker.scope = "row";
    rowMarker.textContent = String(r + 1);
    row.append(rowMarker);
    for (let c = 0; c < columnCount; c += 1) {
      const merge = merged.get(`${r}:${c}`);
      if (merge?.covered) continue;
      const element = document.createElement("td");
      const cell = cells.get(`${r}:${c}`);
      element.textContent = cell ? formatValue(cell) : "";
      element.title = cell?.formula ? `=${cell.formula}` : cell?.address ?? "";
      if (merge?.rowSpan > 1) element.rowSpan = merge.rowSpan;
      if (merge?.colSpan > 1) element.colSpan = merge.colSpan;
      if (cell) applyStyle(element, cell.style);
      row.append(element);
    }
    tbody.append(row);
  }
  table.append(tbody);
  frame.append(table);
  preview.append(frame);
}

function renderResult() {
  tabs.replaceChildren();
  result.workbook.sheets.forEach((sheet, index) => {
    const button = document.createElement("button");
    button.type = "button";
    button.role = "tab";
    button.textContent = sheet.name;
    button.setAttribute("aria-selected", String(index === 0));
    button.addEventListener("click", () => renderSheet(index));
    tabs.append(button);
  });
  renderSheet(Math.min(activeSheet, result.workbook.sheets.length - 1));
  consoleOutput.textContent = result.logs.length
    ? result.logs.map(({ level, text }) => `[${level}] ${text}`).join("\n")
    : "No console output.";
  downloadButton.disabled = false;
}

function runCode() {
  runButton.disabled = true;
  downloadButton.disabled = true;
  setStatus("Running…");
  consoleOutput.textContent = "Running…";
  const worker = new Worker("playground-worker.js", { type: "module" });
  const timeout = window.setTimeout(() => {
    worker.terminate();
    runButton.disabled = false;
    setStatus("Stopped after 5 seconds", "error");
    consoleOutput.textContent = "The example took too long and was stopped. Check for an infinite loop.";
  }, 5000);
  worker.addEventListener("message", ({ data }) => {
    window.clearTimeout(timeout);
    worker.terminate();
    runButton.disabled = false;
    if (!data.ok) {
      result = undefined;
      setStatus("Fix the error and run again", "error");
      consoleOutput.textContent = data.error;
      preview.innerHTML = `<div class="preview-empty">The workbook preview will appear after the code runs successfully.</div>`;
      tabs.replaceChildren();
      return;
    }
    result = data;
    setStatus("Workbook ready", "success");
    renderResult();
  });
  worker.addEventListener("error", (event) => {
    window.clearTimeout(timeout);
    worker.terminate();
    runButton.disabled = false;
    setStatus("Playground error", "error");
    consoleOutput.textContent = event.message;
  });
  worker.postMessage({ source: editor.value });
}

function setSample(name) {
  editor.value = samples[name];
  hideCompletions();
  hideApiTooltip();
  result = undefined;
  downloadButton.disabled = true;
  setStatus("Ready");
}

sample.addEventListener("change", () => { setSample(sample.value); runCode(); });
resetButton.addEventListener("click", () => { setSample(sample.value); runCode(); });
runButton.addEventListener("click", runCode);
editor.addEventListener("input", showCompletions);
editor.addEventListener("click", hideCompletions);
editor.addEventListener("blur", () => window.setTimeout(hideCompletions, 0));
editor.addEventListener("scroll", () => { hideCompletions(); hideApiTooltip(); });
editor.addEventListener("mousemove", (event) => {
  const entry = apiAtPointer(event);
  if (!entry) {
    hideApiTooltip();
    return;
  }
  if (entry.name !== hoveredApiName) showApiTooltip(entry, event.clientX, event.clientY);
});
editor.addEventListener("mouseleave", hideApiTooltip);
editor.addEventListener("keydown", (event) => {
  if ((event.ctrlKey || event.metaKey) && event.key === "Enter") {
    event.preventDefault();
    hideCompletions();
    runCode();
    return;
  }
  if (!completionMenu.hidden && ["ArrowDown", "ArrowUp", "Enter", "Tab", "Escape"].includes(event.key)) {
    event.preventDefault();
    if (event.key === "ArrowDown") completionIndex = (completionIndex + 1) % completionMatches.length;
    if (event.key === "ArrowUp") completionIndex = (completionIndex - 1 + completionMatches.length) % completionMatches.length;
    if (event.key === "Enter" || event.key === "Tab") acceptCompletion();
    if (event.key === "Escape") hideCompletions();
    if (!completionMenu.hidden) updateCompletionSelection();
    return;
  }
  if (event.key === "Tab") {
    event.preventDefault();
    const start = editor.selectionStart;
    editor.setRangeText("  ", start, editor.selectionEnd, "end");
  }
});
downloadButton.addEventListener("click", () => {
  if (!result) return;
  const blob = new Blob([result.bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = result.workbook.name || "playground.xlsx";
  link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
});

setSample("data");
runCode();
loadApiCatalog();
