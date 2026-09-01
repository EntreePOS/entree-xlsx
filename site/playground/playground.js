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
let result;
let activeSheet = 0;

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
  result = undefined;
  downloadButton.disabled = true;
  setStatus("Ready");
}

sample.addEventListener("change", () => { setSample(sample.value); runCode(); });
resetButton.addEventListener("click", () => { setSample(sample.value); runCode(); });
runButton.addEventListener("click", runCode);
editor.addEventListener("keydown", (event) => {
  if ((event.ctrlKey || event.metaKey) && event.key === "Enter") { event.preventDefault(); runCode(); }
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
