import { decodeAddress, decodeColumn, decodeRange, encodeAddress, encodeRange, normalizeRange } from "./address.js";
import { Cell } from "./cell.js";
import { Column } from "./column.js";
import { Range } from "./range.js";
import { escapeXml } from "./xml.js";
import { legacyPasswordHash } from "./protection.js";
import { normalizeStyle } from "./style.js";

function cellAddresses(source) {
  return Object.keys(source).filter((key) => /^[A-Z]+[1-9]\d*$/.test(key));
}

function cloneCell(cell) {
  return structuredClone(cell);
}

function rowShape(data) {
  if (!Array.isArray(data)) throw new TypeError("Rows must be an array.");
  if (!data.length) return undefined;
  const shape = Array.isArray(data[0]) ? "array" : "object";
  const valid = shape === "array"
    ? data.every(Array.isArray)
    : data.every((row) => row !== null && typeof row === "object" && !Array.isArray(row));
  if (!valid) throw new TypeError("Every row must be an array or every row must be an object.");
  return shape;
}

function shiftedFormula(formula, startRow, delta, deletedEnd) {
  if (!formula) return formula;
  return formula.replace(/(\$?[A-Z]{1,3})(\$?)(\d+)/g, (match, column, absolute, rowText) => {
    const row = Number(rowText);
    if (deletedEnd !== undefined && row >= startRow && row <= deletedEnd) return "#REF!";
    if (row < startRow) return match;
    return `${column}${absolute}${Math.max(1, row + delta)}`;
  });
}

function recomputeReference(source) {
  const addresses = cellAddresses(source);
  if (!addresses.length) { delete source["!ref"]; return; }
  const points = addresses.map(decodeAddress);
  source["!ref"] = encodeRange({
    s: { r: Math.min(...points.map((point) => point.r)), c: Math.min(...points.map((point) => point.c)) },
    e: { r: Math.max(...points.map((point) => point.r)), c: Math.max(...points.map((point) => point.c)) }
  });
}

function csvValue(value, delimiter) {
  if (value === null || value === undefined) return "";
  const text = value instanceof Date ? value.toISOString() : String(value);
  return text.includes(delimiter) || /["\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

export class Worksheet {
  constructor(source, getName, onChange = () => {}, onStructureChange = () => {}, resolveStyle = normalizeStyle) {
    this.source = source;
    this.getName = getName;
    this.onChange = onChange;
    this.onStructureChange = onStructureChange;
    this.resolveStyle = resolveStyle;
  }

  get name() {
    return this.getName();
  }

  get unsafeRaw() {
    return this.source;
  }

  get usedRange() {
    return this.source["!ref"];
  }

  cell(address) {
    return new Cell(this.source, address, this.onChange, this.resolveStyle);
  }

  get(address) {
    return this.cell(address).value;
  }

  set(address, value) {
    this.cell(address).value = value;
    return this;
  }

  range(address) {
    return new Range(this.source, address, this.onChange, this.onStructureChange, this.resolveStyle);
  }

  column(column) {
    return new Column(this.source, column, this.onChange, this.onStructureChange, this.resolveStyle);
  }

  find(matcher) {
    return this.findAll(matcher)[0];
  }

  findAll(matcher) {
    const found = [];
    const addresses = cellAddresses(this.source).sort((left, right) => {
      const a = decodeAddress(left);
      const b = decodeAddress(right);
      return a.r - b.r || a.c - b.c;
    });
    for (const address of addresses) {
      const cell = this.cell(address);
      if (cell.value === undefined && !cell.unsafeRaw?.formula) continue;
      const match = typeof matcher === "function" ? matcher(cell) : cell.value === matcher;
      if (match) found.push(cell);
    }
    return found;
  }

  appendRows(data, options = {}) {
    const shape = rowShape(data);
    if (!data.length) return this;
    const existing = this.usedRange ? decodeRange(this.usedRange) : undefined;
    const origin = options.origin === undefined || options.origin === -1
      ? { r: existing ? existing.e.r + 1 : 0, c: 0 }
      : typeof options.origin === "string"
        ? this.cell(options.origin).coordinates
        : typeof options.origin === "number"
          ? { r: options.origin, c: 0 }
          : options.origin;
    const objectRows = shape === "object";
    let rows = data;
    if (objectRows) {
      const headers = options.header ?? [...new Set(data.flatMap((row) => Object.keys(row)))];
      rows = [...(options.skipHeader ? [] : [headers]), ...data.map((row) => headers.map((header) => row[header]))];
    }
    rows.forEach((row, rowOffset) => {
      row.forEach((value, columnOffset) => this.cell({ r: origin.r + rowOffset, c: origin.c + columnOffset }).set(value));
    });
    return this;
  }

  insertRows(beforeRow, count = 1, options = {}) {
    if (!Number.isInteger(beforeRow) || beforeRow < 1) throw new RangeError("Row number must be a positive integer.");
    if (!Number.isInteger(count) || count < 1) throw new RangeError("Row count must be a positive integer.");
    const start = beforeRow - 1;
    const moving = cellAddresses(this.source)
      .map((address) => ({ address, point: decodeAddress(address), cell: this.source[address] }))
      .filter((item) => item.point.r >= start)
      .sort((left, right) => right.point.r - left.point.r || right.point.c - left.point.c);
    for (const item of moving) {
      const target = encodeAddress({ r: item.point.r + count, c: item.point.c });
      delete this.source[item.address];
      this.onChange(item.address, "clear");
      item.cell.formula = shiftedFormula(item.cell.formula, beforeRow, count);
      this.source[target] = item.cell;
      this.onChange(target, item.cell.formula ? "formula" : "value");
    }
    for (const address of cellAddresses(this.source)) {
      const cell = this.source[address];
      if (moving.some((item) => encodeAddress({ r: item.point.r + count, c: item.point.c }) === address)) continue;
      const formula = shiftedFormula(cell.formula, beforeRow, count);
      if (formula !== cell.formula) { cell.formula = formula; this.onChange(address, "formula"); }
    }
    const rows = (this.source["!rows"] ??= []);
    rows.splice(start, 0, ...Array(count));
    this.source["!merges"] = (this.source["!merges"] ?? []).map((range) => {
      const copy = structuredClone(range);
      if (copy.s.r >= start) { copy.s.r += count; copy.e.r += count; }
      else if (copy.e.r >= start) copy.e.r += count;
      return copy;
    });
    if (options.copyFrom !== undefined) {
      const sourceRow = options.copyFrom === "above" ? beforeRow - 1 : options.copyFrom === "below" ? beforeRow + count : Number(options.copyFrom);
      for (let offset = 0; offset < count; offset += 1) this.copyRow(sourceRow, beforeRow + offset, { values: options.values !== false });
    }
    recomputeReference(this.source);
    this.onStructureChange();
    return this;
  }

  deleteRows(startRow, count = 1) {
    if (!Number.isInteger(startRow) || startRow < 1) throw new RangeError("Row number must be a positive integer.");
    if (!Number.isInteger(count) || count < 1) throw new RangeError("Row count must be a positive integer.");
    const start = startRow - 1;
    const end = start + count - 1;
    const items = cellAddresses(this.source)
      .map((address) => ({ address, point: decodeAddress(address), cell: this.source[address] }))
      .sort((left, right) => left.point.r - right.point.r || left.point.c - right.point.c);
    for (const item of items) {
      if (item.point.r < start) continue;
      delete this.source[item.address];
      this.onChange(item.address, "clear");
      if (item.point.r <= end) continue;
      const target = encodeAddress({ r: item.point.r - count, c: item.point.c });
      item.cell.formula = shiftedFormula(item.cell.formula, startRow, -count, startRow + count - 1);
      this.source[target] = item.cell;
      this.onChange(target, item.cell.formula ? "formula" : "value");
    }
    for (const address of cellAddresses(this.source)) {
      const cell = this.source[address];
      const formula = shiftedFormula(cell.formula, startRow, -count, startRow + count - 1);
      if (formula !== cell.formula) { cell.formula = formula; this.onChange(address, "formula"); }
    }
    this.source["!rows"]?.splice(start, count);
    this.source["!merges"] = (this.source["!merges"] ?? []).flatMap((range) => {
      if (range.s.r >= start && range.e.r <= end) return [];
      const copy = structuredClone(range);
      if (copy.s.r > end) { copy.s.r -= count; copy.e.r -= count; }
      else if (copy.e.r > end) copy.e.r -= count;
      else if (copy.e.r >= start) copy.e.r = Math.max(copy.s.r, start - 1);
      return [copy];
    });
    recomputeReference(this.source);
    this.onStructureChange();
    return this;
  }

  copyRow(sourceRow, targetRow, options = {}) {
    if (!Number.isInteger(sourceRow) || sourceRow < 1 || !Number.isInteger(targetRow) || targetRow < 1) throw new RangeError("Source and target rows must be positive integers.");
    if (sourceRow === targetRow) return this;
    const sourceIndex = sourceRow - 1;
    const targetIndex = targetRow - 1;
    for (const address of cellAddresses(this.source).filter((key) => decodeAddress(key).r === targetIndex)) {
      delete this.source[address];
      this.onChange(address, "clear");
    }
    for (const address of cellAddresses(this.source).filter((key) => decodeAddress(key).r === sourceIndex)) {
      const point = decodeAddress(address);
      const target = encodeAddress({ r: targetIndex, c: point.c });
      const cell = cloneCell(this.source[address]);
      if (options.values === false) { cell.type = "blank"; cell.value = undefined; delete cell.formula; }
      else if (cell.formula) cell.formula = shiftedFormula(cell.formula, 1, targetRow - sourceRow);
      this.source[target] = cell;
      this.onChange(target, cell.formula ? "formula" : "value");
    }
    if (this.source["!rows"]?.[sourceIndex]) {
      (this.source["!rows"] ??= [])[targetIndex] = structuredClone(this.source["!rows"][sourceIndex]);
    }
    recomputeReference(this.source);
    this.onStructureChange();
    return this;
  }

  appendData(records, options = {}) {
    if (!Array.isArray(records)) throw new TypeError("Records must be an array.");
    if (!records.length) return this;
    if (records.some((record) => record === null || typeof record !== "object" || Array.isArray(record))) {
      throw new TypeError("Every record must be an object.");
    }

    if (!this.usedRange) {
      return this.appendRows(records, {
        ...options,
        origin: options.origin ?? "A1",
        skipHeader: false
      });
    }

    const used = decodeRange(this.usedRange);
    const headers = options.header ?? Array.from(
      { length: used.e.c - used.s.c + 1 },
      (_, offset) => this.cell({ r: used.s.r, c: used.s.c + offset }).value
    ).map((value, index) => String(value ?? `Column${index + 1}`));

    return this.appendRows(records, {
      ...options,
      origin: options.origin ?? { r: used.e.r + 1, c: used.s.c },
      header: headers,
      skipHeader: true
    });
  }

  setData(data, options = {}) {
    rowShape(data);
    for (const key of cellAddresses(this.source)) {
      delete this.source[key];
      this.onChange(key, "clear");
    }
    delete this.source["!ref"];
    this.onStructureChange();
    return this.appendRows(data, { ...options, origin: "A1", skipHeader: options.skipHeader ?? false });
  }

  toRows(options = {}) {
    if (!this.usedRange) return [];
    const range = normalizeRange(options.range ?? this.usedRange);
    const rows = [];
    for (let row = range.s.r; row <= range.e.r; row += 1) {
      const values = [];
      for (let column = range.s.c; column <= range.e.c; column += 1) {
        const value = this.cell({ r: row, c: column }).value;
        values.push(value === undefined && "defaultValue" in options ? options.defaultValue : value);
      }
      rows.push(values);
    }
    return rows;
  }

  toRecords(options = {}) {
    const rows = this.toRows(options);
    if (!rows.length) return [];
    const headers = options.headers ?? rows[0].map((value, index) => String(value ?? `Column${index + 1}`));
    const start = options.headers ? 0 : 1;
    return rows.slice(start).map((row) => Object.fromEntries(headers.map((header, index) => [header, row[index] ?? options.defaultValue ?? null])));
  }

  toCsv(options = {}) {
    const delimiter = options.delimiter ?? ",";
    const newline = options.newline ?? "\n";
    return this.toRows(options).map((row) => row.map((value) => csvValue(value, delimiter)).join(delimiter)).join(newline);
  }

  toHtml(options = {}) {
    const rows = this.toRows(options);
    const body = rows.map((row, rowIndex) => `<tr>${row.map((value) => {
      const tag = rowIndex === 0 && options.header !== false ? "th" : "td";
      return `<${tag}>${escapeXml(value ?? "")}</${tag}>`;
    }).join("")}</tr>`).join("");
    return `<table>${body}</table>`;
  }

  merge(range) {
    this.range(range).merge();
    return this;
  }

  unmerge(range) {
    this.range(range).unmerge();
    return this;
  }

  setColumnWidth(column, width) {
    if (!Number.isFinite(width) || width <= 0) throw new RangeError("Column width must be a positive number.");
    const index = typeof column === "number" ? column : decodeColumn(column);
    if (!Number.isInteger(index) || index < 0) throw new RangeError("Column index must be a non-negative integer.");
    const columns = (this.source["!cols"] ??= []);
    columns[index] = { ...(columns[index] ?? {}), width };
    this.onStructureChange();
    return this;
  }

  setRowHeight(row, height) {
    if (!Number.isInteger(row) || row < 1) throw new RangeError("Row number must be a positive integer.");
    if (!Number.isFinite(height) || height <= 0) throw new RangeError("Row height must be a positive number.");
    const rows = (this.source["!rows"] ??= []);
    rows[row - 1] = { ...(rows[row - 1] ?? {}), height };
    this.onStructureChange();
    return this;
  }

  autoFilter(range = this.usedRange) {
    if (!range) return this;
    this.source["!autofilter"] = encodeRange(normalizeRange(range));
    this.onStructureChange();
    return this;
  }

  autoFit(options = {}) {
    if (!this.usedRange) return this;
    const range = normalizeRange(this.usedRange);
    const min = options.min ?? 8;
    const max = options.max ?? 60;
    const padding = options.padding ?? 2;
    for (let column = range.s.c; column <= range.e.c; column += 1) {
      let width = min;
      const startRow = options.includeHeader === false ? range.s.r + 1 : range.s.r;
      for (let row = startRow; row <= range.e.r; row += 1) {
        const value = this.source[encodeAddress({ r: row, c: column })]?.value;
        if (value === null || value === undefined) continue;
        const text = value instanceof Date ? value.toISOString() : String(value);
        width = Math.max(width, ...text.split(/\r?\n/).map((line) => line.length + padding));
      }
      this.setColumnWidth(column, Math.min(width, max));
    }
    return this;
  }

  protectSheet(options = {}) {
    const config = typeof options === "string" ? { password: options } : options;
    this.source["!protection"] = {
      sheet: true,
      objects: config.objects !== false,
      scenarios: config.scenarios !== false,
      ...(config.password ? { password: legacyPasswordHash(config.password) } : {}),
      ...(config.selectLockedCells === false ? { selectLockedCells: true } : {}),
      ...(config.selectUnlockedCells === false ? { selectUnlockedCells: true } : {}),
      ...(config.formatCells === false ? { formatCells: true } : {}),
      ...(config.formatColumns === false ? { formatColumns: true } : {}),
      ...(config.formatRows === false ? { formatRows: true } : {}),
      ...(config.insertColumns === false ? { insertColumns: true } : {}),
      ...(config.insertRows === false ? { insertRows: true } : {}),
      ...(config.deleteColumns === false ? { deleteColumns: true } : {}),
      ...(config.deleteRows === false ? { deleteRows: true } : {}),
      ...(config.sort === false ? { sort: true } : {}),
      ...(config.autoFilter === false ? { autoFilter: true } : {}),
      ...(config.pivotTables === false ? { pivotTables: true } : {})
    };
    this.onStructureChange();
    return this;
  }

  unprotectSheet() {
    delete this.source["!protection"];
    this.onStructureChange();
    return this;
  }
}
