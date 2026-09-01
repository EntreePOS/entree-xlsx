import { decodeAddress } from "./address.js";
import { Cell } from "./cell.js";
import { deepMerge, normalizeStyle, styleForCell } from "./style.js";

function matches(cell, matcher) {
  return typeof matcher === "function" ? Boolean(matcher(cell)) : cell.value === matcher;
}

function validateRow(row) {
  if (!Number.isInteger(row) || row < 1) throw new RangeError("Row number must be a positive integer.");
  return row;
}

export class Row {
  constructor(worksheet, row, onChange = () => {}, onStructureChange = () => {}, resolveStyle = normalizeStyle) {
    this.worksheet = worksheet;
    this.number = validateRow(row);
    this.index = this.number - 1;
    this.onChange = onChange;
    this.onStructureChange = onStructureChange;
    this.resolveStyle = resolveStyle;
  }

  find(matcher) {
    return this.findAll(matcher)[0];
  }

  findAll(matcher) {
    const found = [];
    this.forEach((cell) => {
      if (matches(cell, matcher)) found.push(cell);
    });
    return found;
  }

  forEach(callback) {
    const addresses = Object.keys(this.worksheet)
      .filter((address) => /^[A-Z]+[1-9]\d*$/.test(address))
      .filter((address) => decodeAddress(address).r === this.index)
      .filter((address) => {
        const cell = this.worksheet[address];
        return (cell?.value !== undefined && cell?.value !== null) || cell?.formula;
      })
      .sort((left, right) => decodeAddress(left).c - decodeAddress(right).c);
    for (const address of addresses) callback(new Cell(this.worksheet, address, this.onChange, this.resolveStyle));
    return this;
  }

  style(style, mode = "merge") {
    if (!["merge", "replace"].includes(mode)) throw new TypeError('Style mode must be "merge" or "replace".');
    const resolved = styleForCell(this.resolveStyle(style));
    const namedStyle = [...(Array.isArray(style) ? style : [style])].reverse().find((item) => typeof item === "string");
    const rows = (this.worksheet["!rows"] ??= []);
    const row = (rows[this.index] ??= {});
    row.style = mode === "replace" ? structuredClone(resolved) : deepMerge(row.style ?? {}, resolved);
    if (namedStyle) row.namedStyle = namedStyle;
    else if (mode === "replace") delete row.namedStyle;
    delete row.styleIndex;
    for (const address of Object.keys(this.worksheet).filter((key) => /^[A-Z]+[1-9]\d*$/.test(key) && decodeAddress(key).r === this.index)) {
      new Cell(this.worksheet, address, this.onChange, this.resolveStyle).applyStyle(resolved, mode, namedStyle);
    }
    this.onStructureChange();
    return this;
  }

  height(height) {
    if (!Number.isFinite(height) || height <= 0) throw new RangeError("Row height must be a positive number.");
    const rows = (this.worksheet["!rows"] ??= []);
    rows[this.index] = { ...(rows[this.index] ?? {}), height };
    this.onStructureChange();
    return this;
  }
}

export function normalizeRowSelector(selector) {
  const values = Array.isArray(selector)
    ? selector
    : typeof selector === "string"
      ? (() => {
          const match = /^([1-9]\d*)(?::([1-9]\d*))?$/.exec(selector.trim());
          if (!match) throw new TypeError(`Invalid row selector ${JSON.stringify(selector)}.`);
          const start = Number(match[1]);
          const end = Number(match[2] ?? match[1]);
          return Array.from({ length: Math.abs(end - start) + 1 }, (_, offset) => Math.min(start, end) + offset);
        })()
      : [selector];
  return [...new Set(values.map(validateRow))];
}

export class Rows {
  constructor(worksheet, selector, onChange = () => {}, onStructureChange = () => {}, resolveStyle = normalizeStyle) {
    this.worksheet = worksheet;
    this.numbers = normalizeRowSelector(selector);
    this.onChange = onChange;
    this.onStructureChange = onStructureChange;
    this.resolveStyle = resolveStyle;
  }

  find(matcher) {
    return this.findAll(matcher)[0];
  }

  findAll(matcher) {
    const found = [];
    this.forEach((cell) => {
      if (matches(cell, matcher)) found.push(cell);
    });
    return found;
  }

  forEach(callback) {
    const selected = new Set(this.numbers.map((number) => number - 1));
    const addresses = Object.keys(this.worksheet)
      .filter((address) => /^[A-Z]+[1-9]\d*$/.test(address))
      .filter((address) => selected.has(decodeAddress(address).r))
      .filter((address) => {
        const cell = this.worksheet[address];
        return (cell?.value !== undefined && cell?.value !== null) || cell?.formula;
      })
      .sort((left, right) => {
        const a = decodeAddress(left);
        const b = decodeAddress(right);
        return a.r - b.r || a.c - b.c;
      });
    for (const address of addresses) callback(new Cell(this.worksheet, address, this.onChange, this.resolveStyle));
    return this;
  }

  style(style, mode = "merge") {
    for (const number of this.numbers) new Row(this.worksheet, number, this.onChange, this.onStructureChange, this.resolveStyle).style(style, mode);
    return this;
  }

  height(height) {
    for (const number of this.numbers) new Row(this.worksheet, number, this.onChange, this.onStructureChange, this.resolveStyle).height(height);
    return this;
  }
}
