import { decodeAddress, decodeColumn, encodeColumn } from "./address.js";
import { Cell } from "./cell.js";
import { Column } from "./column.js";
import { normalizeStyle } from "./style.js";

function normalizeColumn(column) {
  const index = typeof column === "number" ? column : decodeColumn(column);
  if (!Number.isInteger(index) || index < 0) throw new RangeError("Column index must be a non-negative integer.");
  return index;
}

function matches(cell, matcher) {
  return typeof matcher === "function" ? Boolean(matcher(cell)) : cell.value === matcher;
}

export function normalizeColumnSelector(selector) {
  const values = Array.isArray(selector)
    ? selector
    : typeof selector === "string" && selector.includes(":")
      ? (() => {
          const match = /^([A-Z]+):([A-Z]+)$/i.exec(selector.trim());
          if (!match) throw new TypeError(`Invalid column selector ${JSON.stringify(selector)}.`);
          const start = decodeColumn(match[1]);
          const end = decodeColumn(match[2]);
          return Array.from({ length: Math.abs(end - start) + 1 }, (_, offset) => Math.min(start, end) + offset);
        })()
      : [selector];
  return [...new Set(values.map(normalizeColumn))];
}

export class Columns {
  constructor(worksheet, selector, onChange = () => {}, onStructureChange = () => {}, resolveStyle = normalizeStyle) {
    this.worksheet = worksheet;
    this.indexes = normalizeColumnSelector(selector);
    this.letters = this.indexes.map(encodeColumn);
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
    const selected = new Set(this.indexes);
    const addresses = Object.keys(this.worksheet)
      .filter((address) => /^[A-Z]+[1-9]\d*$/.test(address))
      .filter((address) => selected.has(decodeAddress(address).c))
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
    for (const index of this.indexes) new Column(this.worksheet, index, this.onChange, this.onStructureChange, this.resolveStyle).style(style, mode);
    return this;
  }

  width(width) {
    for (const index of this.indexes) new Column(this.worksheet, index, this.onChange, this.onStructureChange, this.resolveStyle).width(width);
    return this;
  }
}
