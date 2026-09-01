import { decodeAddress, decodeColumn, encodeColumn } from "./address.js";
import { Cell } from "./cell.js";
import { deepMerge, normalizeStyle, styleForCell } from "./style.js";

function matches(cell, matcher) {
  return typeof matcher === "function" ? Boolean(matcher(cell)) : cell.value === matcher;
}

export class Column {
  constructor(worksheet, column, onChange = () => {}, onStructureChange = () => {}, resolveStyle = normalizeStyle) {
    this.worksheet = worksheet;
    this.index = typeof column === "number" ? column : decodeColumn(column);
    if (!Number.isInteger(this.index) || this.index < 0) throw new RangeError("Column index must be a non-negative integer.");
    this.onChange = onChange;
    this.onStructureChange = onStructureChange;
    this.resolveStyle = resolveStyle;
  }

  get letter() {
    return encodeColumn(this.index);
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
      .filter((address) => decodeAddress(address).c === this.index)
      .filter((address) => {
        const cell = this.worksheet[address];
        return (cell?.value !== undefined && cell?.value !== null) || cell?.formula;
      })
      .sort((left, right) => decodeAddress(left).r - decodeAddress(right).r);
    for (const address of addresses) callback(new Cell(this.worksheet, address, this.onChange, this.resolveStyle));
    return this;
  }

  style(style, mode = "merge") {
    if (!["merge", "replace"].includes(mode)) throw new TypeError('Style mode must be "merge" or "replace".');
    const resolved = styleForCell(this.resolveStyle(style));
    const namedStyle = [...(Array.isArray(style) ? style : [style])].reverse().find((item) => typeof item === "string");
    const columns = (this.worksheet["!cols"] ??= []);
    const column = (columns[this.index] ??= {});
    column.style = mode === "replace" ? structuredClone(resolved) : deepMerge(column.style ?? {}, resolved);
    if (namedStyle) column.namedStyle = namedStyle;
    else if (mode === "replace") delete column.namedStyle;
    delete column.styleIndex;
    for (const address of Object.keys(this.worksheet).filter((key) => /^[A-Z]+[1-9]\d*$/.test(key) && decodeAddress(key).c === this.index)) {
      new Cell(this.worksheet, address, this.onChange, this.resolveStyle).applyStyle(resolved, mode, namedStyle);
    }
    this.onStructureChange();
    return this;
  }

  width(width) {
    if (!Number.isFinite(width) || width <= 0) throw new RangeError("Column width must be a positive number.");
    const columns = (this.worksheet["!cols"] ??= []);
    columns[this.index] = { ...(columns[this.index] ?? {}), width };
    this.onStructureChange();
    return this;
  }
}
