import { encodeRange, forEachAddress, normalizeRange } from "./address.js";
import { Cell } from "./cell.js";
import { normalizeStyle, styleForRangeCell } from "./style.js";

function sameRange(left, right) {
  return left.s.r === right.s.r && left.s.c === right.s.c && left.e.r === right.e.r && left.e.c === right.e.c;
}

export class Range {
  constructor(worksheet, range, onChange = () => {}, onStructureChange = () => {}, resolveStyle = normalizeStyle) {
    this.worksheet = worksheet;
    this.coordinates = normalizeRange(range);
    this.onChange = onChange;
    this.onStructureChange = onStructureChange;
    this.resolveStyle = resolveStyle;
  }

  get address() {
    return encodeRange(this.coordinates);
  }

  getValues() {
    const rows = [];
    for (let row = this.coordinates.s.r; row <= this.coordinates.e.r; row += 1) {
      const values = [];
      for (let column = this.coordinates.s.c; column <= this.coordinates.e.c; column += 1) {
        values.push(new Cell(this.worksheet, { r: row, c: column }, this.onChange, this.resolveStyle).value);
      }
      rows.push(values);
    }
    return rows;
  }

  setValues(values) {
    if (!Array.isArray(values) || values.some((row) => !Array.isArray(row))) {
      throw new TypeError("Range values must be an array of rows.");
    }
    const rowCount = this.coordinates.e.r - this.coordinates.s.r + 1;
    const columnCount = this.coordinates.e.c - this.coordinates.s.c + 1;
    if (values.length > rowCount || values.some((row) => row.length > columnCount)) {
      throw new RangeError(`Data does not fit in ${this.address} (${rowCount} rows by ${columnCount} columns).`);
    }
    values.forEach((row, rowOffset) => {
      row.forEach((value, columnOffset) => {
        new Cell(this.worksheet, {
          r: this.coordinates.s.r + rowOffset,
          c: this.coordinates.s.c + columnOffset
        }, this.onChange, this.resolveStyle).value = value;
      });
    });
    return this;
  }

  style(style, mode = "merge") {
    const resolved = this.resolveStyle(style);
    const namedStyle = [...(Array.isArray(style) ? style : [style])].reverse().find((item) => typeof item === "string");
    this.forEach((cell) => cell.applyStyle(styleForRangeCell(resolved, cell.coordinates, this.coordinates), mode, namedStyle));
    return this;
  }

  copyStyleFrom(source, options = {}) {
    const sourceRange = source instanceof Range ? source : new Range(this.worksheet, source, this.onChange, this.onStructureChange, this.resolveStyle);
    const sourceRows = sourceRange.coordinates.e.r - sourceRange.coordinates.s.r + 1;
    const sourceColumns = sourceRange.coordinates.e.c - sourceRange.coordinates.s.c + 1;
    const targetRows = this.coordinates.e.r - this.coordinates.s.r + 1;
    const targetColumns = this.coordinates.e.c - this.coordinates.s.c + 1;
    if (!options.repeat && (sourceRows !== targetRows || sourceColumns !== targetColumns)) {
      throw new RangeError("Source and target ranges must have the same size. Pass { repeat: true } to tile the source styles.");
    }
    for (let row = 0; row < targetRows; row += 1) {
      for (let column = 0; column < targetColumns; column += 1) {
        const sourceCell = new Cell(this.worksheet, {
          r: sourceRange.coordinates.s.r + row % sourceRows,
          c: sourceRange.coordinates.s.c + column % sourceColumns
        }, this.onChange, this.resolveStyle);
        const targetCell = new Cell(this.worksheet, {
          r: this.coordinates.s.r + row,
          c: this.coordinates.s.c + column
        }, this.onChange, this.resolveStyle);
        targetCell.applyStyle(sourceCell.getStyle(), options.mode ?? "replace", sourceCell.unsafeRaw?.namedStyle);
      }
    }
    return this;
  }

  clearStyle(parts) {
    this.forEach((cell) => cell.clearStyle(parts));
    return this;
  }

  clear(options = {}) {
    this.forEach((cell) => cell.clear(options));
    return this;
  }

  merge() {
    const merges = (this.worksheet["!merges"] ??= []);
    if (!merges.some((item) => sameRange(item, this.coordinates))) merges.push(this.coordinates);
    this.onStructureChange();
    return this;
  }

  unmerge() {
    this.worksheet["!merges"] = (this.worksheet["!merges"] ?? []).filter((item) => !sameRange(item, this.coordinates));
    this.onStructureChange();
    return this;
  }

  forEach(callback) {
    forEachAddress(this.coordinates, (address) => callback(new Cell(this.worksheet, address, this.onChange, this.resolveStyle)));
    return this;
  }
}
