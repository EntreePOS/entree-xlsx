import { decodeAddress, expandWorksheetRef, normalizeAddress } from "./address.js";
import { deepMerge, normalizeStyle, styleForCell } from "./style.js";

function inferCell(value) {
  if (value === null || value === undefined) return { type: "blank", value: undefined };
  if (value instanceof Date) return { type: "date", value };
  if (typeof value === "number") return { type: "number", value };
  if (typeof value === "boolean") return { type: "boolean", value };
  return { type: "string", value: String(value) };
}

function deleteStylePath(style, path) {
  const aliases = {
    bold: "font.bold", italic: "font.italic", underline: "font.underline", strike: "font.strike",
    fontName: "font.name", fontSize: "font.size", color: "font.color",
    horizontal: "alignment.horizontal", vertical: "alignment.vertical", wrapText: "alignment.wrapText",
    textRotation: "alignment.textRotation", shrinkToFit: "alignment.shrinkToFit"
  };
  const parts = (aliases[path] ?? path).split(".");
  const parents = [];
  let target = style;
  for (const part of parts.slice(0, -1)) {
    if (!target?.[part] || typeof target[part] !== "object") return;
    parents.push([target, part]);
    target = target[part];
  }
  delete target?.[parts.at(-1)];
  for (const [parent, key] of parents.reverse()) if (!Object.keys(parent[key]).length) delete parent[key];
}

export class Cell {
  constructor(worksheet, address, onChange = () => {}, resolveStyle = normalizeStyle) {
    this.worksheet = worksheet;
    this.address = normalizeAddress(address);
    this.coordinates = decodeAddress(this.address);
    this.onChange = onChange;
    this.resolveStyle = resolveStyle;
  }

  get raw() {
    return this.worksheet[this.address];
  }

  get value() {
    return this.raw?.value;
  }

  set value(value) {
    this.worksheet[this.address] = { ...(this.raw ?? {}), ...inferCell(value), formatted: undefined };
    expandWorksheetRef(this.worksheet, this.coordinates);
    this.onChange(this.address, "value");
  }

  set(value) {
    this.value = value;
    return this;
  }

  formula(formula, result) {
    const valueCell = result === undefined ? { type: "number", value: undefined } : inferCell(result);
    this.worksheet[this.address] = {
      ...(this.raw ?? {}),
      ...valueCell,
      formula: String(formula).replace(/^=/, ""),
      formatted: undefined
    };
    expandWorksheetRef(this.worksheet, this.coordinates);
    this.onChange(this.address, "formula");
    return this;
  }

  style(style, mode = "merge") {
    const namedStyle = [...(Array.isArray(style) ? style : [style])].reverse().find((item) => typeof item === "string");
    return this.applyStyle(styleForCell(this.resolveStyle(style)), mode, namedStyle);
  }

  applyStyle(style, mode = "merge", namedStyle) {
    if (!["merge", "replace"].includes(mode)) throw new TypeError('Style mode must be "merge" or "replace".');
    const cell = this.ensure();
    cell.style = mode === "replace" ? structuredClone(style) : deepMerge(cell.style ?? {}, style);
    if (namedStyle) cell.namedStyle = namedStyle;
    else if (mode === "replace") delete cell.namedStyle;
    cell.styleDirty = true;
    this.onChange(this.address, "style");
    return this;
  }

  getStyle() {
    return structuredClone(this.raw?.style ?? {});
  }

  clearStyle(parts) {
    const cell = this.ensure();
    if (parts === undefined) {
      cell.style = {};
      delete cell.namedStyle;
    } else {
      const paths = Array.isArray(parts) ? parts : [parts];
      if (paths.some((path) => typeof path !== "string" || !path)) throw new TypeError("Style parts must be non-empty strings.");
      for (const path of paths) {
        if (path === "namedStyle") delete cell.namedStyle;
        else deleteStylePath(cell.style ?? {}, path);
      }
    }
    cell.styleDirty = true;
    this.onChange(this.address, "style");
    return this;
  }

  copyStyleFrom(source, mode = "replace") {
    const sourceCell = source instanceof Cell ? source : new Cell(this.worksheet, source, this.onChange, this.resolveStyle);
    return this.applyStyle(sourceCell.getStyle(), mode, sourceCell.raw?.namedStyle);
  }

  numberFormat(format) {
    const cell = this.ensure();
    cell.style = deepMerge(cell.style ?? {}, { numberFormat: String(format) });
    cell.styleDirty = true;
    this.onChange(this.address, "style");
    return this;
  }

  hyperlink(target, tooltip) {
    this.ensure("").hyperlink = { target: String(target), ...(tooltip ? { tooltip: String(tooltip) } : {}) };
    this.onChange(this.address, "hyperlink");
    return this;
  }

  clear(options = {}) {
    const style = options.keepStyle ? this.raw?.style : undefined;
    const styleIndex = options.keepStyle ? this.raw?.styleIndex : undefined;
    const namedStyle = options.keepStyle ? this.raw?.namedStyle : undefined;
    delete this.worksheet[this.address];
    if (style || styleIndex) this.worksheet[this.address] = { type: "blank", ...(style ? { style } : {}), ...(styleIndex ? { styleIndex } : {}), ...(namedStyle ? { namedStyle } : {}) };
    this.onChange(this.address, "clear");
    return this;
  }

  ensure(defaultValue) {
    if (!this.raw) this.value = defaultValue;
    return this.raw;
  }
}
