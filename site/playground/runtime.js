const encoder = new TextEncoder();

function clone(value) {
  return value === undefined ? undefined : structuredClone(value);
}

function merge(left = {}, right = {}) {
  const output = clone(left) ?? {};
  for (const [key, value] of Object.entries(right ?? {})) {
    output[key] = value && typeof value === "object" && !Array.isArray(value) && !(value instanceof Date)
      ? merge(output[key], value)
      : clone(value);
  }
  return output;
}

function columnIndex(reference) {
  if (Number.isInteger(reference) && reference >= 0) return reference;
  if (!/^[A-Z]+$/i.test(reference ?? "")) throw new TypeError(`Invalid column reference: ${reference}`);
  return [...String(reference).toUpperCase()].reduce((value, character) => value * 26 + character.charCodeAt(0) - 64, 0) - 1;
}

function columnLetter(index) {
  let value = index + 1;
  let output = "";
  while (value > 0) {
    value -= 1;
    output = String.fromCharCode(65 + value % 26) + output;
    value = Math.floor(value / 26);
  }
  return output;
}

function decodeAddress(address) {
  const match = /^([A-Z]+)([1-9]\d*)$/i.exec(String(address));
  if (!match) throw new TypeError(`Invalid cell address: ${address}`);
  return { c: columnIndex(match[1]), r: Number(match[2]) - 1 };
}

function encodeAddress({ r, c }) {
  return `${columnLetter(c)}${r + 1}`;
}

function decodeRange(reference) {
  const [start, end = start] = String(reference).split(":");
  const a = decodeAddress(start);
  const b = decodeAddress(end);
  return {
    s: { r: Math.min(a.r, b.r), c: Math.min(a.c, b.c) },
    e: { r: Math.max(a.r, b.r), c: Math.max(a.c, b.c) }
  };
}

function selectionNumbers(selector, decode, minimum) {
  if (Array.isArray(selector)) return [...new Set(selector.map(decode))];
  if (typeof selector === "string" && selector.includes(":")) {
    const [left, right] = selector.split(":").map(decode);
    return Array.from({ length: Math.abs(right - left) + 1 }, (_, index) => Math.min(left, right) + index);
  }
  const value = decode(selector);
  if (value < minimum) throw new RangeError("Selection is outside the worksheet.");
  return [value];
}

function isPopulated(cell) {
  return cell && (cell.value !== undefined || cell.formula);
}

class Cell {
  constructor(sheet, r, c) {
    this.sheet = sheet;
    this.r = r;
    this.c = c;
  }

  get address() { return encodeAddress(this); }
  get value() { return this.sheet._cells.get(this.address)?.value; }

  set(value) {
    const cell = this.sheet._ensure(this.r, this.c);
    cell.value = value;
    delete cell.formula;
    delete cell.result;
    return this;
  }

  get() { return this.value; }

  formula(formula, result) {
    const cell = this.sheet._ensure(this.r, this.c);
    cell.formula = String(formula).replace(/^=/, "");
    cell.result = result;
    return this;
  }

  style(style, mode = "merge") {
    const cell = this.sheet._ensure(this.r, this.c);
    const resolved = this.sheet.workbook.styles.resolve(style);
    cell.style = mode === "replace" ? resolved : merge(cell.style, resolved);
    return this;
  }

  numberFormat(format) { return this.style({ numberFormat: format }); }

  clear(options = {}) {
    const previous = this.sheet._cells.get(this.address);
    if (options.keepStyle && previous?.style) this.sheet._cells.set(this.address, { style: previous.style });
    else this.sheet._cells.delete(this.address);
    return this;
  }
}

class Selection {
  constructor(sheet, coordinates, metadata = {}) {
    this.sheet = sheet;
    this.coordinates = coordinates;
    this.metadata = metadata;
  }

  _existing() {
    return this.coordinates().filter(({ r, c }) => isPopulated(this.sheet._cells.get(encodeAddress({ r, c }))));
  }

  find(matcher) { return this.findAll(matcher)[0]; }

  findAll(matcher) {
    const predicate = typeof matcher === "function" ? matcher : (cell) => cell.value === matcher;
    return this._existing().map(({ r, c }) => new Cell(this.sheet, r, c)).filter(predicate);
  }

  forEach(callback) {
    this._existing().forEach(({ r, c }, index) => callback(new Cell(this.sheet, r, c), index));
    return this;
  }

  style(style, mode = "merge") {
    const resolved = this.sheet.workbook.styles.resolve(style);
    if (this.metadata.rows) {
      for (const row of this.metadata.rows) this.sheet._rowStyles.set(row, mode === "replace" ? resolved : merge(this.sheet._rowStyles.get(row), resolved));
    } else if (this.metadata.columns) {
      for (const column of this.metadata.columns) this.sheet._columnStyles.set(column, mode === "replace" ? resolved : merge(this.sheet._columnStyles.get(column), resolved));
    } else {
      for (const { r, c } of this.coordinates()) new Cell(this.sheet, r, c).style(resolved, mode);
    }
    return this;
  }

  width(width) {
    if (!this.metadata.columns) throw new TypeError("width() is available on column selections.");
    for (const column of this.metadata.columns) this.sheet._columnWidths.set(column, width);
    return this;
  }

  height(height) {
    if (!this.metadata.rows) throw new TypeError("height() is available on row selections.");
    for (const row of this.metadata.rows) this.sheet._rowHeights.set(row, height);
    return this;
  }

  setValues(rows) {
    const range = this.metadata.range;
    if (!range) throw new TypeError("setValues() is available on ranges.");
    rows.forEach((values, rowOffset) => values.forEach((value, columnOffset) => {
      const r = range.s.r + rowOffset;
      const c = range.s.c + columnOffset;
      if (r <= range.e.r && c <= range.e.c) this.sheet.cell({ r, c }).set(value);
    }));
    return this;
  }

  getValues() {
    const range = this.metadata.range;
    if (!range) throw new TypeError("getValues() is available on ranges.");
    return Array.from({ length: range.e.r - range.s.r + 1 }, (_, rowOffset) =>
      Array.from({ length: range.e.c - range.s.c + 1 }, (_, columnOffset) => this.sheet.cell({ r: range.s.r + rowOffset, c: range.s.c + columnOffset }).value));
  }

  merge() {
    if (!this.metadata.range) throw new TypeError("merge() is available on ranges.");
    this.sheet.merge(`${encodeAddress(this.metadata.range.s)}:${encodeAddress(this.metadata.range.e)}`);
    return this;
  }
}

class StyleCollection {
  constructor() { this.definitions = new Map(); }

  define(name, style, options = {}) {
    const parents = Array.isArray(options.extends) ? options.extends : options.extends ? [options.extends] : [];
    this.definitions.set(name, { style: clone(style), parents });
    return this;
  }

  defineMany(definitions) {
    for (const [name, definition] of Object.entries(definitions)) {
      const { extends: parents, ...style } = definition;
      this.define(name, style, { extends: parents });
    }
    return this;
  }

  resolve(input, seen = new Set()) {
    if (Array.isArray(input)) return input.reduce((result, item) => merge(result, this.resolve(item, seen)), {});
    if (typeof input !== "string") return clone(input) ?? {};
    if (seen.has(input)) throw new Error(`Circular style definition: ${input}`);
    const definition = this.definitions.get(input);
    if (!definition) throw new Error(`Unknown style: ${input}`);
    const next = new Set(seen).add(input);
    return merge(definition.parents.reduce((result, parent) => merge(result, this.resolve(parent, next)), {}), definition.style);
  }
}

class Worksheet {
  constructor(workbook, name) {
    this.workbook = workbook;
    this.name = name;
    this._cells = new Map();
    this._columnWidths = new Map();
    this._rowHeights = new Map();
    this._columnStyles = new Map();
    this._rowStyles = new Map();
    this._merges = [];
    this._autoFilter = undefined;
  }

  _ensure(r, c) {
    const address = encodeAddress({ r, c });
    if (!this._cells.has(address)) this._cells.set(address, {});
    return this._cells.get(address);
  }

  cell(address) {
    const position = typeof address === "string" ? decodeAddress(address) : address;
    return new Cell(this, position.r, position.c);
  }

  set(address, value) { this.cell(address).set(value); return this; }
  get(address) { return this.cell(address).value; }

  setData(data) {
    this._cells.clear();
    if (!data.length) return this;
    const rows = Array.isArray(data[0])
      ? data
      : [Object.keys(data[0]), ...data.map((record) => Object.keys(data[0]).map((key) => record[key]))];
    rows.forEach((row, r) => row.forEach((value, c) => this.cell({ r, c }).set(value)));
    return this;
  }

  appendRows(rows) {
    const start = this._bounds()?.e.r + 1 || 0;
    rows.forEach((row, offset) => row.forEach((value, c) => this.cell({ r: start + offset, c }).set(value)));
    return this;
  }

  range(reference) {
    const range = decodeRange(reference);
    return new Selection(this, () => {
      const cells = [];
      for (let r = range.s.r; r <= range.e.r; r += 1) for (let c = range.s.c; c <= range.e.c; c += 1) cells.push({ r, c });
      return cells;
    }, { range });
  }

  row(number) { return this.rows(number); }

  rows(selector) {
    const rows = selectionNumbers(selector, (value) => {
      const number = Number(value);
      if (!Number.isInteger(number) || number < 1) throw new TypeError(`Invalid row: ${value}`);
      return number - 1;
    }, 0);
    return new Selection(this, () => {
      const end = Math.max(this._bounds()?.e.c ?? 0, 0);
      return rows.flatMap((r) => Array.from({ length: end + 1 }, (_, c) => ({ r, c })));
    }, { rows });
  }

  column(reference) { return this.columns(reference); }

  columns(selector) {
    const columns = selectionNumbers(selector, columnIndex, 0);
    return new Selection(this, () => {
      const end = Math.max(this._bounds()?.e.r ?? 0, 0);
      return Array.from({ length: end + 1 }, (_, r) => columns.map((c) => ({ r, c }))).flat();
    }, { columns });
  }

  find(matcher) { return this.findAll(matcher)[0]; }
  findAll(matcher) {
    const bounds = this._bounds();
    if (!bounds) return [];
    return this.range(`${encodeAddress(bounds.s)}:${encodeAddress(bounds.e)}`).findAll(matcher);
  }

  merge(reference) { this._merges.push(decodeRange(reference)); return this; }
  unmerge(reference) {
    const target = JSON.stringify(decodeRange(reference));
    this._merges = this._merges.filter((range) => JSON.stringify(range) !== target);
    return this;
  }
  autoFilter(reference) {
    const bounds = this._bounds();
    this._autoFilter = reference ?? (bounds ? `${encodeAddress(bounds.s)}:${encodeAddress(bounds.e)}` : undefined);
    return this;
  }
  setColumnWidth(reference, width) { this._columnWidths.set(columnIndex(reference), width); return this; }
  setRowHeight(row, height) { this._rowHeights.set(row - 1, height); return this; }

  autoFit(options = {}) {
    const bounds = this._bounds();
    if (!bounds) return this;
    for (let c = bounds.s.c; c <= bounds.e.c; c += 1) {
      let width = options.min ?? 8;
      for (let r = bounds.s.r; r <= bounds.e.r; r += 1) width = Math.max(width, String(this.cell({ r, c }).value ?? "").length + (options.padding ?? 2));
      this._columnWidths.set(c, Math.min(width, options.max ?? 48));
    }
    return this;
  }

  toRows() {
    const bounds = this._bounds();
    if (!bounds) return [];
    return Array.from({ length: bounds.e.r + 1 }, (_, r) => Array.from({ length: bounds.e.c + 1 }, (_, c) => this.cell({ r, c }).value ?? null));
  }

  toRecords() {
    const [headers = [], ...rows] = this.toRows();
    return rows.map((row) => Object.fromEntries(headers.map((header, index) => [String(header), row[index]])));
  }

  toCsv() {
    return this.toRows().map((row) => row.map((value) => {
      const text = String(value ?? "");
      return /[",\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
    }).join(",")).join("\n");
  }

  _bounds() {
    const positions = [...this._cells.entries()].filter(([, cell]) => isPopulated(cell)).map(([address]) => decodeAddress(address));
    if (!positions.length) return undefined;
    return {
      s: { r: Math.min(...positions.map(({ r }) => r)), c: Math.min(...positions.map(({ c }) => c)) },
      e: { r: Math.max(...positions.map(({ r }) => r)), c: Math.max(...positions.map(({ c }) => c)) }
    };
  }

  _effectiveStyle(r, c, cell = {}) {
    return merge(merge(this._columnStyles.get(c), this._rowStyles.get(r)), cell.style);
  }
}

class Workbook {
  constructor(name = "Sheet1") {
    this.styles = new StyleCollection();
    this._sheets = [new Worksheet(this, name)];
    this.outputName = "playground.xlsx";
  }

  sheet(reference = 0) {
    const sheet = typeof reference === "number" ? this._sheets[reference] : this._sheets.find(({ name }) => name === reference);
    if (!sheet) throw new Error(`Worksheet not found: ${reference}`);
    return sheet;
  }

  findSheet(reference = 0) {
    return typeof reference === "number" ? this._sheets[reference] : this._sheets.find(({ name }) => name === reference);
  }

  addSheet(name, data) {
    if (this.findSheet(name)) throw new Error(`Worksheet already exists: ${name}`);
    const sheet = new Worksheet(this, name);
    this._sheets.push(sheet);
    if (data) sheet.setData(data);
    return sheet;
  }

  async save(name = "playground.xlsx") { this.outputName = name; return this; }
}

export function createWorkbook(name) {
  return new Workbook(name);
}

function xml(value) {
  return String(value).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&apos;");
}

function styleFingerprint(value) {
  if (Array.isArray(value)) return `[${value.map(styleFingerprint).join(",")}]`;
  if (value && typeof value === "object") return `{${Object.keys(value).sort().map((key) => `${JSON.stringify(key)}:${styleFingerprint(value[key])}`).join(",")}}`;
  return JSON.stringify(value);
}

function normalizedStyle(style = {}) {
  const font = merge(style.font, {
    ...(style.bold !== undefined ? { bold: style.bold } : {}),
    ...(style.italic !== undefined ? { italic: style.italic } : {}),
    ...(style.fontName ? { name: style.fontName } : {}),
    ...(style.fontSize ? { size: style.fontSize } : {}),
    ...(style.color ? { color: style.color } : {})
  });
  const alignment = merge({
    ...(style.horizontal ? { horizontal: style.horizontal } : {}),
    ...(style.vertical ? { vertical: style.vertical } : {}),
    ...(style.wrapText ? { wrapText: style.wrapText } : {})
  }, style.alignment);
  return {
    ...(Object.keys(font).length ? { font } : {}),
    ...(style.fill ? { fill: style.fill } : {}),
    ...(style.numberFormat ? { numberFormat: style.numberFormat } : {}),
    ...(Object.keys(alignment).length ? { alignment } : {})
  };
}

function buildStyles(workbook) {
  const styles = [{}];
  const indexes = new Map([[styleFingerprint({}), 0]]);
  const index = (input) => {
    const style = normalizedStyle(input);
    const key = styleFingerprint(style);
    if (!indexes.has(key)) { indexes.set(key, styles.length); styles.push(style); }
    return indexes.get(key);
  };
  for (const sheet of workbook._sheets) {
    for (const [address, cell] of sheet._cells) {
      const { r, c } = decodeAddress(address);
      index(sheet._effectiveStyle(r, c, cell));
    }
    for (const style of sheet._rowStyles.values()) index(style);
    for (const style of sheet._columnStyles.values()) index(style);
  }

  const fonts = [{ name: "Aptos", size: 11 }];
  const fills = [null, { gray125: true }];
  const formats = [];
  const xfs = [{ fontId: 0, fillId: 0, numFmtId: 0 }];
  styles.slice(1).forEach((style) => {
    const fontId = style.font ? fonts.push(style.font) - 1 : 0;
    const fillId = style.fill ? fills.push(style.fill) - 1 : 0;
    const numFmtId = style.numberFormat ? 164 + formats.push(style.numberFormat) - 1 : 0;
    xfs.push({ fontId, fillId, numFmtId, alignment: style.alignment });
  });

  const color = (value) => String(typeof value === "object" ? value.rgb ?? "000000" : value).replace(/^#/, "").toUpperCase();
  const fontXml = fonts.map((font) => `<font>${font.bold ? "<b/>" : ""}${font.italic ? "<i/>" : ""}<sz val="${font.size ?? 11}"/><name val="${xml(font.name ?? "Aptos")}"/>${font.color ? `<color rgb="FF${color(font.color)}"/>` : ""}</font>`).join("");
  const fillXml = fills.map((fill) => fill?.gray125
    ? '<fill><patternFill patternType="gray125"/></fill>'
    : fill ? `<fill><patternFill patternType="solid"><fgColor rgb="FF${color(typeof fill === "object" ? fill.foreground?.rgb ?? fill.color ?? "FFFFFF" : fill)}"/><bgColor indexed="64"/></patternFill></fill>`
      : '<fill><patternFill patternType="none"/></fill>').join("");
  const formatXml = formats.map((format, offset) => `<numFmt numFmtId="${164 + offset}" formatCode="${xml(format)}"/>`).join("");
  const xfXml = xfs.map((xf) => {
    const alignment = xf.alignment ? `<alignment${xf.alignment.horizontal ? ` horizontal="${xml(xf.alignment.horizontal)}"` : ""}${xf.alignment.vertical ? ` vertical="${xml(xf.alignment.vertical)}"` : ""}${xf.alignment.wrapText ? ' wrapText="1"' : ""}/>` : "";
    return `<xf numFmtId="${xf.numFmtId}" fontId="${xf.fontId}" fillId="${xf.fillId}" borderId="0" xfId="0"${xf.numFmtId ? ' applyNumberFormat="1"' : ""}${xf.fontId ? ' applyFont="1"' : ""}${xf.fillId ? ' applyFill="1"' : ""}${alignment ? ' applyAlignment="1"' : ""}>${alignment}</xf>`;
  }).join("");
  const output = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${formats.length ? `<numFmts count="${formats.length}">${formatXml}</numFmts>` : ""}<fonts count="${fonts.length}">${fontXml}</fonts><fills count="${fills.length}">${fillXml}</fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfXml}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
  return { output, index };
}

function sheetXml(sheet, styleIndex) {
  const bounds = sheet._bounds() ?? { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
  const rows = new Map();
  for (const [address, cell] of sheet._cells) {
    if (!isPopulated(cell)) continue;
    const { r, c } = decodeAddress(address);
    const style = styleIndex(sheet._effectiveStyle(r, c, cell));
    let value;
    let type = "";
    if (cell.formula) {
      value = cell.result ?? 0;
      if (typeof value === "string") type = ' t="str"';
      else if (typeof value === "boolean") { value = value ? 1 : 0; type = ' t="b"'; }
    } else if (cell.value instanceof Date) {
      value = (cell.value.getTime() - Date.UTC(1899, 11, 30)) / 86400000;
    } else if (typeof cell.value === "number") value = cell.value;
    else if (typeof cell.value === "boolean") { value = cell.value ? 1 : 0; type = ' t="b"'; }
    else { value = cell.value ?? cell.result ?? ""; type = ' t="inlineStr"'; }
    const formula = cell.formula ? `<f>${xml(cell.formula)}</f>` : "";
    const body = type.includes("inlineStr") ? `<is><t>${xml(value)}</t></is>` : `<v>${xml(value)}</v>`;
    const cellXml = `<c r="${address}"${type}${style ? ` s="${style}"` : ""}>${formula}${body}</c>`;
    if (!rows.has(r)) rows.set(r, []);
    rows.get(r).push({ c, xml: cellXml });
  }
  for (const r of new Set([...sheet._rowStyles.keys(), ...sheet._rowHeights.keys()])) if (!rows.has(r)) rows.set(r, []);
  const rowXml = [...rows].sort(([a], [b]) => a - b).map(([r, cells]) => {
    const height = sheet._rowHeights.get(r);
    const style = sheet._rowStyles.has(r) ? styleIndex(sheet._rowStyles.get(r)) : 0;
    return `<row r="${r + 1}"${height ? ` ht="${height}" customHeight="1"` : ""}${style ? ` s="${style}" customFormat="1"` : ""}>${cells.sort((a, b) => a.c - b.c).map((cell) => cell.xml).join("")}</row>`;
  }).join("");
  const columnSet = new Set([...sheet._columnWidths.keys(), ...sheet._columnStyles.keys()]);
  const columns = [...columnSet].sort((a, b) => a - b).map((c) => {
    const width = sheet._columnWidths.get(c);
    const style = sheet._columnStyles.has(c) ? styleIndex(sheet._columnStyles.get(c)) : 0;
    return `<col min="${c + 1}" max="${c + 1}"${width ? ` width="${width}" customWidth="1"` : ""}${style ? ` style="${style}"` : ""}/>`;
  }).join("");
  const merges = sheet._merges.map((range) => `<mergeCell ref="${encodeAddress(range.s)}:${encodeAddress(range.e)}"/>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="${encodeAddress(bounds.s)}:${encodeAddress(bounds.e)}"/>${columns ? `<cols>${columns}</cols>` : ""}<sheetData>${rowXml}</sheetData>${merges ? `<mergeCells count="${sheet._merges.length}">${merges}</mergeCells>` : ""}${sheet._autoFilter ? `<autoFilter ref="${xml(sheet._autoFilter)}"/>` : ""}</worksheet>`;
}

function crc32(bytes) {
  let crc = 0xFFFFFFFF;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ (0xEDB88320 & -(crc & 1));
  }
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

function join(parts) {
  const output = new Uint8Array(parts.reduce((size, part) => size + part.length, 0));
  let offset = 0;
  for (const part of parts) { output.set(part, offset); offset += part.length; }
  return output;
}

function zip(files) {
  const locals = [];
  const centrals = [];
  let offset = 0;
  for (const [name, source] of Object.entries(files)) {
    const nameBytes = encoder.encode(name);
    const data = source instanceof Uint8Array ? source : encoder.encode(source);
    const crc = crc32(data);
    const local = new Uint8Array(30 + nameBytes.length + data.length);
    const lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034B50, true); lv.setUint16(4, 20, true); lv.setUint16(6, 0x800, true);
    lv.setUint32(14, crc, true); lv.setUint32(18, data.length, true); lv.setUint32(22, data.length, true); lv.setUint16(26, nameBytes.length, true);
    local.set(nameBytes, 30); local.set(data, 30 + nameBytes.length); locals.push(local);
    const central = new Uint8Array(46 + nameBytes.length);
    const cv = new DataView(central.buffer);
    cv.setUint32(0, 0x02014B50, true); cv.setUint16(4, 20, true); cv.setUint16(6, 20, true); cv.setUint16(8, 0x800, true);
    cv.setUint32(16, crc, true); cv.setUint32(20, data.length, true); cv.setUint32(24, data.length, true); cv.setUint16(28, nameBytes.length, true); cv.setUint32(42, offset, true);
    central.set(nameBytes, 46); centrals.push(central); offset += local.length;
  }
  const directory = join(centrals);
  const end = new Uint8Array(22);
  const view = new DataView(end.buffer);
  view.setUint32(0, 0x06054B50, true); view.setUint16(8, centrals.length, true); view.setUint16(10, centrals.length, true); view.setUint32(12, directory.length, true); view.setUint32(16, offset, true);
  return join([...locals, directory, end]);
}

export function workbookToXlsx(workbook) {
  const styles = buildStyles(workbook);
  const sheetOverrides = workbook._sheets.map((_, index) => `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("");
  const sheets = workbook._sheets.map((sheet, index) => `<sheet name="${xml(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join("");
  const relationships = workbook._sheets.map((_, index) => `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`).join("");
  const files = {
    "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheetOverrides}</Types>`,
    "_rels/.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    "xl/workbook.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets}</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`,
    "xl/_rels/workbook.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships}<Relationship Id="rId${workbook._sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
    "xl/styles.xml": styles.output
  };
  workbook._sheets.forEach((sheet, index) => { files[`xl/worksheets/sheet${index + 1}.xml`] = sheetXml(sheet, styles.index); });
  return zip(files);
}

export function serializeWorkbook(workbook) {
  return {
    name: workbook.outputName,
    sheets: workbook._sheets.map((sheet) => ({
      name: sheet.name,
      cells: [...sheet._cells.entries()].filter(([, cell]) => isPopulated(cell)).map(([address, cell]) => {
        const { r, c } = decodeAddress(address);
        return { r, c, address, value: cell.value, formula: cell.formula, result: cell.result, style: sheet._effectiveStyle(r, c, cell) };
      }),
      columnWidths: Object.fromEntries(sheet._columnWidths),
      rowHeights: Object.fromEntries(sheet._rowHeights),
      merges: sheet._merges
    }))
  };
}

export async function executePlayground(source) {
  let lastWorkbook;
  const logs = [];
  const playgroundCreateWorkbook = (name) => (lastWorkbook = createWorkbook(name));
  const playgroundConsole = Object.fromEntries(["log", "info", "warn", "error"].map((level) => [level, (...values) => logs.push({
    level,
    text: values.map((value) => typeof value === "string" ? value : JSON.stringify(value)).join(" ")
  })]));
  const cleaned = source.replace(/^\s*import\s+[^;]+;?\s*$/gm, "");
  const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
  const run = new AsyncFunction("createWorkbook", "console", `"use strict";\n${cleaned}`);
  await run(playgroundCreateWorkbook, playgroundConsole);
  if (!lastWorkbook) throw new Error("Create a workbook with createWorkbook() before running the preview.");
  return { workbook: serializeWorkbook(lastWorkbook), bytes: workbookToXlsx(lastWorkbook), logs };
}
