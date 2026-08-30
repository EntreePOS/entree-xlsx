import { writeFileSync } from "node:fs";
import { writeFile } from "node:fs/promises";
import { DuplicateSheetError, InvalidSheetNameError, SheetNotFoundError } from "./errors.js";
import { ChartCollection } from "./charts.js";
import { encryptWorkbookBuffer } from "./encryption.js";
import { writeXlsx } from "./ooxml.js";
import { getPackageState } from "./package-state.js";
import { PivotCollection } from "./pivots.js";
import { legacyPasswordHash } from "./protection.js";
import { Worksheet } from "./worksheet.js";
import { StyleCollection } from "./style.js";

function validateSheetName(name) {
  if (typeof name !== "string") throw new InvalidSheetNameError(String(name), "the name must be a string");
  if (!name.trim()) throw new InvalidSheetNameError(name, "the name cannot be empty");
  if (name.length > 31) throw new InvalidSheetNameError(name, "Excel limits names to 31 characters");
  if (/[\\/?*:[\]]/.test(name)) throw new InvalidSheetNameError(name, "the characters \\ / ? * : [ ] are not allowed");
  if (name.startsWith("'") || name.endsWith("'")) throw new InvalidSheetNameError(name, "the name cannot start or end with an apostrophe");
}

function assertXlsxPath(path) {
  if (!/\.xls[mx]$/i.test(path)) throw new TypeError("Use a filename ending in .xlsx or .xlsm.");
}

export class Workbook {
  constructor(source) {
    this.source = source;
  }

  get sheetNames() {
    return [...this.source.SheetNames];
  }

  get sheetCount() {
    return this.source.SheetNames.length;
  }

  get properties() {
    return { ...(this.source.Props ?? {}) };
  }

  get charts() {
    return new ChartCollection(this.source);
  }

  get pivots() {
    return new PivotCollection(this.source);
  }

  get styles() {
    return new StyleCollection(this.source);
  }

  protect(options = {}) {
    const config = typeof options === "string" ? { password: options } : options;
    this.source.WorkbookProtection = {
      lockStructure: config.structure !== false,
      lockWindows: config.windows === true,
      ...(config.password ? { workbookPassword: legacyPasswordHash(config.password) } : {})
    };
    const state = getPackageState(this.source);
    if (state) state.workbookDirty = true;
    return this;
  }

  unprotect() {
    delete this.source.WorkbookProtection;
    const state = getPackageState(this.source);
    if (state) state.workbookDirty = true;
    return this;
  }

  set properties(properties) {
    this.source.Props = { ...(this.source.Props ?? {}), ...properties };
    const state = getPackageState(this.source);
    if (state) state.propertiesDirty = true;
  }

  sheet(reference = 0) {
    const name = typeof reference === "number" ? this.source.SheetNames[reference] : reference;
    if (!name || !this.source.Sheets[name]) throw new SheetNotFoundError(reference, this.source.SheetNames);
    const state = getPackageState(this.source);
    return new Worksheet(
      this.source.Sheets[name],
      () => name,
      (address, kind) => state?.markCell(name, address, kind),
      () => state?.markSheetStructure(name),
      (style) => this.styles.resolve(style)
    );
  }

  trySheet(reference = 0) {
    try {
      return this.sheet(reference);
    } catch (error) {
      if (error instanceof SheetNotFoundError) return undefined;
      throw error;
    }
  }

  addSheet(name, data = []) {
    validateSheetName(name);
    if (this.source.Sheets[name]) throw new DuplicateSheetError(name);
    this.source.SheetNames.push(name);
    this.source.Sheets[name] = {};
    const state = getPackageState(this.source);
    if (state) state.workbookDirty = true;
    const sheet = this.sheet(name);
    if (data.length) sheet.addRows(data, { origin: "A1", skipHeader: false });
    return sheet;
  }

  removeSheet(reference) {
    if (this.source.SheetNames.length <= 1) throw new RangeError("A workbook must contain at least one worksheet.");
    const name = this.resolveName(reference);
    delete this.source.Sheets[name];
    this.source.SheetNames.splice(this.source.SheetNames.indexOf(name), 1);
    const state = getPackageState(this.source);
    if (state) state.workbookDirty = true;
    return this;
  }

  renameSheet(reference, newName) {
    validateSheetName(newName);
    const oldName = this.resolveName(reference);
    if (oldName !== newName && this.source.Sheets[newName]) throw new DuplicateSheetError(newName);
    if (oldName === newName) return this.sheet(oldName);
    const index = this.source.SheetNames.indexOf(oldName);
    this.source.Sheets[newName] = this.source.Sheets[oldName];
    delete this.source.Sheets[oldName];
    this.source.SheetNames[index] = newName;
    getPackageState(this.source)?.renameSheet(oldName, newName);
    return this.sheet(newName);
  }

  toBuffer(options = {}) {
    const buffer = writeXlsx(this.source);
    return options.password ? encryptWorkbookBuffer(buffer, options.password, options.encryption) : buffer;
  }

  toUint8Array(options = {}) {
    return new Uint8Array(this.toBuffer(options));
  }

  toBase64(options = {}) {
    return this.toBuffer(options).toString("base64");
  }

  async save(path, options = {}) {
    assertXlsxPath(path);
    await writeFile(path, this.toBuffer(options));
    return this;
  }

  saveSync(path, options = {}) {
    assertXlsxPath(path);
    writeFileSync(path, this.toBuffer(options));
    return this;
  }

  toJSON(options = {}) {
    return Object.fromEntries(this.source.SheetNames.map((name) => [name, this.sheet(name).toRecords(options)]));
  }

  resolveName(reference) {
    const name = typeof reference === "number" ? this.source.SheetNames[reference] : reference;
    if (!name || !this.source.Sheets[name]) throw new SheetNotFoundError(reference, this.source.SheetNames);
    return name;
  }
}
