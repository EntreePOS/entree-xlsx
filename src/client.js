import { readFileSync } from "node:fs";
import { readFile } from "node:fs/promises";
import { InvalidSourceError } from "./errors.js";
import { decryptWorkbookBuffer, isCompoundFile } from "./encryption.js";
import { readXlsx } from "./ooxml.js";
import { Workbook } from "./workbook.js";

function bytes(value) {
  if (Buffer.isBuffer(value)) return value;
  if (value instanceof Uint8Array) return Buffer.from(value.buffer, value.byteOffset, value.byteLength);
  if (value instanceof ArrayBuffer) return Buffer.from(value);
  throw new InvalidSourceError("Workbook data must be a Buffer, Uint8Array, or ArrayBuffer.");
}

export class XlsxClient {
  create(firstSheetName = "Sheet1") {
    const workbook = new Workbook({ SheetNames: [], Sheets: {}, Props: {} });
    workbook.addSheet(firstSheetName);
    return workbook;
  }

  parse(data, options = {}) {
    try {
      let input = bytes(data);
      if (isCompoundFile(input)) {
        if (!options.password) throw new InvalidSourceError("This workbook is password encrypted. Pass { password } to open or parse it.");
        input = decryptWorkbookBuffer(input, options.password, options.encryption);
      }
      return new Workbook(readXlsx(input));
    } catch (error) {
      if (error instanceof InvalidSourceError) throw error;
      throw new InvalidSourceError("Unable to parse XLSX data.", { cause: error });
    }
  }

  async open(source, options = {}) {
    if (source instanceof URL || (typeof source === "string" && /^https?:\/\//i.test(source))) {
      const response = await fetch(source);
      if (!response.ok) throw new InvalidSourceError(`Unable to download workbook: ${response.status} ${response.statusText}.`);
      return this.parse(await response.arrayBuffer(), options);
    }
    if (source instanceof ArrayBuffer || source instanceof Uint8Array) return this.parse(source, options);
    if (typeof source !== "string") throw new InvalidSourceError("Workbook source must be a path, URL, Buffer, Uint8Array, or ArrayBuffer.");
    try {
      return this.parse(await readFile(source), options);
    } catch (error) {
      if (error instanceof InvalidSourceError) throw error;
      throw new InvalidSourceError(`Unable to open XLSX file ${JSON.stringify(source)}.`, { cause: error });
    }
  }

  openSync(path, options = {}) {
    try {
      return this.parse(readFileSync(path), options);
    } catch (error) {
      if (error instanceof InvalidSourceError) throw error;
      throw new InvalidSourceError(`Unable to open XLSX file ${JSON.stringify(path)}.`, { cause: error });
    }
  }
}
