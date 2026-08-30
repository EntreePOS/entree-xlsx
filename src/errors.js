export class XlsxError extends Error {
  constructor(message, code = "XLSX_ERROR", options) {
    super(message, options);
    this.name = new.target.name;
    this.code = code;
  }
}

export class SheetNotFoundError extends XlsxError {
  constructor(sheet, available) {
    super(
      `Worksheet ${JSON.stringify(sheet)} was not found. Available worksheets: ${available.length ? available.join(", ") : "none"}.`,
      "SHEET_NOT_FOUND"
    );
  }
}

export class DuplicateSheetError extends XlsxError {
  constructor(name) {
    super(`A worksheet named ${JSON.stringify(name)} already exists.`, "DUPLICATE_SHEET");
  }
}

export class InvalidSheetNameError extends XlsxError {
  constructor(name, reason) {
    super(`Invalid worksheet name ${JSON.stringify(name)}: ${reason}.`, "INVALID_SHEET_NAME");
  }
}

export class InvalidSourceError extends XlsxError {
  constructor(message, options) {
    super(message, "INVALID_SOURCE", options);
  }
}

export class EncryptionError extends XlsxError {
  constructor(message, options) {
    super(message, "ENCRYPTION_ERROR", options);
  }
}
