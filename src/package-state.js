export const PACKAGE_STATE = Symbol("@entree_pos/xlsx.package-state");

export class PackageState {
  constructor({ files, workbookPath, workbookRelationshipsPath, stylesPath, sharedStringsPath, sheetPaths, corePath }) {
    this.files = new Map(files);
    this.workbookPath = workbookPath;
    this.workbookRelationshipsPath = workbookRelationshipsPath;
    this.stylesPath = stylesPath;
    this.sharedStringsPath = sharedStringsPath;
    this.sheetPaths = new Map(sheetPaths);
    this.corePath = corePath;
    this.dirtyCells = new Map();
    this.dirtySheetStructures = new Set();
    this.workbookDirty = false;
    this.propertiesDirty = false;
  }

  markCell(sheetName, address, kind = "value") {
    if (!this.dirtyCells.has(sheetName)) this.dirtyCells.set(sheetName, new Map());
    const cells = this.dirtyCells.get(sheetName);
    const kinds = cells.get(address) ?? new Set();
    kinds.add(kind);
    cells.set(address, kinds);
  }

  markSheetStructure(sheetName) {
    this.dirtySheetStructures.add(sheetName);
  }

  renameSheet(oldName, newName) {
    const path = this.sheetPaths.get(oldName);
    if (path) {
      this.sheetPaths.delete(oldName);
      this.sheetPaths.set(newName, path);
    }
    const cells = this.dirtyCells.get(oldName);
    if (cells) {
      this.dirtyCells.delete(oldName);
      this.dirtyCells.set(newName, cells);
    }
    if (this.dirtySheetStructures.delete(oldName)) this.dirtySheetStructures.add(newName);
    this.workbookDirty = true;
  }

  get dirty() {
    return this.workbookDirty || this.propertiesDirty || this.dirtyCells.size > 0 || this.dirtySheetStructures.size > 0;
  }
}

export function attachPackageState(workbook, state) {
  Object.defineProperty(workbook, PACKAGE_STATE, {
    value: state,
    configurable: false,
    enumerable: false,
    writable: false
  });
  return workbook;
}

export function getPackageState(workbook) {
  return workbook?.[PACKAGE_STATE];
}
