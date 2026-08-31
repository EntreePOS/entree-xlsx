import { decodeAddress, decodeRange, encodeAddress, encodeRange, expandWorksheetRef } from "./address.js";
import { ensurePackageState } from "./ooxml.js";
import {
  addContentTypeOverride,
  addRelationship,
  getTextFile,
  nextPartPath,
  readRelationships,
  relationshipTarget,
  relationshipsPath,
  removeContentTypeOverride,
  removeRelationship,
  setTextFile
} from "./package-xml.js";
import { allTags, decodeXml, escapeXml } from "./xml.js";

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CACHE_DEF_REL = `${REL_NS}/pivotCacheDefinition`;
const CACHE_RECORDS_REL = `${REL_NS}/pivotCacheRecords`;
const TABLE_REL = `${REL_NS}/pivotTable`;

function cellValue(sheet, row, column) {
  return sheet[encodeAddress({ r: row, c: column })]?.value;
}

function sourceData(workbook, source) {
  const sheet = workbook.Sheets[source.sheet];
  if (!sheet) throw new RangeError(`Pivot source worksheet ${source.sheet} does not exist.`);
  const range = decodeRange(source.range ?? sheet["!ref"]);
  if (!range || range.e.r <= range.s.r) throw new RangeError("Pivot source needs a header row and at least one data row.");
  const fields = [];
  for (let column = range.s.c; column <= range.e.c; column += 1) {
    fields.push(String(cellValue(sheet, range.s.r, column) ?? `Column${column - range.s.c + 1}`));
  }
  const records = [];
  for (let row = range.s.r + 1; row <= range.e.r; row += 1) {
    records.push(fields.map((_, index) => cellValue(sheet, row, range.s.c + index)));
  }
  return { fields, records, range: encodeRange(range) };
}

function fieldIndexes(fields, names, area) {
  return (names ?? []).map((name) => {
    const index = fields.indexOf(name);
    if (index < 0) throw new RangeError(`Pivot ${area} field ${name} is not present in the source header.`);
    return index;
  });
}

function normalizeValues(fields, values) {
  if (!values?.length) throw new TypeError("A pivot table requires at least one value field.");
  return values.map((item) => {
    const config = typeof item === "string" ? { field: item } : item;
    const field = fields.indexOf(config.field);
    if (field < 0) throw new RangeError(`Pivot value field ${config.field} is not present in the source header.`);
    const summarize = config.summarize ?? "sum";
    if (!["sum", "count", "average", "min", "max"].includes(summarize)) throw new RangeError(`Unsupported pivot summary: ${summarize}.`);
    return { field, fieldName: config.field, summarize, name: config.name ?? `${summarize === "sum" ? "Sum" : summarize[0].toUpperCase() + summarize.slice(1)} of ${config.field}` };
  });
}

function normalizeConfig(workbook, config) {
  if (!config?.source?.sheet || !config?.target?.sheet || !config?.target?.cell) {
    throw new TypeError("Pivot source.sheet, target.sheet, and target.cell are required.");
  }
  if (!workbook.Sheets[config.target.sheet]) throw new RangeError(`Pivot target worksheet ${config.target.sheet} does not exist.`);
  const data = sourceData(workbook, config.source);
  return {
    name: config.name,
    source: { sheet: config.source.sheet, range: data.range },
    target: { sheet: config.target.sheet, cell: encodeAddress(decodeAddress(config.target.cell)) },
    fields: data.fields,
    records: data.records,
    rows: fieldIndexes(data.fields, config.rows, "row"),
    columns: fieldIndexes(data.fields, config.columns, "column"),
    filters: fieldIndexes(data.fields, config.filters, "filter"),
    values: normalizeValues(data.fields, config.values),
    refreshOnLoad: config.refreshOnLoad !== false,
    showGrandTotals: config.showGrandTotals !== false,
    style: config.style ?? "PivotStyleMedium9"
  };
}

function typedCacheValue(value) {
  if (value === undefined || value === null || value === "") return "<m/>";
  if (value instanceof Date) return `<d v="${escapeXml(value.toISOString())}"/>`;
  if (typeof value === "boolean") return `<b v="${value ? 1 : 0}"/>`;
  if (typeof value === "number") return `<n v="${value}"/>`;
  return `<s v="${escapeXml(value)}"/>`;
}

function cacheFieldXml(name, values) {
  const strings = values.filter((value) => typeof value === "string");
  const numbers = values.filter((value) => typeof value === "number");
  const dates = values.filter((value) => value instanceof Date);
  const blanks = values.filter((value) => value === undefined || value === null || value === "");
  const uniqueStrings = [...new Set(strings)];
  const attributes = [
    uniqueStrings.length ? 'containsString="1"' : "",
    numbers.length ? 'containsNumber="1"' : "",
    dates.length ? 'containsDate="1"' : "",
    blanks.length ? 'containsBlank="1"' : "",
    `count="${uniqueStrings.length}"`
  ].filter(Boolean).join(" ");
  return `<cacheField name="${escapeXml(name)}" numFmtId="0"><sharedItems ${attributes}>${uniqueStrings.map((value) => `<s v="${escapeXml(value)}"/>`).join("")}</sharedItems></cacheField>`;
}

function cacheDefinitionXml(config, recordRelationshipId) {
  const fields = config.fields.map((name, index) => cacheFieldXml(name, config.records.map((record) => record[index]))).join("");
  return `${XML_HEADER}<pivotCacheDefinition xmlns="${MAIN_NS}" xmlns:r="${REL_NS}" r:id="${recordRelationshipId}" saveData="1" refreshOnLoad="${config.refreshOnLoad ? 1 : 0}" recordCount="${config.records.length}" createdVersion="3" refreshedVersion="8" minRefreshableVersion="3"><cacheSource type="worksheet"><worksheetSource ref="${escapeXml(config.source.range)}" sheet="${escapeXml(config.source.sheet)}"/></cacheSource><cacheFields count="${config.fields.length}">${fields}</cacheFields></pivotCacheDefinition>`;
}

function cacheRecordsXml(config) {
  return `${XML_HEADER}<pivotCacheRecords xmlns="${MAIN_NS}" count="${config.records.length}">${config.records.map((record) => `<r>${record.map(typedCacheValue).join("")}</r>`).join("")}</pivotCacheRecords>`;
}

function aggregate(values, summarize) {
  const present = values.filter((value) => value !== undefined && value !== null && value !== "");
  if (summarize === "count") return present.length;
  const numbers = present.map(Number).filter(Number.isFinite);
  if (!numbers.length) return 0;
  if (summarize === "average") return numbers.reduce((sum, value) => sum + value, 0) / numbers.length;
  if (summarize === "min") return Math.min(...numbers);
  if (summarize === "max") return Math.max(...numbers);
  return numbers.reduce((sum, value) => sum + value, 0);
}

function keyFor(record, indexes) {
  return indexes.map((index) => String(record[index] ?? "(blank)"));
}

function uniqueKeys(records, indexes) {
  const seen = new Set();
  const result = [];
  for (const record of records) {
    const parts = keyFor(record, indexes);
    const key = JSON.stringify(parts);
    if (!seen.has(key)) { seen.add(key); result.push(parts); }
  }
  return indexes.length ? result : [[]];
}

function renderPivot(config) {
  const rowKeys = uniqueKeys(config.records, config.rows);
  const columnKeys = uniqueKeys(config.records, config.columns);
  const headers = config.rows.map((index) => config.fields[index]);
  for (const columnKey of columnKeys) {
    for (const value of config.values) {
      headers.push(config.values.length === 1
        ? columnKey.join(" - ")
        : [...columnKey, value.name].filter(Boolean).join(" - "));
    }
  }
  if (config.showGrandTotals) {
    for (const value of config.values) {
      headers.push(config.values.length === 1 ? "Grand Total" : `Grand Total - ${value.name}`);
    }
  }
  const output = [headers];
  for (const rowKey of rowKeys) {
    const row = [...rowKey];
    for (const columnKey of columnKeys) {
      const matching = config.records.filter((record) =>
        JSON.stringify(keyFor(record, config.rows)) === JSON.stringify(rowKey) &&
        JSON.stringify(keyFor(record, config.columns)) === JSON.stringify(columnKey));
      for (const value of config.values) row.push(aggregate(matching.map((record) => record[value.field]), value.summarize));
    }
    if (config.showGrandTotals) {
      const matching = config.records.filter((record) =>
        JSON.stringify(keyFor(record, config.rows)) === JSON.stringify(rowKey));
      for (const value of config.values) {
        row.push(aggregate(matching.map((record) => record[value.field]), value.summarize));
      }
    }
    output.push(row);
  }
  if (config.showGrandTotals) {
    const row = [...config.rows.map((_, index) => index === 0 ? "Total" : "")];
    for (const columnKey of columnKeys) {
      const matching = config.records.filter((record) => JSON.stringify(keyFor(record, config.columns)) === JSON.stringify(columnKey));
      for (const value of config.values) row.push(aggregate(matching.map((record) => record[value.field]), value.summarize));
    }
    for (const value of config.values) {
      row.push(aggregate(config.records.map((record) => record[value.field]), value.summarize));
    }
    output.push(row);
  }
  return output;
}

function writeRenderedCells(workbook, state, config) {
  const sheet = workbook.Sheets[config.target.sheet];
  const start = decodeAddress(config.target.cell);
  const output = renderPivot(config);
  output.forEach((row, rowOffset) => row.forEach((value, columnOffset) => {
    const point = { r: start.r + rowOffset, c: start.c + columnOffset };
    const address = encodeAddress(point);
    sheet[address] = { type: typeof value === "number" ? "number" : "string", value };
    expandWorksheetRef(sheet, point);
    state?.markCell(config.target.sheet, address, "value");
  }));
  const end = { r: start.r + output.length - 1, c: start.c + Math.max(...output.map((row) => row.length)) - 1 };
  return encodeRange({ s: start, e: end });
}

function pivotFieldXml(config, index) {
  const row = config.rows.includes(index);
  const column = config.columns.includes(index);
  const filter = config.filters.includes(index);
  const data = config.values.some((value) => value.field === index);
  const axis = row ? ' axis="axisRow"' : column ? ' axis="axisCol"' : filter ? ' axis="axisPage"' : "";
  return `<pivotField${axis}${data ? ' dataField="1"' : ""} showAll="0"><items count="1"><item t="default"/></items></pivotField>`;
}

function pivotTableXml(config, name, cacheId, location) {
  const rowFields = config.rows.length ? `<rowFields count="${config.rows.length}">${config.rows.map((index) => `<field x="${index}"/>`).join("")}</rowFields>` : "";
  const columnIndexes = [...config.columns, ...(config.values.length > 1 ? [-2] : [])];
  const columnFields = columnIndexes.length ? `<colFields count="${columnIndexes.length}">${columnIndexes.map((index) => `<field x="${index}"/>`).join("")}</colFields>` : "";
  const pageFields = config.filters.length ? `<pageFields count="${config.filters.length}">${config.filters.map((index) => `<pageField fld="${index}" item="0" hier="-1"/>`).join("")}</pageFields>` : "";
  const dataFields = `<dataFields count="${config.values.length}">${config.values.map((value) => `<dataField name="${escapeXml(value.name)}" fld="${value.field}" subtotal="${value.summarize}"/>`).join("")}</dataFields>`;
  return `${XML_HEADER}<pivotTableDefinition xmlns="${MAIN_NS}" xmlns:r="${REL_NS}" name="${escapeXml(name)}" cacheId="${cacheId}" dataCaption="Values" grandTotalCaption="Grand Total" updatedVersion="8" minRefreshableVersion="3" createdVersion="3" useAutoFormatting="1" itemPrintTitles="1" indent="0" compact="1" compactData="1" rowGrandTotals="${config.showGrandTotals ? 1 : 0}" colGrandTotals="${config.showGrandTotals ? 1 : 0}"><location ref="${escapeXml(location)}" firstHeaderRow="1" firstDataRow="1" firstDataCol="${config.rows.length}"/><pivotFields count="${config.fields.length}">${config.fields.map((_, index) => pivotFieldXml(config, index)).join("")}</pivotFields>${rowFields}${columnFields}${pageFields}${dataFields}<pivotTableStyleInfo name="${escapeXml(config.style)}" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>`;
}

function addPivotCacheToWorkbook(state, cacheId, relationshipId) {
  let xml = getTextFile(state, state.workbookPath);
  const item = `<pivotCache cacheId="${cacheId}" r:id="${relationshipId}"/>`;
  if (/<(?:\w+:)?pivotCaches\b/i.test(xml)) xml = xml.replace(/<\/(?:\w+:)?pivotCaches\s*>/i, `${item}</pivotCaches>`);
  else if (/<(?:\w+:)?calcPr\b/i.test(xml)) {
    xml = xml.replace(/(<(?:\w+:)?calcPr\b[^>]*(?:\/>|>[\s\S]*?<\/(?:\w+:)?calcPr>))/i, `$1<pivotCaches>${item}</pivotCaches>`);
  } else xml = xml.replace(/<\/(?:\w+:)?workbook\s*>/i, `<pivotCaches>${item}</pivotCaches></workbook>`);
  setTextFile(state, state.workbookPath, xml);
}

function addPivotToWorksheet(state, sheetPath, relationshipId) {
  let xml = getTextFile(state, sheetPath);
  const item = `<pivotTablePart r:id="${relationshipId}"/>`;
  const container = allTags(xml, "pivotTableParts")[0];
  if (container) {
    const count = Number(container.attributes.count ?? allTags(container.content, "pivotTablePart").length) + 1;
    xml = xml.replace(container.xml, container.xml.replace(/\bcount="\d+"/i, `count="${count}"`).replace(/<\/(?:\w+:)?pivotTableParts\s*>/i, `${item}</pivotTableParts>`));
  } else xml = xml.replace(/<\/(?:\w+:)?worksheet\s*>/i, `<pivotTableParts count="1">${item}</pivotTableParts></worksheet>`);
  setTextFile(state, sheetPath, xml);
}

function pivotEntries(state, workbook) {
  const entries = [];
  for (const sheetName of workbook.SheetNames) {
    const sheetPath = state.sheetPaths.get(sheetName);
    const sheetXml = getTextFile(state, sheetPath);
    for (const part of allTags(sheetXml, "pivotTablePart")) {
      const relationshipId = part.attributes["r:id"];
      const partPath = relationshipTarget(state, sheetPath, relationshipId);
      if (!partPath) continue;
      const xml = getTextFile(state, partPath);
      const root = allTags(xml, "pivotTableDefinition")[0]?.attributes ?? {};
      const cacheRelationship = readRelationships(state, partPath).items.find((item) => item.type === CACHE_DEF_REL);
      const cachePath = cacheRelationship ? relationshipTarget(state, partPath, cacheRelationship.id) : undefined;
      const cacheXml = cachePath ? getTextFile(state, cachePath) : "";
      const source = allTags(cacheXml, "worksheetSource")[0]?.attributes ?? {};
      const cacheFields = allTags(cacheXml, "cacheField").map((item) => item.attributes.name);
      const fieldNames = (container, attribute = "x") => allTags(allTags(xml, container)[0]?.content ?? "", container === "pageFields" ? "pageField" : "field").map((item) => cacheFields[Number(item.attributes[attribute])]).filter(Boolean);
      const values = allTags(allTags(xml, "dataFields")[0]?.content ?? "", "dataField").map((item) => ({ field: cacheFields[Number(item.attributes.fld)], summarize: item.attributes.subtotal ?? "sum", name: item.attributes.name }));
      entries.push({
        id: partPath,
        name: root.name,
        cacheId: Number(root.cacheId),
        source: { sheet: source.sheet, range: source.ref },
        target: { sheet: sheetName, cell: allTags(xml, "location")[0]?.attributes.ref?.split(":")[0] },
        rows: fieldNames("rowFields"),
        columns: fieldNames("colFields").filter(Boolean),
        filters: fieldNames("pageFields", "fld"),
        values,
        style: allTags(xml, "pivotTableStyleInfo")[0]?.attributes.name,
        sheetPath,
        relationshipId,
        cachePath,
        cacheRelationshipId: cacheRelationship?.id
      });
    }
  }
  return entries;
}

export class PivotCollection {
  constructor(workbook) {
    this.workbook = workbook;
  }

  list(sheet) {
    const state = ensurePackageState(this.workbook);
    return pivotEntries(state, this.workbook).filter((item) => sheet === undefined || item.target.sheet === sheet || this.workbook.SheetNames[sheet] === item.target.sheet).map(({ sheetPath, relationshipId, cacheRelationshipId, ...item }) => item);
  }

  add(config) {
    const normalized = normalizeConfig(this.workbook, config);
    let state = ensurePackageState(this.workbook);
    const name = normalized.name ?? `PivotTable${pivotEntries(state, this.workbook).length + 1}`;
    if (pivotEntries(state, this.workbook).some((item) => item.name === name)) throw new RangeError(`Pivot table ${name} already exists.`);
    const location = writeRenderedCells(this.workbook, state, normalized);
    const cachePath = nextPartPath(state, "xl/pivotCache", "pivotCacheDefinition");
    const recordsPath = nextPartPath(state, "xl/pivotCache", "pivotCacheRecords");
    const tablePath = nextPartPath(state, "xl/pivotTables", "pivotTable");
    const recordRelationshipId = addRelationship(state, cachePath, CACHE_RECORDS_REL, recordsPath);
    setTextFile(state, cachePath, cacheDefinitionXml(normalized, recordRelationshipId));
    setTextFile(state, recordsPath, cacheRecordsXml(normalized));
    const tableCacheRelationshipId = addRelationship(state, tablePath, CACHE_DEF_REL, cachePath);
    const workbookRelationships = readRelationships(state, state.workbookPath).items;
    const cacheIds = allTags(getTextFile(state, state.workbookPath), "pivotCache").map((item) => Number(item.attributes.cacheId)).filter(Number.isFinite);
    const cacheId = cacheIds.length ? Math.max(...cacheIds) + 1 : 1;
    const workbookCacheRelationshipId = addRelationship(state, state.workbookPath, CACHE_DEF_REL, cachePath);
    addPivotCacheToWorkbook(state, cacheId, workbookCacheRelationshipId);
    setTextFile(state, tablePath, pivotTableXml(normalized, name, cacheId, location));
    const sheetPath = state.sheetPaths.get(normalized.target.sheet);
    const sheetRelationshipId = addRelationship(state, sheetPath, TABLE_REL, tablePath);
    addPivotToWorksheet(state, sheetPath, sheetRelationshipId);
    addContentTypeOverride(state, cachePath, "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml");
    addContentTypeOverride(state, recordsPath, "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml");
    addContentTypeOverride(state, tablePath, "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml");
    return this.list(normalized.target.sheet).find((item) => item.id === tablePath);
  }

  update(reference, changes) {
    const current = this.list().find((item) => item.id === reference || item.name === reference);
    if (!current) throw new RangeError(`Pivot table ${String(reference)} does not exist.`);
    const config = {
      name: changes.name ?? current.name,
      source: changes.source ?? current.source,
      target: changes.target ?? current.target,
      rows: changes.rows ?? current.rows,
      columns: changes.columns ?? current.columns,
      filters: changes.filters ?? current.filters,
      values: changes.values ?? current.values,
      style: changes.style ?? current.style,
      refreshOnLoad: changes.refreshOnLoad,
      showGrandTotals: changes.showGrandTotals
    };
    this.remove(reference);
    return this.add(config);
  }

  remove(reference) {
    const state = ensurePackageState(this.workbook);
    const entry = pivotEntries(state, this.workbook).find((item) => item.id === reference || item.name === reference);
    if (!entry) return false;
    let sheetXml = getTextFile(state, entry.sheetPath);
    const container = allTags(sheetXml, "pivotTableParts")[0];
    if (container) {
      const kept = allTags(container.content, "pivotTablePart").filter((item) => item.attributes["r:id"] !== entry.relationshipId);
      const replacement = kept.length ? `<pivotTableParts count="${kept.length}">${kept.map((item) => item.xml).join("")}</pivotTableParts>` : "";
      sheetXml = sheetXml.replace(container.xml, replacement);
      setTextFile(state, entry.sheetPath, sheetXml);
    }
    removeRelationship(state, entry.sheetPath, entry.relationshipId);
    const workbookXml = getTextFile(state, state.workbookPath);
    const caches = allTags(workbookXml, "pivotCaches")[0];
    if (caches) {
      const cacheItems = allTags(caches.content, "pivotCache");
      const removedCache = cacheItems.find((item) => Number(item.attributes.cacheId) === entry.cacheId);
      const kept = cacheItems.filter((item) => item !== removedCache);
      setTextFile(state, state.workbookPath, workbookXml.replace(caches.xml, kept.length ? `<pivotCaches>${kept.map((item) => item.xml).join("")}</pivotCaches>` : ""));
      if (removedCache?.attributes["r:id"]) removeRelationship(state, state.workbookPath, removedCache.attributes["r:id"]);
    }
    const cacheRecordsRelationship = entry.cachePath ? readRelationships(state, entry.cachePath).items.find((item) => item.type === CACHE_RECORDS_REL) : undefined;
    const recordsPath = cacheRecordsRelationship ? relationshipTarget(state, entry.cachePath, cacheRecordsRelationship.id) : undefined;
    for (const path of [entry.id, relationshipsPath(entry.id), entry.cachePath, entry.cachePath ? relationshipsPath(entry.cachePath) : undefined, recordsPath].filter(Boolean)) state.files.delete(path);
    for (const path of [entry.id, entry.cachePath, recordsPath].filter(Boolean)) removeContentTypeOverride(state, path);
    return true;
  }
}
