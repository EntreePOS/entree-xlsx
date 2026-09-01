import { posix } from "node:path";
import { decodeAddress, decodeRange, encodeAddress, encodeRange, expandWorksheetRef } from "./address.js";
import { PackageState, attachPackageState, getPackageState } from "./package-state.js";
import { createZip, extractZip } from "./zip.js";
import { StyleCollection } from "./style.js";
import { allTags, decodeXml, escapeXml, parseAttributes, tagContent } from "./xml.js";

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const builtinFormats = new Map([
  [0, "General"], [1, "0"], [2, "0.00"], [3, "#,##0"], [4, "#,##0.00"],
  [9, "0%"], [10, "0.00%"], [11, "0.00E+00"], [14, "m/d/yy"], [15, "d-mmm-yy"],
  [16, "d-mmm"], [17, "mmm-yy"], [18, "h:mm AM/PM"], [19, "h:mm:ss AM/PM"],
  [20, "h:mm"], [21, "h:mm:ss"], [22, "m/d/yy h:mm"], [49, "@"]
]);

function stable(value) {
  if (Array.isArray(value)) return value.map(stable);
  if (value && typeof value === "object") {
    return Object.fromEntries(Object.keys(value).sort().map((key) => [key, stable(value[key])]));
  }
  return value;
}

function styleKey(style) {
  return JSON.stringify(stable(style ?? {}));
}

function argb(value) {
  if (!value) return undefined;
  const raw = typeof value === "string" ? value : value.rgb;
  if (!raw) return undefined;
  const clean = raw.replace(/^#/, "").toUpperCase();
  return clean.length === 6 ? `FF${clean}` : clean;
}

function colorXml(tag, value) {
  if (!value) return "";
  if (typeof value === "string" || value.rgb) return `<${tag} rgb="${escapeXml(argb(value))}"${value.tint !== undefined ? ` tint="${value.tint}"` : ""}/>`;
  if (value.theme !== undefined) return `<${tag} theme="${value.theme}"${value.tint !== undefined ? ` tint="${value.tint}"` : ""}/>`;
  if (value.indexed !== undefined) return `<${tag} indexed="${value.indexed}"/>`;
  if (value.auto !== undefined) return `<${tag} auto="${value.auto ? 1 : 0}"/>`;
  return "";
}

function parseColor(content, tag = "color") {
  const match = new RegExp(`<(?:\\w+:)?${tag}\\b([^>]*)\\/>`, "i").exec(content);
  if (!match) return undefined;
  const attrs = parseAttributes(match[1]);
  if (attrs.rgb) return { rgb: attrs.rgb.length === 8 && attrs.rgb.startsWith("FF") ? attrs.rgb.slice(2) : attrs.rgb, ...(attrs.tint !== undefined ? { tint: Number(attrs.tint) } : {}) };
  if (attrs.theme !== undefined) return { theme: Number(attrs.theme), ...(attrs.tint !== undefined ? { tint: Number(attrs.tint) } : {}) };
  if (attrs.indexed !== undefined) return { indexed: Number(attrs.indexed) };
  if (attrs.auto !== undefined) return { auto: /^(?:1|true)$/i.test(attrs.auto) };
  return undefined;
}

function fontXml(font = {}) {
  const flag = (tag, value) => value ? `<${tag}/>` : "";
  return `<font>${flag("b", font.bold)}${flag("i", font.italic)}${font.underline ? `<u${font.underline === "single" || font.underline === true ? "" : ` val="${escapeXml(font.underline)}"`}/>` : ""}${flag("strike", font.strike)}${flag("outline", font.outline)}${flag("shadow", font.shadow)}${flag("condense", font.condense)}${flag("extend", font.extend)}${font.size !== undefined ? `<sz val="${font.size}"/>` : ""}${colorXml("color", font.color)}${font.name ? `<name val="${escapeXml(font.name)}"/>` : ""}${font.family !== undefined ? `<family val="${font.family}"/>` : ""}${font.charset !== undefined ? `<charset val="${font.charset}"/>` : ""}${font.scheme && font.scheme !== "none" ? `<scheme val="${escapeXml(font.scheme)}"/>` : ""}${font.verticalAlign && font.verticalAlign !== "baseline" ? `<vertAlign val="${escapeXml(font.verticalAlign)}"/>` : ""}</font>`;
}

function fillXml(fill = {}) {
  if (fill.type === "gradient" || fill.stops) {
    const attributes = fill.gradientType === "path"
      ? ` type="path"${["left", "right", "top", "bottom"].filter((key) => fill[key] !== undefined).map((key) => ` ${key}="${fill[key]}"`).join("")}`
      : ` degree="${fill.degree ?? 0}"`;
    const stops = (fill.stops ?? []).map((stop) => `<stop position="${stop.position}">${colorXml("color", stop.color)}</stop>`).join("");
    return `<fill><gradientFill${attributes}>${stops}</gradientFill></fill>`;
  }
  const foreground = fill.foreground ?? fill.color;
  const background = fill.background;
  const patternType = fill.patternType ?? (foreground ? "solid" : "none");
  return `<fill><patternFill patternType="${escapeXml(patternType)}">${colorXml("fgColor", foreground)}${colorXml("bgColor", background) || (foreground ? '<bgColor indexed="64"/>' : "")}</patternFill></fill>`;
}

function borderSideXml(name, side = {}) {
  return `<${name}${side.style ? ` style="${escapeXml(side.style)}"` : ""}>${colorXml("color", side.color)}</${name}>`;
}

function borderXml(border = {}) {
  const attributes = ["diagonalUp", "diagonalDown", "outline"].filter((key) => border[key] !== undefined).map((key) => ` ${key}="${border[key] ? 1 : 0}"`).join("");
  const standard = ["left", "right", "top", "bottom", "diagonal"].map((side) => borderSideXml(side, border[side])).join("");
  const interior = ["vertical", "horizontal"].filter((side) => border[side]).map((side) => borderSideXml(side, border[side])).join("");
  return `<border${attributes}>${standard}${interior}</border>`;
}

function styleChildXml(style) {
  const attributes = (values) => Object.entries(values).map(([key, value]) => ` ${key}="${typeof value === "boolean" ? value ? 1 : 0 : escapeXml(value)}"`).join("");
  const alignment = style.alignment && Object.keys(style.alignment).length ? `<alignment${attributes(style.alignment)}/>` : "";
  const protection = style.protection && Object.keys(style.protection).length ? `<protection${attributes(style.protection)}/>` : "";
  return { alignment, protection };
}

function xfXml(style, ids, options = {}) {
  const { alignment, protection } = styleChildXml(style);
  const attributes = `numFmtId="${ids.numFmtId}" fontId="${ids.fontId}" fillId="${ids.fillId}" borderId="${ids.borderId}"${options.cell ? ` xfId="${options.xfId ?? 0}"` : ""}${style.font ? ' applyFont="1"' : ""}${style.fill ? ' applyFill="1"' : ""}${style.border ? ' applyBorder="1"' : ""}${ids.numFmtId ? ' applyNumberFormat="1"' : ""}${alignment ? ' applyAlignment="1"' : ""}${protection ? ' applyProtection="1"' : ""}`;
  return alignment || protection ? `<xf ${attributes}>${alignment}${protection}</xf>` : `<xf ${attributes}/>`;
}

class StyleRegistry {
  constructor(workbook) {
    this.workbook = workbook;
    this.styles = [{ style: {}, namedStyle: undefined }];
    this.styleIndexes = new Map([[styleKey([{}, undefined]), 0]]);
    const collection = new StyleCollection(workbook);
    this.namedStyles = collection.names.filter((name) => name !== "Normal").map((name) => ({ name, style: collection.get(name) }));
    this.namedIndexes = new Map(this.namedStyles.map((entry, index) => [entry.name, index + 1]));
  }

  index(style, namedStyle) {
    const normalized = style ?? {};
    const name = this.namedIndexes.has(namedStyle) ? namedStyle : undefined;
    const key = styleKey([normalized, name]);
    if (this.styleIndexes.has(key)) return this.styleIndexes.get(key);
    const index = this.styles.length;
    this.styles.push({ style: normalized, namedStyle: name });
    this.styleIndexes.set(key, index);
    return index;
  }

  toXml() {
    const fonts = [{ name: "Calibri", size: 11, family: 2, scheme: "minor", color: { theme: 1 } }];
    const fills = [{ type: "pattern", patternType: "none" }, { type: "pattern", patternType: "gray125" }];
    const borders = [{}];
    const fontMap = new Map([[styleKey(fonts[0]), 0]]);
    const fillMap = new Map(fills.map((fill, index) => [styleKey(fill), index]));
    const borderMap = new Map([[styleKey({}), 0]]);
    const formats = new Map();
    let nextFormat = 164;

    const getIndex = (collection, map, value) => {
      if (!value) return 0;
      const key = styleKey(value);
      if (map.has(key)) return map.get(key);
      const index = collection.length;
      collection.push(value);
      map.set(key, index);
      return index;
    };
    const idsFor = (style) => {
      let numFmtId = 0;
      if (style.numberFormat && style.numberFormat !== "General") {
        const builtin = [...builtinFormats].find(([, format]) => format === style.numberFormat);
        if (builtin) numFmtId = builtin[0];
        else {
          if (!formats.has(style.numberFormat)) formats.set(style.numberFormat, nextFormat++);
          numFmtId = formats.get(style.numberFormat);
        }
      }
      return {
        fontId: getIndex(fonts, fontMap, style.font),
        fillId: getIndex(fills, fillMap, style.fill),
        borderId: getIndex(borders, borderMap, style.border),
        numFmtId
      };
    };
    const namedXfs = [xfXml({}, { fontId: 0, fillId: 0, borderId: 0, numFmtId: 0 })];
    for (const entry of this.namedStyles) namedXfs.push(xfXml(entry.style, idsFor(entry.style)));
    const xfs = this.styles.map(({ style, namedStyle }) => xfXml(style, idsFor(style), { cell: true, xfId: this.namedIndexes.get(namedStyle) ?? 0 }));
    const numFormats = formats.size
      ? `<numFmts count="${formats.size}">${[...formats].map(([format, id]) => `<numFmt numFmtId="${id}" formatCode="${escapeXml(format)}"/>`).join("")}</numFmts>`
      : "";
    const cellStyles = [`<cellStyle name="Normal" xfId="0" builtinId="0"/>`, ...this.namedStyles.map((entry, index) => `<cellStyle name="${escapeXml(entry.name)}" xfId="${index + 1}"/>`)].join("");
    return `${XML_HEADER}<styleSheet xmlns="${MAIN_NS}">${numFormats}<fonts count="${fonts.length}">${fonts.map(fontXml).join("")}</fonts><fills count="${fills.length}">${fills.map(fillXml).join("")}</fills><borders count="${borders.length}">${borders.map(borderXml).join("")}</borders><cellStyleXfs count="${namedXfs.length}">${namedXfs.join("")}</cellStyleXfs><cellXfs count="${xfs.length}">${xfs.join("")}</cellXfs><cellStyles count="${this.namedStyles.length + 1}">${cellStyles}</cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>`;
  }
}

function excelSerial(date) {
  return (date.getTime() - Date.UTC(1899, 11, 30)) / 86_400_000;
}

function dateFromSerial(value) {
  return new Date(Date.UTC(1899, 11, 30) + value * 86_400_000);
}

function cellXml(address, cell, styles) {
  let style = cell.style ?? {};
  if (cell.type === "date" && !style.numberFormat) style = { ...style, numberFormat: "yyyy-mm-dd hh:mm:ss" };
  const styleIndex = styles.index(style, cell.namedStyle);
  const styleAttribute = styleIndex ? ` s="${styleIndex}"` : "";
  const formula = cell.formula ? `<f>${escapeXml(cell.formula)}</f>` : "";
  if (cell.type === "blank" || cell.value === undefined || cell.value === null) {
    return formula || styleIndex ? `<c r="${address}"${styleAttribute}>${formula}</c>` : "";
  }
  if (cell.type === "string") {
    if (cell.formula) return `<c r="${address}"${styleAttribute} t="str">${formula}<v>${escapeXml(cell.value)}</v></c>`;
    const preserve = /^\s|\s$/.test(cell.value) ? ' xml:space="preserve"' : "";
    return `<c r="${address}"${styleAttribute} t="inlineStr"><is><t${preserve}>${escapeXml(cell.value)}</t></is></c>`;
  }
  if (cell.type === "boolean") return `<c r="${address}"${styleAttribute} t="b">${formula}<v>${cell.value ? 1 : 0}</v></c>`;
  const value = cell.type === "date" ? excelSerial(cell.value) : cell.value;
  return `<c r="${address}"${styleAttribute}>${formula}<v>${escapeXml(value)}</v></c>`;
}

function columnXml(column, index, styles) {
  if (!column) return "";
  const styleIndex = column.style ? styles.index(column.style, column.namedStyle) : Number(column.styleIndex ?? 0);
  if (column.style) column.styleIndex = styleIndex;
  const width = column.width !== undefined ? ` width="${column.width}" customWidth="1"` : "";
  const hidden = column.hidden ? ' hidden="1"' : "";
  const style = styleIndex ? ` style="${styleIndex}"` : "";
  return `<col min="${index + 1}" max="${index + 1}"${width}${hidden}${style}/>`;
}

function sheetXml(worksheet, styles) {
  const addresses = Object.keys(worksheet)
    .filter((key) => /^[A-Z]+[1-9]\d*$/.test(key))
    .sort((left, right) => {
      const a = decodeAddress(left);
      const b = decodeAddress(right);
      return a.r - b.r || a.c - b.c;
    });
  const rows = new Map();
  for (const address of addresses) {
    const point = decodeAddress(address);
    const xml = cellXml(address, worksheet[address], styles);
    if (xml) {
      if (!rows.has(point.r)) rows.set(point.r, []);
      rows.get(point.r).push(xml);
    }
  }
  for (const [index, properties] of (worksheet["!rows"] ?? []).entries()) {
    if (properties && !rows.has(index)) rows.set(index, []);
  }
  const rowXml = [...rows].sort(([a], [b]) => a - b).map(([row, cells]) => {
    const properties = worksheet["!rows"]?.[row] ?? {};
    const styleIndex = properties.style ? styles.index(properties.style, properties.namedStyle) : Number(properties.styleIndex ?? 0);
    if (properties.style) properties.styleIndex = styleIndex;
    const height = properties.height ? ` ht="${properties.height}" customHeight="1"` : "";
    const hidden = properties.hidden ? ' hidden="1"' : "";
    const style = styleIndex ? ` s="${styleIndex}" customFormat="1"` : "";
    return `<row r="${row + 1}"${height}${hidden}${style}>${cells.join("")}</row>`;
  }).join("");
  const columns = (worksheet["!cols"] ?? []).map((column, index) => columnXml(column, index, styles)).join("");
  const merges = worksheet["!merges"]?.length
    ? `<mergeCells count="${worksheet["!merges"].length}">${worksheet["!merges"].map((range) => `<mergeCell ref="${encodeRange(range)}"/>`).join("")}</mergeCells>`
    : "";
  const hyperlinks = [];
  const relationships = [];
  for (const address of addresses) {
    const link = worksheet[address]?.hyperlink;
    if (!link) continue;
    if (link.target.startsWith("#")) hyperlinks.push(`<hyperlink ref="${address}" location="${escapeXml(link.target.slice(1))}"${link.tooltip ? ` tooltip="${escapeXml(link.tooltip)}"` : ""}/>`);
    else {
      const id = `rId${relationships.length + 1}`;
      hyperlinks.push(`<hyperlink ref="${address}" r:id="${id}"${link.tooltip ? ` tooltip="${escapeXml(link.tooltip)}"` : ""}/>`);
      relationships.push({ id, target: link.target });
    }
  }
  const hyperlinkXml = hyperlinks.length ? `<hyperlinks>${hyperlinks.join("")}</hyperlinks>` : "";
  const filterXml = worksheet["!autofilter"] ? `<autoFilter ref="${escapeXml(worksheet["!autofilter"])}"/>` : "";
  const protectionXml = worksheet["!protection"] ? `<sheetProtection${Object.entries(worksheet["!protection"]).map(([key, value]) => ` ${key}="${typeof value === "boolean" ? value ? 1 : 0 : escapeXml(value)}"`).join("")}/>` : "";
  const xml = `${XML_HEADER}<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><dimension ref="${escapeXml(worksheet["!ref"] ?? "A1")}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>${columns ? `<cols>${columns}</cols>` : ""}<sheetData>${rowXml}</sheetData>${protectionXml}${filterXml}${merges}${hyperlinkXml}<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`;
  return { xml, relationships };
}

function relationshipsXml(items) {
  return `${XML_HEADER}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items.map((item) => `<Relationship Id="${item.id}" Type="${item.type}" Target="${escapeXml(item.target)}"${item.external ? ' TargetMode="External"' : ""}/>`).join("")}</Relationships>`;
}

function workbookProtectionXml(protection) {
  return protection ? `<workbookProtection${Object.entries(protection).map(([key, value]) => ` ${key}="${typeof value === "boolean" ? value ? 1 : 0 : escapeXml(value)}"`).join("")}/>` : "";
}

function writeNewXlsx(workbook) {
  const styles = new StyleRegistry(workbook);
  const files = {};
  const sheetOverrides = [];
  const workbookRelationships = [];
  const sheetEntries = [];

  workbook.SheetNames.forEach((name, index) => {
    const path = `xl/worksheets/sheet${index + 1}.xml`;
    const output = sheetXml(workbook.Sheets[name], styles);
    files[path] = output.xml;
    if (output.relationships.length) {
      files[`xl/worksheets/_rels/sheet${index + 1}.xml.rels`] = relationshipsXml(output.relationships.map((relationship) => ({
        ...relationship,
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        external: true
      })));
    }
    sheetOverrides.push(`<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
    workbookRelationships.push({ id: `rId${index + 1}`, type: `${REL_NS}/worksheet`, target: `worksheets/sheet${index + 1}.xml` });
    sheetEntries.push(`<sheet name="${escapeXml(name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`);
  });

  const stylesId = workbookRelationships.length + 1;
  workbookRelationships.push({ id: `rId${stylesId}`, type: `${REL_NS}/styles`, target: "styles.xml" });
  files["xl/styles.xml"] = styles.toXml();
  files["xl/workbook.xml"] = `${XML_HEADER}<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">${workbookProtectionXml(workbook.WorkbookProtection)}<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="24000" windowHeight="12000"/></bookViews><sheets>${sheetEntries.join("")}</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`;
  files["xl/_rels/workbook.xml.rels"] = relationshipsXml(workbookRelationships);
  files["_rels/.rels"] = relationshipsXml([
    { id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", target: "xl/workbook.xml" },
    { id: "rId2", type: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", target: "docProps/core.xml" },
    { id: "rId3", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", target: "docProps/app.xml" }
  ]);
  const properties = workbook.Props ?? {};
  const now = new Date().toISOString();
  files["docProps/core.xml"] = `${XML_HEADER}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${escapeXml(properties.title ?? "")}</dc:title><dc:creator>${escapeXml(properties.author ?? "Entree POS")}</dc:creator><cp:lastModifiedBy>${escapeXml(properties.lastAuthor ?? properties.author ?? "Entree POS")}</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${escapeXml(properties.createdAt instanceof Date ? properties.createdAt.toISOString() : properties.createdAt ?? now)}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`;
  files["docProps/app.xml"] = `${XML_HEADER}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>@entree_pos/xlsx</Application><AppVersion>1.0</AppVersion></Properties>`;
  files["[Content_Types].xml"] = `${XML_HEADER}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>${sheetOverrides.join("")}</Types>`;
  return createZip(files);
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function preservedCellXml(address, cell) {
  if (!cell) return "";
  const styleIndex = Number(cell.styleIndex ?? 0);
  const styleAttribute = styleIndex ? ` s="${styleIndex}"` : "";
  const formula = cell.formula ? `<f>${escapeXml(cell.formula)}</f>` : "";
  if (cell.type === "blank" || cell.value === undefined || cell.value === null) {
    return formula || styleIndex ? `<c r="${address}"${styleAttribute}>${formula}</c>` : "";
  }
  if (cell.type === "string") {
    if (cell.formula) return `<c r="${address}"${styleAttribute} t="str">${formula}<v>${escapeXml(cell.value)}</v></c>`;
    const preserve = /^\s|\s$/.test(cell.value) ? ' xml:space="preserve"' : "";
    return `<c r="${address}"${styleAttribute} t="inlineStr"><is><t${preserve}>${escapeXml(cell.value)}</t></is></c>`;
  }
  if (cell.type === "boolean") return `<c r="${address}"${styleAttribute} t="b">${formula}<v>${cell.value ? 1 : 0}</v></c>`;
  const value = cell.type === "date" ? excelSerial(cell.value) : cell.value;
  return `<c r="${address}"${styleAttribute}>${formula}<v>${escapeXml(value)}</v></c>`;
}

function patchDimension(xml, reference) {
  if (!reference) return xml;
  const dimension = /<(?:\w+:)?dimension\b[^>]*\bref="[^"]*"[^>]*\/?\s*>/i;
  if (dimension.test(xml)) return xml.replace(dimension, (tag) => tag.replace(/\bref="[^"]*"/i, `ref="${escapeXml(reference)}"`));
  const sheetData = /<(?:\w+:)?sheetData\b/i;
  return sheetData.test(xml) ? xml.replace(sheetData, `<dimension ref="${escapeXml(reference)}"/><sheetData`) : xml;
}

function patchCell(xml, address, cell) {
  const escapedAddress = escapeRegExp(address);
  const existing = new RegExp(`<(?:\\w+:)?c\\b(?=[^>]*\\br="${escapedAddress}"(?:\\s|/|>))[^>]*(?:\\/>|>[\\s\\S]*?<\\/(?:\\w+:)?c\\s*>)`, "i");
  const replacement = preservedCellXml(address, cell);
  if (existing.test(xml)) return xml.replace(existing, replacement);
  if (!replacement) return xml;

  const rowNumber = decodeAddress(address).r + 1;
  const row = new RegExp(`(<(?:\\w+:)?row\\b(?=[^>]*\\br="${rowNumber}"(?:\\s|/|>))[^>]*>)([\\s\\S]*?)(<\\/(?:\\w+:)?row\\s*>)`, "i");
  if (row.test(xml)) return xml.replace(row, `$1$2${replacement}$3`);
  const emptyRow = new RegExp(`<(?:\\w+:)?row\\b(?=[^>]*\\br="${rowNumber}"(?:\\s|/|>))([^>]*)\\/>`, "i");
  if (emptyRow.test(xml)) return xml.replace(emptyRow, (tag) => `${tag.replace(/\/>$/, ">")}${replacement}</row>`);
  const sheetDataEnd = /<\/(?:\w+:)?sheetData\s*>/i;
  if (!sheetDataEnd.test(xml)) throw new Error("Invalid XLSX worksheet: missing sheetData.");
  return xml.replace(sheetDataEnd, `<row r="${rowNumber}">${replacement}</row></sheetData>`);
}

function sortSheetData(xml) {
  const sheetData = allTags(xml, "sheetData")[0];
  if (!sheetData?.content) return xml;
  const rows = allTags(sheetData.content, "row");
  if (!rows.length) return xml;

  const normalizedRows = rows.map((row) => {
    if (!row.content) return row;
    const cells = allTags(row.content, "c");
    if (cells.length < 2) return row;
    const sortedCells = [...cells].sort((left, right) => {
      const leftPoint = decodeAddress(left.attributes.r);
      const rightPoint = decodeAddress(right.attributes.r);
      return leftPoint.c - rightPoint.c;
    });
    let otherContent = row.content;
    for (const cell of cells) otherContent = otherContent.replace(cell.xml, "");
    const content = `${otherContent}${sortedCells.map((cell) => cell.xml).join("")}`;
    return { ...row, xml: row.xml.replace(row.content, content) };
  }).sort((left, right) => Number(left.attributes.r) - Number(right.attributes.r));

  let otherContent = sheetData.content;
  for (const row of rows) otherContent = otherContent.replace(row.xml, "");
  const content = `${otherContent}${normalizedRows.map((row) => row.xml).join("")}`;
  return xml.replace(sheetData.xml, sheetData.xml.replace(sheetData.content, content));
}

function replaceProperty(xml, qualifiedName, value, attributes = "") {
  const escapedName = escapeRegExp(qualifiedName);
  const pattern = new RegExp(`<${escapedName}\\b[^>]*>[\\s\\S]*?<\\/${escapedName}>`, "i");
  const replacement = `<${qualifiedName}${attributes}>${escapeXml(value ?? "")}</${qualifiedName}>`;
  if (pattern.test(xml)) return xml.replace(pattern, replacement);
  return xml.replace(/<\/(?:\w+:)?coreProperties\s*>/i, `${replacement}</cp:coreProperties>`);
}

function patchCoreProperties(xml, properties) {
  let output = xml;
  output = replaceProperty(output, "dc:title", properties.title ?? "");
  output = replaceProperty(output, "dc:creator", properties.author ?? "Entree POS");
  output = replaceProperty(output, "cp:lastModifiedBy", properties.lastAuthor ?? properties.author ?? "Entree POS");
  if (properties.createdAt) {
    const created = properties.createdAt instanceof Date ? properties.createdAt.toISOString() : properties.createdAt;
    output = replaceProperty(output, "dcterms:created", created, ' xsi:type="dcterms:W3CDTF"');
  }
  return replaceProperty(output, "dcterms:modified", new Date().toISOString(), ' xsi:type="dcterms:W3CDTF"');
}

function appendCountedElement(xml, containerName, element) {
  const container = allTags(xml, containerName)[0];
  if (!container) throw new Error(`Invalid styles part: missing ${containerName}.`);
  const previousCount = Number(container.attributes.count ?? allTags(container.content, containerName === "cellStyles" ? "cellStyle" : containerName === "cellXfs" || containerName === "cellStyleXfs" ? "xf" : containerName.slice(0, -1)).length);
  const count = previousCount + 1;
  const counted = /\bcount="\d+"/i.test(container.xml)
    ? container.xml.replace(/\bcount="\d+"/i, `count="${count}"`)
    : container.xml.replace(/^(<[^>]+)/, `$1 count="${count}"`);
  const updated = counted
    .replace(new RegExp(`<\\/(?:\\w+:)?${containerName}\\s*>$`, "i"), `${element}</${containerName}>`);
  return { xml: xml.replace(container.xml, updated), index: count - 1 };
}

function removeCountedElement(xml, containerName, elementXml) {
  const container = allTags(xml, containerName)[0];
  if (!container || !container.xml.includes(elementXml)) return xml;
  const previousCount = Number(container.attributes.count ?? 1);
  let updated = container.xml.replace(elementXml, "");
  updated = /\bcount="\d+"/i.test(updated)
    ? updated.replace(/\bcount="\d+"/i, `count="${Math.max(0, previousCount - 1)}"`)
    : updated;
  return xml.replace(container.xml, updated);
}

class PreservedStyleRegistry {
  constructor(xml) {
    this.xml = xml;
    this.fonts = allTags(tagContent(xml, "fonts") ?? "", "font").map(({ content }) => parseFont(content));
    this.fills = allTags(tagContent(xml, "fills") ?? "", "fill").map(({ content }) => parseFill(content));
    this.borders = allTags(tagContent(xml, "borders") ?? "", "border").map(({ content, attributes }) => parseBorder(content, attributes));
    this.fontMap = new Map(this.fonts.map((value, index) => [styleKey(value), index]));
    this.fillMap = new Map(this.fills.map((value, index) => [styleKey(value), index]));
    this.borderMap = new Map(this.borders.map((value, index) => [styleKey(value), index]));
    this.formats = new Map(builtinFormats);
    for (const { attributes } of allTags(tagContent(xml, "numFmts") ?? "", "numFmt")) this.formats.set(Number(attributes.numFmtId), attributes.formatCode);
    this.formatIds = new Map([...this.formats].map(([id, format]) => [format, id]));
    this.cellRecords = parseStyleRecords(xml, "cellXfs");
    this.cellMap = new Map(this.cellRecords.map((record, index) => [this.#formatKey(record.style, record.xfId), index]));
    this.masterRecords = parseStyleRecords(xml, "cellStyleXfs");
    this.masterMap = new Map(this.masterRecords.map((record, index) => [styleKey(record.style), index]));
    this.namedIndexes = new Map(allTags(tagContent(xml, "cellStyles") ?? "", "cellStyle").map(({ attributes }) => [attributes.name, Number(attributes.xfId)]));
  }

  ensureNamedStyles(workbook) {
    for (const name of workbook.RemovedStyleNames ?? []) {
      const existing = allTags(tagContent(this.xml, "cellStyles") ?? "", "cellStyle").find((item) => item.attributes.name === name);
      if (existing) this.xml = removeCountedElement(this.xml, "cellStyles", existing.xml);
      this.namedIndexes.delete(name);
    }
    const styles = new StyleCollection(workbook);
    for (const name of styles.names.filter((candidate) => candidate !== "Normal")) {
      const style = styles.get(name);
      let xfId = this.masterMap.get(styleKey(style));
      if (xfId === undefined) {
        const ids = this.#ids(style);
        const appended = appendCountedElement(this.xml, "cellStyleXfs", xfXml(style, ids));
        this.xml = appended.xml;
        xfId = appended.index;
        this.masterMap.set(styleKey(style), xfId);
      }
      const existing = allTags(tagContent(this.xml, "cellStyles") ?? "", "cellStyle").find((item) => item.attributes.name === name);
      if (existing) {
        const updated = /\bxfId="\d+"/i.test(existing.xml)
          ? existing.xml.replace(/\bxfId="\d+"/i, `xfId="${xfId}"`)
          : existing.xml.replace(/\/>$/, ` xfId="${xfId}"/>`);
        this.xml = this.xml.replace(existing.xml, updated);
      } else {
        this.xml = appendCountedElement(this.xml, "cellStyles", `<cellStyle name="${escapeXml(name)}" xfId="${xfId}"/>`).xml;
      }
      this.namedIndexes.set(name, xfId);
    }
  }

  index(style, cellType, namedStyle) {
    let normalized = structuredClone(style ?? {});
    delete normalized.rangeBorder;
    if (cellType === "date" && !normalized.numberFormat) normalized.numberFormat = "yyyy-mm-dd hh:mm:ss";
    const xfId = this.namedIndexes.get(namedStyle) ?? 0;
    const key = this.#formatKey(normalized, xfId);
    if (this.cellMap.has(key)) return this.cellMap.get(key);
    const ids = this.#ids(normalized);
    const appended = appendCountedElement(this.xml, "cellXfs", xfXml(normalized, ids, { cell: true, xfId }));
    this.xml = appended.xml;
    this.cellMap.set(key, appended.index);
    return appended.index;
  }

  #formatKey(style, xfId) {
    return styleKey([style ?? {}, Number(xfId ?? 0)]);
  }

  #resource(value, collection, map, container, writer) {
    if (!value) return 0;
    const key = styleKey(value);
    if (map.has(key)) return map.get(key);
    const appended = appendCountedElement(this.xml, container, writer(value));
    this.xml = appended.xml;
    collection.push(value);
    map.set(key, appended.index);
    return appended.index;
  }

  #numberFormat(format) {
    if (!format || format === "General") return 0;
    if (this.formatIds.has(format)) return this.formatIds.get(format);
    const ids = [...this.formats.keys()].filter(Number.isFinite);
    const id = Math.max(163, ...ids) + 1;
    const element = `<numFmt numFmtId="${id}" formatCode="${escapeXml(format)}"/>`;
    const existing = allTags(this.xml, "numFmts")[0];
    if (existing) {
      const count = Number(existing.attributes.count ?? allTags(existing.content, "numFmt").length) + 1;
      let updated = existing.xml.replace(/<\/(?:\w+:)?numFmts\s*>$/i, `${element}</numFmts>`);
      updated = /\bcount="\d+"/i.test(updated) ? updated.replace(/\bcount="\d+"/i, `count="${count}"`) : updated.replace(/^(<[^>]+)/, `$1 count="${count}"`);
      this.xml = this.xml.replace(existing.xml, updated);
    } else this.xml = this.xml.replace(/<(?:\w+:)?fonts\b/i, `<numFmts count="1">${element}</numFmts><fonts`);
    this.formats.set(id, format);
    this.formatIds.set(format, id);
    return id;
  }

  #ids(style) {
    return {
      fontId: this.#resource(style.font, this.fonts, this.fontMap, "fonts", fontXml),
      fillId: this.#resource(style.fill, this.fills, this.fillMap, "fills", fillXml),
      borderId: this.#resource(style.border, this.borders, this.borderMap, "borders", borderXml),
      numFmtId: this.#numberFormat(style.numberFormat)
    };
  }
}

function removeElement(xml, name) {
  return xml.replace(new RegExp(`<(?:\\w+:)?${name}\\b[^>]*(?:\\/>|>[\\s\\S]*?<\\/(?:\\w+:)?${name}\\s*>)`, "gi"), "");
}

function patchRowProperties(xml, rows = []) {
  let output = xml;
  rows.forEach((properties, index) => {
    if (!properties) return;
    const rowNumber = index + 1;
    const opening = new RegExp(`<(?:\\w+:)?row\\b(?=[^>]*\\br="${rowNumber}"(?:\\s|/|>))[^>]*`, "i");
    const apply = (tag) => {
      const selfClosing = tag.endsWith("/");
      let updated = tag.replace(/\/$/, "").replace(/\s(?:ht|customHeight|hidden|s|customFormat)="[^"]*"/gi, "");
      if (properties.height) updated += ` ht="${properties.height}" customHeight="1"`;
      if (properties.hidden) updated += ' hidden="1"';
      if (properties.styleIndex) updated += ` s="${properties.styleIndex}" customFormat="1"`;
      return `${updated}${selfClosing ? "/" : ""}`;
    };
    if (opening.test(output)) output = output.replace(opening, apply);
    else {
      const attributes = `${properties.height ? ` ht="${properties.height}" customHeight="1"` : ""}${properties.hidden ? ' hidden="1"' : ""}${properties.styleIndex ? ` s="${properties.styleIndex}" customFormat="1"` : ""}`;
      output = output.replace(/<\/(?:\w+:)?sheetData\s*>/i, `<row r="${rowNumber}"${attributes}/></sheetData>`);
    }
  });
  return output;
}

function patchWorksheetStructures(xml, worksheet) {
  let output = removeElement(xml, "cols");
  const columns = (worksheet["!cols"] ?? []).map((column, index) => {
    if (!column) return "";
    const width = column.width !== undefined ? ` width="${column.width}" customWidth="1"` : "";
    const hidden = column.hidden ? ' hidden="1"' : "";
    const style = column.styleIndex ? ` style="${column.styleIndex}"` : "";
    return `<col min="${index + 1}" max="${index + 1}"${width}${hidden}${style}/>`;
  }).join("");
  if (columns) output = output.replace(/<(?:\w+:)?sheetData\b/i, `<cols>${columns}</cols><sheetData`);
  output = patchRowProperties(output, worksheet["!rows"]);
  for (const name of ["sheetProtection", "autoFilter", "mergeCells"]) output = removeElement(output, name);
  const protection = worksheet["!protection"]
    ? `<sheetProtection${Object.entries(worksheet["!protection"]).map(([key, value]) => ` ${key}="${typeof value === "boolean" ? value ? 1 : 0 : escapeXml(value)}"`).join("")}/>`
    : "";
  const filter = worksheet["!autofilter"] ? `<autoFilter ref="${escapeXml(worksheet["!autofilter"])}"/>` : "";
  const merges = worksheet["!merges"]?.length
    ? `<mergeCells count="${worksheet["!merges"].length}">${worksheet["!merges"].map((range) => `<mergeCell ref="${encodeRange(range)}"/>`).join("")}</mergeCells>`
    : "";
  output = output.replace(/(<\/(?:\w+:)?sheetData\s*>)/i, `$1${protection}${filter}${merges}`);
  return patchDimension(output, worksheet["!ref"]);
}

function patchWorkbookProtection(xml, protection) {
  let output = removeElement(xml, "workbookProtection");
  const element = workbookProtectionXml(protection);
  if (element) output = output.replace(/<(?:\w+:)?bookViews\b/i, `${element}<bookViews`);
  return output;
}

function nextSheetPath(files) {
  let number = 1;
  while (files[`xl/worksheets/sheet${number}.xml`]) number += 1;
  return `xl/worksheets/sheet${number}.xml`;
}

function reconcileWorkbookSheets(workbook, state, files) {
  let workbookXml = Buffer.isBuffer(files[state.workbookPath]) ? files[state.workbookPath].toString("utf8") : String(files[state.workbookPath]);
  const relationshipItems = allTags(Buffer.isBuffer(files[state.workbookRelationshipsPath]) ? files[state.workbookRelationshipsPath].toString("utf8") : String(files[state.workbookRelationshipsPath]), "Relationship")
    .map((item) => ({ id: item.attributes.Id, type: item.attributes.Type, target: item.attributes.Target, ...(item.attributes.TargetMode ? { external: true } : {}) }));
  const relationshipsById = new Map(relationshipItems.map((item) => [item.id, item]));
  const existing = allTags(tagContent(workbookXml, "sheets") ?? "", "sheet").map((item) => {
    const relationship = relationshipsById.get(item.attributes["r:id"]);
    return { ...item, relationship, path: relationship ? resolvePart(state.workbookPath, relationship.target) : undefined };
  });
  const desiredNames = new Set(workbook.SheetNames);
  const retainedPaths = new Set([...state.sheetPaths].filter(([name]) => desiredNames.has(name)).map(([, path]) => path));
  for (const entry of existing) {
    if (!entry.path || retainedPaths.has(entry.path)) continue;
    delete files[entry.path];
    delete files[sheetRelationshipsPath(entry.path)];
    const relationshipIndex = relationshipItems.findIndex((item) => item.id === entry.relationship?.id);
    if (relationshipIndex >= 0) relationshipItems.splice(relationshipIndex, 1);
    let contentTypes = Buffer.isBuffer(files["[Content_Types].xml"]) ? files["[Content_Types].xml"].toString("utf8") : String(files["[Content_Types].xml"]);
    const escapedPath = escapeRegExp(`/${entry.path}`);
    contentTypes = contentTypes.replace(new RegExp(`<(?:\\w+:)?Override\\b(?=[^>]*\\bPartName="${escapedPath}")[^>]*/>`, "i"), "");
    files["[Content_Types].xml"] = contentTypes;
  }
  for (const name of [...state.sheetPaths.keys()]) if (!desiredNames.has(name)) state.sheetPaths.delete(name);

  const usedRelationshipIds = new Set(relationshipItems.map((item) => item.id));
  const nextRelationshipId = () => {
    let number = 1;
    while (usedRelationshipIds.has(`rId${number}`)) number += 1;
    usedRelationshipIds.add(`rId${number}`);
    return `rId${number}`;
  };
  let nextSheetId = Math.max(0, ...existing.map((entry) => Number(entry.attributes.sheetId)).filter(Number.isFinite)) + 1;
  const tags = workbook.SheetNames.map((name) => {
    let path = state.sheetPaths.get(name);
    let entry = path ? existing.find((item) => item.path === path) : undefined;
    if (!path) {
      path = nextSheetPath(files);
      state.sheetPaths.set(name, path);
      const id = nextRelationshipId();
      relationshipItems.push({ id, type: `${REL_NS}/worksheet`, target: posix.relative(posix.dirname(state.workbookPath), path) });
      files[path] = `${XML_HEADER}<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData></sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`;
      let contentTypes = Buffer.isBuffer(files["[Content_Types].xml"]) ? files["[Content_Types].xml"].toString("utf8") : String(files["[Content_Types].xml"]);
      contentTypes = contentTypes.replace(/<\/(?:\w+:)?Types\s*>/i, `<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`);
      files["[Content_Types].xml"] = contentTypes;
      entry = { attributes: { sheetId: nextSheetId++, "r:id": id }, xml: "" };
    }
    const id = entry.attributes["r:id"];
    return `<sheet name="${escapeXml(name)}" sheetId="${entry.attributes.sheetId}" r:id="${id}"${entry.attributes.state ? ` state="${escapeXml(entry.attributes.state)}"` : ""}/>`;
  });
  const sheets = allTags(workbookXml, "sheets")[0];
  workbookXml = workbookXml.replace(sheets.xml, `<sheets>${tags.join("")}</sheets>`);
  files[state.workbookRelationshipsPath] = relationshipsXml(relationshipItems);
  files[state.workbookPath] = workbookXml;
}

function patchHyperlink(files, sheetPath, sheetXml, address, hyperlink) {
  const links = allTags(sheetXml, "hyperlinks")[0];
  const existing = allTags(links?.content ?? "", "hyperlink").find((item) => item.attributes.ref === address);
  let relationshipId = existing?.attributes["r:id"];
  let relsPath = sheetRelationshipsPath(sheetPath);
  let relsXml = files[relsPath] ? (Buffer.isBuffer(files[relsPath]) ? files[relsPath].toString("utf8") : String(files[relsPath])) : "";
  if (relationshipId && relsXml) {
    const escapedId = escapeRegExp(relationshipId);
    relsXml = relsXml.replace(new RegExp(`<(?:\\w+:)?Relationship\\b(?=[^>]*\\bId="${escapedId}")[^>]*/>`, "i"), "");
    relationshipId = undefined;
  }
  let item = "";
  if (hyperlink) {
    if (hyperlink.target.startsWith("#")) {
      item = `<hyperlink ref="${address}" location="${escapeXml(hyperlink.target.slice(1))}"${hyperlink.tooltip ? ` tooltip="${escapeXml(hyperlink.tooltip)}"` : ""}/>`;
    } else {
      const ids = allTags(relsXml, "Relationship").map((relation) => relation.attributes.Id);
      let number = 1;
      while (ids.includes(`rId${number}`)) number += 1;
      relationshipId = `rId${number}`;
      item = `<hyperlink ref="${address}" r:id="${relationshipId}"${hyperlink.tooltip ? ` tooltip="${escapeXml(hyperlink.tooltip)}"` : ""}/>`;
      const relationship = `<Relationship Id="${relationshipId}" Type="${REL_NS}/hyperlink" Target="${escapeXml(hyperlink.target)}" TargetMode="External"/>`;
      relsXml = relsXml
        ? relsXml.replace(/<\/(?:\w+:)?Relationships\s*>/i, `${relationship}</Relationships>`)
        : relationshipsXml([{ id: relationshipId, type: `${REL_NS}/hyperlink`, target: hyperlink.target, external: true }]);
    }
  }
  if (relsXml) files[relsPath] = relsXml;
  if (links) {
    const kept = allTags(links.content, "hyperlink").filter((entry) => entry.attributes.ref !== address).map((entry) => entry.xml);
    if (item) kept.push(item);
    sheetXml = sheetXml.replace(links.xml, kept.length ? `<hyperlinks>${kept.join("")}</hyperlinks>` : "");
  } else if (item) sheetXml = sheetXml.replace(/<(?:\w+:)?pageMargins\b/i, `<hyperlinks>${item}</hyperlinks><pageMargins`);
  return sheetXml;
}

function writePreservedXlsx(workbook, state) {
  const files = Object.fromEntries(state.files);
  if (state.workbookDirty) reconcileWorkbookSheets(workbook, state, files);
  let stylesXml = state.stylesPath && files[state.stylesPath]
    ? (Buffer.isBuffer(files[state.stylesPath]) ? files[state.stylesPath].toString("utf8") : String(files[state.stylesPath]))
    : undefined;
  const styleRegistry = stylesXml ? new PreservedStyleRegistry(stylesXml) : undefined;
  if (styleRegistry) {
    styleRegistry.ensureNamedStyles(workbook);
    stylesXml = styleRegistry.xml;
  }
  for (const [sheetName, cells] of state.dirtyCells) {
    const path = state.sheetPaths.get(sheetName);
    if (!path || !files[path]) throw new Error(`Cannot locate the original XML part for worksheet ${sheetName}.`);
    let xml = Buffer.isBuffer(files[path]) ? files[path].toString("utf8") : String(files[path]);
    for (const [address, kinds] of cells) {
      const cell = workbook.Sheets[sheetName]?.[address];
      if (kinds.has("style")) {
        if (!styleRegistry) throw new Error("Cannot edit template styles because the workbook has no styles part.");
        const styleIndex = styleRegistry.index(cell?.style, cell?.type, cell?.namedStyle);
        stylesXml = styleRegistry.xml;
        if (cell) {
          cell.styleIndex = styleIndex;
          cell.styleDirty = false;
        }
      }
      if (kinds.has("hyperlink")) xml = patchHyperlink(files, path, xml, address, cell?.hyperlink);
      xml = patchCell(xml, address, cell);
    }
    files[path] = sortSheetData(patchDimension(xml, workbook.Sheets[sheetName]?.["!ref"]));
  }
  for (const sheetName of state.dirtySheetStructures) {
    const path = state.sheetPaths.get(sheetName);
    if (!path || !files[path]) throw new Error(`Cannot locate the original XML part for worksheet ${sheetName}.`);
    const worksheet = workbook.Sheets[sheetName];
    for (const column of worksheet?.["!cols"] ?? []) {
      if (!column?.style) continue;
      if (!styleRegistry) throw new Error("Cannot edit template column styles because the workbook has no styles part.");
      column.styleIndex = styleRegistry.index(column.style, undefined, column.namedStyle);
      stylesXml = styleRegistry.xml;
    }
    for (const row of worksheet?.["!rows"] ?? []) {
      if (!row?.style) continue;
      if (!styleRegistry) throw new Error("Cannot edit template row styles because the workbook has no styles part.");
      row.styleIndex = styleRegistry.index(row.style, undefined, row.namedStyle);
      stylesXml = styleRegistry.xml;
    }
    const xml = Buffer.isBuffer(files[path]) ? files[path].toString("utf8") : String(files[path]);
    files[path] = sortSheetData(patchWorksheetStructures(xml, worksheet));
  }
  if (state.workbookDirty) {
    const xml = Buffer.isBuffer(files[state.workbookPath]) ? files[state.workbookPath].toString("utf8") : String(files[state.workbookPath]);
    files[state.workbookPath] = patchWorkbookProtection(xml, workbook.WorkbookProtection);
  }
  if (stylesXml && state.stylesPath) files[state.stylesPath] = stylesXml;
  if (state.propertiesDirty && state.corePath && files[state.corePath]) {
    const core = Buffer.isBuffer(files[state.corePath]) ? files[state.corePath].toString("utf8") : String(files[state.corePath]);
    files[state.corePath] = patchCoreProperties(core, workbook.Props ?? {});
  }
  const output = createZip(files);
  state.files = new Map(Object.entries(files).map(([path, value]) => [path, Buffer.isBuffer(value) ? value : Buffer.from(value)]));
  state.dirtyCells.clear();
  state.dirtySheetStructures.clear();
  state.workbookDirty = false;
  state.propertiesDirty = false;
  workbook.RemovedStyleNames = [];
  return output;
}

export function writeXlsx(workbook) {
  const state = getPackageState(workbook);
  return state ? writePreservedXlsx(workbook, state) : writeNewXlsx(workbook);
}

export function ensurePackageState(workbook) {
  const existing = getPackageState(workbook);
  if (existing) return existing;
  const parsed = readXlsx(writeNewXlsx(workbook));
  for (const sheetName of workbook.SheetNames) {
    const sourceSheet = workbook.Sheets[sheetName];
    const parsedSheet = parsed.Sheets[sheetName];
    for (const address of Object.keys(sourceSheet).filter((key) => /^[A-Z]+[1-9]\d*$/.test(key))) {
      sourceSheet[address].styleIndex = parsedSheet[address]?.styleIndex ?? 0;
    }
  }
  attachPackageState(workbook, getPackageState(parsed));
  return getPackageState(workbook);
}

function readPart(files, path, required = true) {
  const value = files.get(path);
  if (!value && required) throw new Error(`Invalid XLSX file: missing ${path}.`);
  return value?.toString("utf8") ?? "";
}

function relationshipMap(xml) {
  return new Map(allTags(xml, "Relationship").map(({ attributes }) => [attributes.Id, attributes]));
}

function resolvePart(base, target) {
  if (target.startsWith("/")) return target.slice(1);
  return posix.normalize(posix.join(posix.dirname(base), target));
}

function parseFont(content) {
  const flag = (name) => {
    const tag = allTags(content, name)[0];
    return tag ? !/^(?:0|false)$/i.test(tag.attributes.val ?? "1") : undefined;
  };
  const value = (name, transform = (item) => item) => {
    const tag = allTags(content, name)[0];
    return tag?.attributes.val === undefined ? undefined : transform(tag.attributes.val);
  };
  return {
    ...(flag("b") !== undefined ? { bold: flag("b") } : {}),
    ...(flag("i") !== undefined ? { italic: flag("i") } : {}),
    ...(flag("strike") !== undefined ? { strike: flag("strike") } : {}),
    ...(flag("outline") !== undefined ? { outline: flag("outline") } : {}),
    ...(flag("shadow") !== undefined ? { shadow: flag("shadow") } : {}),
    ...(flag("condense") !== undefined ? { condense: flag("condense") } : {}),
    ...(flag("extend") !== undefined ? { extend: flag("extend") } : {}),
    ...(allTags(content, "u")[0] ? { underline: allTags(content, "u")[0].attributes.val ?? "single" } : {}),
    ...(value("name") !== undefined ? { name: value("name") } : {}),
    ...(value("sz", Number) !== undefined ? { size: value("sz", Number) } : {}),
    ...(value("family", Number) !== undefined ? { family: value("family", Number) } : {}),
    ...(value("charset", Number) !== undefined ? { charset: value("charset", Number) } : {}),
    ...(value("scheme") !== undefined ? { scheme: value("scheme") } : {}),
    ...(value("vertAlign") !== undefined ? { verticalAlign: value("vertAlign") } : {}),
    ...(parseColor(content) ? { color: parseColor(content) } : {})
  };
}

function parseFill(content) {
  const gradient = allTags(content, "gradientFill")[0];
  if (gradient) {
    const gradientType = gradient.attributes.type === "path" ? "path" : "linear";
    return {
      type: "gradient",
      gradientType,
      ...(gradientType === "linear" ? { degree: Number(gradient.attributes.degree ?? 0) } : {}),
      ...Object.fromEntries(["left", "right", "top", "bottom"].filter((key) => gradient.attributes[key] !== undefined).map((key) => [key, Number(gradient.attributes[key])])),
      stops: allTags(gradient.content, "stop").map((stop) => ({ position: Number(stop.attributes.position), color: parseColor(stop.content) })).filter((stop) => stop.color)
    };
  }
  const pattern = allTags(content, "patternFill")[0];
  if (!pattern) return {};
  const patternType = pattern.attributes.patternType ?? "none";
  const foreground = parseColor(pattern.content, "fgColor");
  const background = parseColor(pattern.content, "bgColor");
  if (patternType === "none" && !foreground && !background) return {};
  return { type: "pattern", patternType, ...(foreground ? { foreground } : {}), ...(background ? { background } : {}) };
}

function parseBorder(content, attributes = {}) {
  const border = Object.fromEntries(["diagonalUp", "diagonalDown", "outline"].filter((key) => attributes[key] !== undefined).map((key) => [key, /^(?:1|true)$/i.test(attributes[key])]));
  for (const side of ["left", "right", "top", "bottom", "diagonal", "vertical", "horizontal"]) {
    const tag = allTags(content, side)[0];
    if (!tag) continue;
    const color = parseColor(tag.content);
    if (tag.attributes.style || color) border[side] = { ...(tag.attributes.style ? { style: tag.attributes.style } : {}), ...(color ? { color } : {}) };
  }
  return border;
}

function parseScalar(value) {
  if (/^(?:0|1)$/.test(value)) return value === "1";
  if (/^(?:true|false)$/i.test(value)) return value.toLowerCase() === "true";
  return Number.isNaN(Number(value)) ? value : Number(value);
}

function parseStyleRecords(xml, containerName) {
  if (!xml) return [{ style: {}, xfId: 0 }];
  const formats = new Map(builtinFormats);
  for (const { attributes } of allTags(tagContent(xml, "numFmts") ?? "", "numFmt")) formats.set(Number(attributes.numFmtId), attributes.formatCode);
  const fonts = allTags(tagContent(xml, "fonts") ?? "", "font").map(({ content }) => parseFont(content));
  const fills = allTags(tagContent(xml, "fills") ?? "", "fill").map(({ content }) => parseFill(content));
  const borders = allTags(tagContent(xml, "borders") ?? "", "border").map(({ content, attributes }) => parseBorder(content, attributes));
  return allTags(tagContent(xml, containerName) ?? "", "xf").map(({ attributes, content }) => {
    const alignmentMatch = /<(?:\w+:)?alignment\b([^>]*)\/>/i.exec(content);
    const protectionMatch = /<(?:\w+:)?protection\b([^>]*)\/>/i.exec(content);
    const format = formats.get(Number(attributes.numFmtId));
    const font = fonts[Number(attributes.fontId)] ?? {};
    const fill = fills[Number(attributes.fillId)] ?? {};
    const border = borders[Number(attributes.borderId)] ?? {};
    const includeFont = Number(attributes.fontId ?? 0) !== 0 || /^(?:1|true)$/i.test(attributes.applyFont ?? "0");
    const includeFill = Number(attributes.fillId ?? 0) !== 0 || /^(?:1|true)$/i.test(attributes.applyFill ?? "0");
    const includeBorder = Number(attributes.borderId ?? 0) !== 0 || /^(?:1|true)$/i.test(attributes.applyBorder ?? "0");
    return { xfId: Number(attributes.xfId ?? 0), style: {
      ...(includeFont && Object.keys(font).length ? { font } : {}),
      ...(includeFill && Object.keys(fill).length ? { fill } : {}),
      ...(includeBorder && Object.keys(border).length ? { border } : {}),
      ...(format && format !== "General" ? { numberFormat: format } : {}),
      ...(alignmentMatch ? { alignment: Object.fromEntries(Object.entries(parseAttributes(alignmentMatch[1])).map(([key, value]) => [key, parseScalar(value)])) } : {}),
      ...(protectionMatch ? { protection: Object.fromEntries(Object.entries(parseAttributes(protectionMatch[1])).map(([key, value]) => [key, /^(?:1|true)$/i.test(value)])) } : {})
    } };
  });
}

function parseStyles(xml) {
  if (!xml) return { styles: [{}], namedStyles: [], definedStyles: {} };
  const cells = parseStyleRecords(xml, "cellXfs");
  const masters = parseStyleRecords(xml, "cellStyleXfs");
  const namesByXfId = new Map(allTags(tagContent(xml, "cellStyles") ?? "", "cellStyle").map(({ attributes }) => [Number(attributes.xfId), attributes.name]));
  const definedStyles = {};
  for (const [xfId, name] of namesByXfId) {
    if (name && name !== "Normal" && masters[xfId]) definedStyles[name] = { style: masters[xfId].style, extends: [] };
  }
  return {
    styles: cells.map((record) => record.style),
    namedStyles: cells.map((record) => namesByXfId.get(record.xfId)),
    definedStyles
  };
}

function isDateFormat(format = "") {
  const cleaned = format.replace(/"[^"]*"|\[[^\]]*\]|\\./g, "").toLowerCase();
  return /[ymdhis]/.test(cleaned) && !/^general$/.test(cleaned);
}

function sharedStrings(xml) {
  return allTags(xml, "si").map(({ content }) => allTags(content, "t").map((tag) => decodeXml(tag.content)).join(""));
}

function sheetRelationshipsPath(sheetPath) {
  return posix.join(posix.dirname(sheetPath), "_rels", `${posix.basename(sheetPath)}.rels`);
}

function readSheet(xml, styleInfo, strings, relationships) {
  const worksheet = {};
  const dimension = allTags(xml, "dimension")[0]?.attributes.ref;
  if (dimension) worksheet["!ref"] = dimension;
  for (const { attributes, content } of allTags(xml, "c")) {
    const address = attributes.r;
    if (!address) continue;
    const styleIndex = Number(attributes.s ?? 0);
    const style = styleInfo.styles[styleIndex] ?? {};
    const namedStyle = styleInfo.namedStyles[styleIndex];
    const formulaMatch = /<(?:\w+:)?f\b[^>]*>([\s\S]*?)<\/(?:\w+:)?f>/i.exec(content);
    const valueMatch = /<(?:\w+:)?v\b[^>]*>([\s\S]*?)<\/(?:\w+:)?v>/i.exec(content);
    const inline = tagContent(content, "is");
    const rawValue = valueMatch ? decodeXml(valueMatch[1]) : undefined;
    let type = "blank";
    let value;
    if (attributes.t === "inlineStr") { type = "string"; value = allTags(inline ?? "", "t").map((tag) => decodeXml(tag.content)).join(""); }
    else if (attributes.t === "s") { type = "string"; value = strings[Number(rawValue)] ?? ""; }
    else if (attributes.t === "str") { type = "string"; value = rawValue ?? ""; }
    else if (attributes.t === "b") { type = "boolean"; value = rawValue === "1"; }
    else if (attributes.t === "d") { type = "date"; value = new Date(rawValue); }
    else if (rawValue !== undefined && rawValue !== "") {
      const numeric = Number(rawValue);
      if (isDateFormat(style.numberFormat)) { type = "date"; value = dateFromSerial(numeric); }
      else { type = "number"; value = numeric; }
    }
    worksheet[address] = {
      type,
      value,
      styleIndex,
      ...(formulaMatch ? { formula: decodeXml(formulaMatch[1]) } : {}),
      ...(styleIndex ? { style } : {}),
      ...(namedStyle && namedStyle !== "Normal" ? { namedStyle } : {})
    };
    if (!worksheet["!ref"]) expandWorksheetRef(worksheet, decodeAddress(address));
  }
  const merges = allTags(xml, "mergeCell").map(({ attributes }) => decodeRange(attributes.ref)).filter(Boolean);
  if (merges.length) worksheet["!merges"] = merges;
  const columns = [];
  for (const { attributes } of allTags(xml, "col")) {
    for (let index = Number(attributes.min) - 1; index < Number(attributes.max); index += 1) {
      const styleIndex = Number(attributes.style ?? 0);
      const style = styleInfo.styles[styleIndex] ?? {};
      const namedStyle = styleInfo.namedStyles[styleIndex];
      columns[index] = {
        ...(attributes.width ? { width: Number(attributes.width) } : {}),
        ...(attributes.hidden === "1" ? { hidden: true } : {}),
        ...(styleIndex ? { styleIndex, style } : {}),
        ...(namedStyle && namedStyle !== "Normal" ? { namedStyle } : {})
      };
    }
  }
  if (columns.length) worksheet["!cols"] = columns;
  const rows = [];
  for (const { attributes } of allTags(xml, "row")) {
    const styleIndex = Number(attributes.s ?? 0);
    const style = styleInfo.styles[styleIndex] ?? {};
    const namedStyle = styleInfo.namedStyles[styleIndex];
    if (attributes.ht || attributes.hidden === "1" || styleIndex) {
      rows[Number(attributes.r) - 1] = {
        ...(attributes.ht ? { height: Number(attributes.ht) } : {}),
        ...(attributes.hidden === "1" ? { hidden: true } : {}),
        ...(styleIndex ? { styleIndex, style } : {}),
        ...(namedStyle && namedStyle !== "Normal" ? { namedStyle } : {})
      };
    }
  }
  if (rows.length) worksheet["!rows"] = rows;
  const filter = allTags(xml, "autoFilter")[0]?.attributes.ref;
  if (filter) worksheet["!autofilter"] = filter;
  const protection = allTags(xml, "sheetProtection")[0]?.attributes;
  if (protection) worksheet["!protection"] = Object.fromEntries(Object.entries(protection).map(([key, value]) => [key, /^(?:0|1)$/.test(value) ? value === "1" : value]));
  for (const { attributes } of allTags(xml, "hyperlink")) {
    const cell = worksheet[attributes.ref];
    if (!cell) continue;
    const relation = relationships.get(attributes["r:id"]);
    const target = attributes.location ? `#${attributes.location}` : relation?.Target;
    if (target) cell.hyperlink = { target, ...(attributes.tooltip ? { tooltip: attributes.tooltip } : {}) };
  }
  return worksheet;
}

export function readXlsx(input) {
  const files = extractZip(input);
  const packageRelationships = relationshipMap(readPart(files, "_rels/.rels"));
  const officeDocument = [...packageRelationships.values()].find((relationship) => relationship.Type?.endsWith("/officeDocument"));
  const workbookPath = officeDocument ? officeDocument.Target.replace(/^\//, "") : "xl/workbook.xml";
  const workbookXml = readPart(files, workbookPath);
  const workbookRelationshipsPath = posix.join(posix.dirname(workbookPath), "_rels", `${posix.basename(workbookPath)}.rels`);
  const workbookRelationships = relationshipMap(readPart(files, workbookRelationshipsPath));
  const stylesRelationship = [...workbookRelationships.values()].find((relationship) => relationship.Type?.endsWith("/styles"));
  const stringsRelationship = [...workbookRelationships.values()].find((relationship) => relationship.Type?.endsWith("/sharedStrings"));
  const stylesXml = stylesRelationship ? readPart(files, resolvePart(workbookPath, stylesRelationship.Target), false) : "";
  const styleInfo = parseStyles(stylesXml);
  const strings = stringsRelationship ? sharedStrings(readPart(files, resolvePart(workbookPath, stringsRelationship.Target), false)) : [];
  const workbook = { SheetNames: [], Sheets: {}, Props: {}, DefinedStyles: styleInfo.definedStyles };
  const workbookProtection = allTags(workbookXml, "workbookProtection")[0]?.attributes;
  if (workbookProtection) workbook.WorkbookProtection = Object.fromEntries(Object.entries(workbookProtection).map(([key, value]) => [key, /^(?:0|1)$/.test(value) ? value === "1" : value]));
  const sheetPaths = new Map();
  for (const { attributes } of allTags(workbookXml, "sheet")) {
    const relationship = workbookRelationships.get(attributes["r:id"]);
    if (!relationship) continue;
    const sheetPath = resolvePart(workbookPath, relationship.Target);
    const sheetRels = relationshipMap(readPart(files, sheetRelationshipsPath(sheetPath), false));
    workbook.SheetNames.push(attributes.name);
    sheetPaths.set(attributes.name, sheetPath);
    workbook.Sheets[attributes.name] = readSheet(readPart(files, sheetPath), styleInfo, strings, sheetRels);
  }
  const coreRelationship = [...packageRelationships.values()].find((relationship) => relationship.Type?.endsWith("/core-properties"));
  const corePath = coreRelationship ? coreRelationship.Target.replace(/^\//, "") : "docProps/core.xml";
  const core = readPart(files, corePath, false);
  const property = (name) => decodeXml(tagContent(core, name) ?? "");
  const createdAt = property("created");
  workbook.Props = {
    ...(property("title") ? { title: property("title") } : {}),
    ...(property("creator") ? { author: property("creator") } : {}),
    ...(property("lastModifiedBy") ? { lastAuthor: property("lastModifiedBy") } : {}),
    ...(createdAt ? { createdAt: new Date(createdAt) } : {})
  };
  if (!workbook.SheetNames.length) throw new Error("Invalid XLSX file: workbook has no worksheets.");
  return attachPackageState(workbook, new PackageState({
    files,
    workbookPath,
    workbookRelationshipsPath,
    stylesPath: stylesRelationship ? resolvePart(workbookPath, stylesRelationship.Target) : undefined,
    sharedStringsPath: stringsRelationship ? resolvePart(workbookPath, stringsRelationship.Target) : undefined,
    sheetPaths,
    corePath
  }));
}
