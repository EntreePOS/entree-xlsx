import { decodeAddress, decodeRange, encodeAddress } from "./address.js";
import { ensurePackageState } from "./ooxml.js";
import {
  addContentTypeOverride,
  addRelationship,
  getTextFile,
  nextPartPath,
  readRelationships,
  relationshipTarget,
  removeContentTypeOverride,
  removeRelationship,
  setTextFile
} from "./package-xml.js";
import { allTags, decodeXml, escapeXml } from "./xml.js";

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const DRAWING_MAIN_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const DRAWING_REL = `${REL_NS}/drawing`;
const CHART_REL = `${REL_NS}/chart`;

function absoluteCell(address) {
  const point = decodeAddress(address);
  const encoded = encodeAddress(point);
  return `$${encoded.replace(/(\d+)$/, "$$$1")}`;
}

function formula(sheetName, range) {
  if (range.includes("!")) return range;
  const [start, end] = range.split(":");
  const quoted = `'${sheetName.replace(/'/g, "''")}'`;
  return `${quoted}!${absoluteCell(start)}${end ? `:${absoluteCell(end)}` : ""}`;
}

function marker(address) {
  const point = decodeAddress(address);
  return `<xdr:col>${point.c}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${point.r}</xdr:row><xdr:rowOff>0</xdr:rowOff>`;
}

function titleXml(title) {
  if (!title) return "";
  return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${escapeXml(title)}</a:t></a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val="0"/></c:title>`;
}

function seriesNameXml(series) {
  if (series.nameCell) return `<c:tx><c:strRef><c:f>${escapeXml(series.nameCell)}</c:f></c:strRef></c:tx>`;
  return series.name ? `<c:tx><c:v>${escapeXml(series.name)}</c:v></c:tx>` : "";
}

function categoryXml(reference) {
  return reference ? `<c:cat><c:strRef><c:f>${escapeXml(reference)}</c:f></c:strRef></c:cat>` : "";
}

function standardSeriesXml(series, index) {
  return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>${seriesNameXml(series)}${categoryXml(series.categories)}<c:val><c:numRef><c:f>${escapeXml(series.values)}</c:f></c:numRef></c:val></c:ser>`;
}

function scatterSeriesXml(series, index) {
  return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>${seriesNameXml(series)}<c:xVal><c:numRef><c:f>${escapeXml(series.xValues ?? series.categories)}</c:f></c:numRef></c:xVal><c:yVal><c:numRef><c:f>${escapeXml(series.values)}</c:f></c:numRef></c:yVal></c:ser>`;
}

function chartBody(type, series) {
  const content = series.map((item, index) => type === "scatter" ? scatterSeriesXml(item, index) : standardSeriesXml(item, index)).join("");
  if (type === "pie") return `<c:pieChart><c:varyColors val="1"/>${content}<c:firstSliceAng val="0"/></c:pieChart>`;
  if (type === "line") return `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${content}<c:marker val="1"/><c:smooth val="0"/><c:axId val="48650112"/><c:axId val="48672768"/></c:lineChart>`;
  if (type === "scatter") return `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${content}<c:axId val="48650112"/><c:axId val="48672768"/></c:scatterChart>`;
  const direction = type === "bar" ? "bar" : "col";
  return `<c:barChart><c:barDir val="${direction}"/><c:grouping val="clustered"/><c:varyColors val="0"/>${content}<c:gapWidth val="150"/><c:axId val="48650112"/><c:axId val="48672768"/></c:barChart>`;
}

function axesXml(type) {
  if (type === "pie") return "";
  const horizontal = type === "scatter" ? "valAx" : "catAx";
  const first = horizontal === "catAx"
    ? '<c:catAx><c:axId val="48650112"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:tickLblPos val="nextTo"/><c:crossAx val="48672768"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>'
    : '<c:valAx><c:axId val="48650112"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="48672768"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>';
  return `${first}<c:valAx><c:axId val="48672768"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorGridlines/><c:tickLblPos val="nextTo"/><c:crossAx val="48650112"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
}

function chartXml(config) {
  const legend = config.legend === false ? "" : `<c:legend><c:legendPos val="${escapeXml(config.legendPosition ?? "r")}"/><c:layout/><c:overlay val="0"/></c:legend>`;
  return `${XML_HEADER}<c:chartSpace xmlns:c="${CHART_NS}" xmlns:a="${DRAWING_MAIN_NS}" xmlns:r="${REL_NS}"><c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/><c:chart>${titleXml(config.title)}<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>${chartBody(config.type, config.series)}${axesXml(config.type)}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/></c:chart><c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>`;
}

function normalizeSeries(sheetName, series) {
  return series.map((item, index) => {
    if (!item?.values) throw new TypeError(`Chart series ${index + 1} requires a values range.`);
    return {
      ...(item.name ? { name: String(item.name) } : {}),
      ...(item.nameCell ? { nameCell: formula(sheetName, item.nameCell) } : {}),
      ...(item.categories ? { categories: formula(sheetName, item.categories) } : {}),
      ...(item.xValues ? { xValues: formula(sheetName, item.xValues) } : {}),
      values: formula(sheetName, item.values)
    };
  });
}

function seriesFromRange(sheetName, range) {
  const decoded = decodeRange(range);
  if (decoded.e.r <= decoded.s.r || decoded.e.c <= decoded.s.c) throw new RangeError("Chart range needs a header row and at least two columns.");
  const categories = `${encodeAddress({ r: decoded.s.r + 1, c: decoded.s.c })}:${encodeAddress({ r: decoded.e.r, c: decoded.s.c })}`;
  const result = [];
  for (let column = decoded.s.c + 1; column <= decoded.e.c; column += 1) {
    result.push({
      nameCell: encodeAddress({ r: decoded.s.r, c: column }),
      categories,
      values: `${encodeAddress({ r: decoded.s.r + 1, c: column })}:${encodeAddress({ r: decoded.e.r, c: column })}`
    });
  }
  return normalizeSeries(sheetName, result);
}

function normalizeConfig(sheetName, config) {
  const type = config.type ?? "column";
  if (!["column", "bar", "line", "pie", "scatter"].includes(type)) throw new RangeError(`Unsupported chart type: ${type}.`);
  const series = config.series ? normalizeSeries(sheetName, config.series) : config.range ? seriesFromRange(sheetName, config.range) : undefined;
  if (!series?.length) throw new TypeError("A chart requires either range or series.");
  return {
    type,
    name: config.name,
    title: config.title,
    series,
    legend: config.legend,
    legendPosition: config.legendPosition,
    position: { from: config.position?.from ?? "E2", to: config.position?.to ?? "M18" }
  };
}

function drawingAnchor(id, name, relationshipId, position) {
  return `<xdr:twoCellAnchor editAs="oneCell"><xdr:from>${marker(position.from)}</xdr:from><xdr:to>${marker(position.to)}</xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${escapeXml(name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="${CHART_NS}"><c:chart xmlns:c="${CHART_NS}" xmlns:r="${REL_NS}" r:id="${relationshipId}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
}

function worksheetDrawing(state, sheetPath) {
  const sheetXml = getTextFile(state, sheetPath);
  const drawing = allTags(sheetXml, "drawing")[0];
  if (!drawing?.attributes["r:id"]) return undefined;
  const drawingPath = relationshipTarget(state, sheetPath, drawing.attributes["r:id"]);
  return drawingPath ? { drawingPath, sheetRelationshipId: drawing.attributes["r:id"] } : undefined;
}

function ensureDrawing(state, sheetPath) {
  const existing = worksheetDrawing(state, sheetPath);
  if (existing) return existing;
  const drawingPath = nextPartPath(state, "xl/drawings", "drawing");
  setTextFile(state, drawingPath, `${XML_HEADER}<xdr:wsDr xmlns:xdr="${DRAWING_NS}" xmlns:a="${DRAWING_MAIN_NS}"/>`);
  addContentTypeOverride(state, drawingPath, "application/vnd.openxmlformats-officedocument.drawing+xml");
  const relationshipId = addRelationship(state, sheetPath, DRAWING_REL, drawingPath);
  let sheetXml = getTextFile(state, sheetPath);
  sheetXml = sheetXml.replace(/<\/(?:\w+:)?worksheet\s*>/i, `<drawing r:id="${relationshipId}"/></worksheet>`);
  setTextFile(state, sheetPath, sheetXml);
  return { drawingPath, sheetRelationshipId: relationshipId };
}

function chartType(xml) {
  if (/<(?:\w+:)?pieChart\b/i.test(xml)) return "pie";
  if (/<(?:\w+:)?lineChart\b/i.test(xml)) return "line";
  if (/<(?:\w+:)?scatterChart\b/i.test(xml)) return "scatter";
  if (/<(?:\w+:)?barChart\b/i.test(xml)) return /<(?:\w+:)?barDir\b[^>]*\bval="bar"/i.test(xml) ? "bar" : "column";
  return "unknown";
}

function firstText(xml, tag) {
  return allTags(xml, tag)[0]?.content ? decodeXml(allTags(xml, tag)[0].content) : undefined;
}

function chartEntries(state, sheetName) {
  const sheetPath = state.sheetPaths.get(sheetName);
  if (!sheetPath) return [];
  const drawing = worksheetDrawing(state, sheetPath);
  if (!drawing) return [];
  const drawingXml = getTextFile(state, drawing.drawingPath);
  const relationships = readRelationships(state, drawing.drawingPath).items;
  return allTags(drawingXml, "twoCellAnchor").flatMap(({ content, xml: anchorXml }) => {
    const chartTag = allTags(content, "chart")[0];
    const relationshipId = chartTag?.attributes["r:id"];
    const relationship = relationships.find((item) => item.id === relationshipId && item.type === CHART_REL);
    if (!relationship) return [];
    const partPath = relationshipTarget(state, drawing.drawingPath, relationshipId);
    const xml = getTextFile(state, partPath);
    const properties = allTags(content, "cNvPr")[0]?.attributes ?? {};
    const markers = allTags(content, "from")[0]?.content;
    const endMarkers = allTags(content, "to")[0]?.content;
    const markerAddress = (markerXml) => markerXml ? encodeAddress({ c: Number(firstText(markerXml, "col")), r: Number(firstText(markerXml, "row")) }) : undefined;
    return [{
      id: partPath,
      name: properties.name,
      type: chartType(xml),
      title: allTags(allTags(xml, "title")[0]?.content ?? "", "t").map((item) => decodeXml(item.content)).join("") || undefined,
      position: { from: markerAddress(markers), to: markerAddress(endMarkers) },
      partPath,
      drawingPath: drawing.drawingPath,
      relationshipId,
      anchorXml,
      xml
    }];
  });
}

function findChart(state, workbook, reference) {
  for (const sheetName of workbook.SheetNames) {
    const entry = chartEntries(state, sheetName).find((item) => item.id === reference || item.name === reference);
    if (entry) return { ...entry, sheet: sheetName };
  }
  return undefined;
}

export class ChartCollection {
  constructor(workbook) {
    this.workbook = workbook;
  }

  list(sheet) {
    const state = ensurePackageState(this.workbook);
    const names = sheet === undefined ? this.workbook.SheetNames : [typeof sheet === "number" ? this.workbook.SheetNames[sheet] : sheet];
    return names.flatMap((sheetName) => chartEntries(state, sheetName).map(({ anchorXml, xml, ...entry }) => ({ ...entry, sheet: sheetName })));
  }

  add(config) {
    if (!config || typeof config !== "object" || Array.isArray(config)) throw new TypeError("Chart options must be an object.");
    const { sheet, ...chartConfig } = config;
    if (sheet === undefined) throw new TypeError("A chart requires a sheet.");
    const sheetName = typeof sheet === "number" ? this.workbook.SheetNames[sheet] : sheet;
    if (!this.workbook.Sheets[sheetName]) throw new RangeError(`Worksheet ${String(sheet)} does not exist.`);
    const normalized = normalizeConfig(sheetName, chartConfig);
    const state = ensurePackageState(this.workbook);
    const sheetPath = state.sheetPaths.get(sheetName);
    const drawing = ensureDrawing(state, sheetPath);
    const chartPath = nextPartPath(state, "xl/charts", "chart");
    const relationshipId = addRelationship(state, drawing.drawingPath, CHART_REL, chartPath);
    const existing = chartEntries(state, sheetName);
    const name = normalized.name ?? `Chart ${existing.length + 1}`;
    let drawingXml = getTextFile(state, drawing.drawingPath);
    const anchor = drawingAnchor(existing.length + 2, name, relationshipId, normalized.position);
    drawingXml = /<xdr:wsDr\b[^>]*\/>/i.test(drawingXml)
      ? drawingXml.replace(/<xdr:wsDr\b([^>]*)\/>/i, `<xdr:wsDr$1>${anchor}</xdr:wsDr>`)
      : drawingXml.replace(/<\/xdr:wsDr\s*>/i, `${anchor}</xdr:wsDr>`);
    setTextFile(state, drawing.drawingPath, drawingXml);
    setTextFile(state, chartPath, chartXml(normalized));
    addContentTypeOverride(state, chartPath, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml");
    return this.list(sheetName).find((item) => item.id === chartPath);
  }

  update(reference, changes) {
    const state = ensurePackageState(this.workbook);
    const entry = findChart(state, this.workbook, reference);
    if (!entry) throw new RangeError(`Chart ${String(reference)} does not exist.`);
    if (!changes.range && !changes.series) throw new TypeError("Updating a chart requires range or series so the data mapping stays explicit.");
    const normalized = normalizeConfig(entry.sheet, { type: changes.type ?? entry.type, title: changes.title ?? entry.title, position: changes.position ?? entry.position, ...changes });
    setTextFile(state, entry.partPath, chartXml(normalized));
    if (changes.position || changes.name) {
      let drawingXml = getTextFile(state, entry.drawingPath);
      const id = Number(allTags(entry.anchorXml, "cNvPr")[0]?.attributes.id ?? 2);
      drawingXml = drawingXml.replace(entry.anchorXml, drawingAnchor(id, changes.name ?? entry.name, entry.relationshipId, normalized.position));
      setTextFile(state, entry.drawingPath, drawingXml);
    }
    return this.list(entry.sheet).find((item) => item.id === entry.id);
  }

  remove(reference) {
    const state = ensurePackageState(this.workbook);
    const entry = findChart(state, this.workbook, reference);
    if (!entry) return false;
    let drawingXml = getTextFile(state, entry.drawingPath);
    drawingXml = drawingXml.replace(entry.anchorXml, "");
    setTextFile(state, entry.drawingPath, drawingXml);
    removeRelationship(state, entry.drawingPath, entry.relationshipId);
    state.files.delete(entry.partPath);
    removeContentTypeOverride(state, entry.partPath);
    return true;
  }
}
