import assert from "node:assert/strict";
import test from "node:test";
import { createWorkbook, normalizeColor, normalizeStyle, parseWorkbook } from "../src/index.js";
import { extractZip } from "../src/zip.js";
import { allTags, tagContent } from "../src/xml.js";

test("friendly styles normalize to the native model", () => {
  assert.deepEqual(normalizeStyle({ bold: true, fontSize: 14, fill: "#ffcc00", horizontal: "center" }), {
    font: { bold: true, size: 14 },
    fill: { type: "pattern", patternType: "solid", foreground: { rgb: "FFCC00" } },
    alignment: { horizontal: "center" }
  });
});

test("malformed colors fail early", () => {
  assert.throws(() => normalizeColor("red"), /Invalid color/);
});

test("supports rich nested styles, gradients, and friendly aliases", () => {
  assert.deepEqual(normalizeStyle({
    font: { name: "Aptos Display", size: 18, verticalAlign: "superscript", color: { theme: 4, tint: 0.25 } },
    fill: { type: "gradient", degree: 90, stops: [{ position: 0, color: "#FFFFFF" }, { position: 1, color: "#17324D" }] },
    alignment: { shrinkToFit: true, indent: 2 },
    rotation: -45,
    editable: true
  }), {
    font: { name: "Aptos Display", size: 18, verticalAlign: "superscript", color: { theme: 4, tint: 0.25 } },
    fill: { type: "gradient", gradientType: "linear", degree: 90, stops: [{ position: 0, color: { rgb: "FFFFFF" } }, { position: 1, color: { rgb: "17324D" } }] },
    alignment: { shrinkToFit: true, indent: 2, textRotation: 135 },
    protection: { locked: false }
  });
});

test("defines, inherits, composes, and persists named Excel styles", () => {
  const workbook = createWorkbook("Report");
  workbook.styles
    .define("reportBase", { fontName: "Aptos", fontSize: 11, vertical: "center" })
    .define("reportHeader", { bold: true, color: "#FFFFFF", fill: "#17324D", horizontal: "center" }, { extends: "reportBase" })
    .define("currency", { numberFormat: "$#,##0.00;[Red]-$#,##0.00" });

  const sheet = workbook.sheet();
  sheet.range("A1:D1").setValues([["Item", "Qty", "Price", "Total"]]).style("reportHeader");
  sheet.cell("C2").set(12.5).style(["currency", { italic: true }]);
  sheet.cell("D2").copyStyleFrom("C2");

  const reopened = parseWorkbook(workbook.toBuffer());
  assert.deepEqual(reopened.styles.names.sort(), ["currency", "reportBase", "reportHeader"]);
  assert.equal(reopened.sheet().cell("A1").raw.namedStyle, "reportHeader");
  assert.equal(reopened.sheet().cell("A1").raw.style.font.bold, true);
  assert.equal(reopened.sheet().cell("A1").raw.style.font.name, "Aptos");
  assert.equal(reopened.sheet().cell("D2").raw.style.numberFormat, "$#,##0.00;[Red]-$#,##0.00");
});

test("applies range outline and interior borders positionally", () => {
  const workbook = createWorkbook("Grid");
  const sheet = workbook.sheet();
  sheet.range("A1:B2").style({ border: {
    outline: { style: "medium", color: "#17324D" },
    inside: { style: "thin", color: "#CCCCCC" }
  } });
  assert.equal(sheet.cell("A1").raw.style.border.top.style, "medium");
  assert.equal(sheet.cell("A1").raw.style.border.right.style, "thin");
  assert.equal(sheet.cell("B2").raw.style.border.bottom.style, "medium");
  assert.equal(sheet.cell("B2").raw.style.border.left.style, "thin");
});

test("copies and selectively clears style properties", () => {
  const workbook = createWorkbook("Styles");
  const sheet = workbook.sheet();
  sheet.cell("A1").style({ bold: true, fill: "#17324D", wrapText: true });
  sheet.cell("B1").copyStyleFrom("A1").clearStyle(["fill", "bold"]);
  assert.equal(sheet.cell("B1").raw.style.fill, undefined);
  assert.equal(sheet.cell("B1").raw.style.font, undefined);
  assert.equal(sheet.cell("B1").raw.style.alignment.wrapText, true);
  sheet.range("A1:B1").clearStyle("wrapText");
  assert.equal(sheet.cell("A1").raw.style.alignment, undefined);
});

test("deduplicates repeated styles in preserved templates", () => {
  const source = createWorkbook("Template");
  source.sheet().cell("A1").set("seed");
  const template = parseWorkbook(source.toBuffer());
  template.sheet().range("A1:J20").style({ bold: true, fill: "#FFF2CC" });
  const files = extractZip(template.toBuffer());
  const stylesXml = files.get("xl/styles.xml").toString("utf8");
  const cellXfs = allTags(tagContent(stylesXml, "cellXfs"), "xf");
  const fonts = allTags(tagContent(stylesXml, "fonts"), "font");
  const fills = allTags(tagContent(stylesXml, "fills"), "fill");
  assert.equal(cellXfs.length, 2);
  assert.equal(fonts.length, 2);
  assert.equal(fills.length, 3);
});

test("removes named styles from preserved workbooks", () => {
  const source = createWorkbook("Report");
  source.styles.define("temporary", { bold: true });
  const edited = parseWorkbook(source.toBuffer());
  assert.equal(edited.styles.remove("temporary"), true);
  const reopened = parseWorkbook(edited.toBuffer());
  assert.equal(reopened.styles.has("temporary"), false);
});
