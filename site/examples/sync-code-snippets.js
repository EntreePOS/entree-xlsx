import { readFile, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const examplesDirectory = dirname(fileURLToPath(import.meta.url));
const siteDirectory = join(examplesDirectory, "..");
const pagePath = join(siteDirectory, "index.html");
const checkOnly = process.argv.includes("--check");

const sourceFiles = [
  "01-add-data.js",
  "02-export-records.js",
  "03-change-cells.js",
  "04-multiple-sheets.js",
  "05-first-style.js",
  "06-reusable-styles.js",
  "07-formulas-and-formats.js",
  "08-layout-and-filters.js",
  "09-edit-a-template.js",
  "10-format-dates-and-percentages.js",
  "11-create-a-chart.js",
  "12-edit-a-chart.js",
  "13-create-a-pivot-table.js",
  "14-protect-a-sheet.js",
  "15-encrypt-a-workbook.js"
];

function escapeHtml(source) {
  return source
    .trimEnd()
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

let page = await readFile(pagePath, "utf8");
const originalPage = page;
for (const [index, fileName] of sourceFiles.entries()) {
  const lesson = index + 1;
  const sourceText = await readFile(join(examplesDirectory, "code", fileName), "utf8");
  const lineCount = sourceText.trimEnd().split("\n").length;
  if (lesson === 1 && lineCount > 15) throw new Error("Lesson 1 must stay beginner friendly and fit within 15 lines.");
  if (lesson <= 5 && lineCount > 30) throw new Error(`Foundation lesson ${lesson} must fit within 30 lines.`);
  if (!sourceText.includes('from "@entree_pos/xlsx"') || !sourceText.includes(".save(")) {
    throw new Error(`Lesson ${lesson} is not a complete runnable example.`);
  }
  const source = escapeHtml(sourceText);
  const pattern = new RegExp(
    `(<article class="[^"]*\\blesson\\b[^"]*" id="lesson-${lesson}">[\\s\\S]*?<pre><code>)[\\s\\S]*?(</code></pre>)`
  );

  if (!pattern.test(page)) {
    throw new Error(`Could not find the code block for lesson ${lesson}`);
  }
  page = page.replace(pattern, `$1${source}$2`);
}

if (checkOnly) {
  if (page !== originalPage) {
    throw new Error("Tutorial code blocks are out of sync. Run npm run tutorial:sync.");
  }
  console.log("Tutorial code blocks match the runnable example files.");
} else {
  await writeFile(pagePath, page);
  console.log("Updated tutorial code blocks from the runnable example files.");
}
