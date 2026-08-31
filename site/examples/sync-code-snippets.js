import { readFile, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const examplesDirectory = dirname(fileURLToPath(import.meta.url));
const siteDirectory = join(examplesDirectory, "..");
const pagePath = join(siteDirectory, "index.html");
const checkOnly = process.argv.includes("--check");

const sourceFiles = [
  "01-create-workbook.js",
  "02-reusable-styles.js",
  "03-formulas-and-formats.js",
  "04-layout-and-filters.js",
  "05-edit-a-template.js",
  "06-create-a-chart.js",
  "07-create-a-pivot-table.js",
  "08-protection-and-encryption.js"
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
let previousLineCount = 0;

for (const [index, fileName] of sourceFiles.entries()) {
  const lesson = index + 1;
  const sourceText = await readFile(join(examplesDirectory, "code", fileName), "utf8");
  const lineCount = sourceText.trimEnd().split("\n").length;
  if (lineCount <= previousLineCount) {
    throw new Error(`Lesson ${lesson} must be more detailed than lesson ${lesson - 1}.`);
  }
  if (!sourceText.includes('from "@entree_pos/xlsx"') || !sourceText.includes(".save(")) {
    throw new Error(`Lesson ${lesson} is not a complete runnable example.`);
  }
  previousLineCount = lineCount;
  const source = escapeHtml(sourceText);
  const pattern = new RegExp(
    `(<article class="lesson" id="lesson-${lesson}">[\\s\\S]*?<pre><code>)[\\s\\S]*?(</code></pre>)`
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
