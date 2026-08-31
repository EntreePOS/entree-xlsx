import { spawn } from "node:child_process";
import { mkdir, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { openWorkbook } from "@entree_pos/xlsx";

const examplesDirectory = dirname(fileURLToPath(import.meta.url));
const siteDirectory = join(examplesDirectory, "..");
const codeDirectory = join(examplesDirectory, "code");
const outputDirectory = join(siteDirectory, "assets", "examples");
const previewDirectory = join(siteDirectory, "..", ".codex-tmp-tutorial");

await mkdir(outputDirectory, { recursive: true });
await mkdir(previewDirectory, { recursive: true });

async function runExample(fileName) {
  await new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [join(codeDirectory, fileName)], {
      cwd: outputDirectory,
      stdio: "inherit"
    });
    child.once("error", reject);
    child.once("exit", (code) => {
      if (code === 0) resolve();
      else reject(new Error(`${fileName} exited with code ${code}`));
    });
  });
}

async function createSecurityPreview() {
  const workbook = await openWorkbook(
    join(outputDirectory, "15-encrypt-a-workbook.xlsx"),
    { password: "demo" }
  );
  await writeFile(join(previewDirectory, "15-security-preview.xlsx"), workbook.toBuffer());
}

for (const fileName of [
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
]) {
  await runExample(fileName);
}

await createSecurityPreview();

console.log(`Created tutorial workbooks in ${outputDirectory}`);
