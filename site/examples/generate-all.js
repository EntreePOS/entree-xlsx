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
    join(outputDirectory, "08-protection-and-encryption.xlsx"),
    { password: "demo" }
  );
  await writeFile(join(previewDirectory, "08-security-preview.xlsx"), workbook.toBuffer());
}

for (const fileName of [
  "01-create-workbook.js",
  "02-reusable-styles.js",
  "03-formulas-and-formats.js",
  "04-layout-and-filters.js",
  "05-edit-a-template.js",
  "06-create-a-chart.js",
  "07-create-a-pivot-table.js",
  "08-protection-and-encryption.js"
]) {
  await runExample(fileName);
}

await createSecurityPreview();

console.log(`Created tutorial workbooks in ${outputDirectory}`);
