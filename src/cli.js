#!/usr/bin/env node
import { extname, resolve } from "node:path";
import { writeFile } from "node:fs/promises";
import { openWorkbook } from "./index.js";

function help() {
  console.log(`entree-xlsx

Usage:
  entree-xlsx inspect <input.xlsx>
  entree-xlsx convert <input.xlsx> <output.csv|json|xlsx> [--sheet <name>]

Examples:
  entree-xlsx inspect sales.xlsx
  entree-xlsx convert sales.xlsx sales.json --sheet Orders
  entree-xlsx convert sales.xlsx orders.csv --sheet Orders`);
  process.exit(0);
}

function option(name) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : undefined;
}

async function main() {
  const [, , command, input, output] = process.argv;
  if (!command || command === "--help" || command === "-h") help();
  if (!input) throw new Error("An input workbook path is required. Run entree-xlsx --help for usage.");
  const workbook = await openWorkbook(resolve(input));

  if (command === "inspect") {
    console.log(JSON.stringify({
      file: resolve(input),
      sheets: workbook.sheetNames.map((name) => ({ name, range: workbook.sheet(name).usedRange }))
    }, null, 2));
    return;
  }

  if (command !== "convert" || !output) help();
  const selected = option("--sheet");
  const extension = extname(output).toLowerCase();
  if (extension === ".json") {
    const value = selected ? workbook.sheet(selected).toRecords() : workbook.toJSON();
    await writeFile(resolve(output), `${JSON.stringify(value, null, 2)}\n`, "utf8");
  } else if (extension === ".csv") {
    await writeFile(resolve(output), workbook.sheet(selected ?? 0).toCsv(), "utf8");
  } else {
    await workbook.save(resolve(output));
  }
  console.log(`Wrote ${resolve(output)}`);
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exitCode = 1;
});
