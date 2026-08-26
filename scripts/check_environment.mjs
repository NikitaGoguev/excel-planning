import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import ExcelJS from "exceljs";
import JSZip from "jszip";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const nodeMajor = Number(process.versions.node.split(".")[0]);
if (nodeMajor !== 24) {
  throw new Error(`Node.js 24.x is required; current version is ${process.versions.node}`);
}

for (const relativePath of [
  "assets/Шаблон Экспресс оценки.xlsx",
  "assets/vba/vbaProject.step2.bin",
  "config/workbook-limits.json",
  "contracts/vba.contract.json",
  "data/test_data_quarter_planning.json",
]) {
  await fs.access(path.join(repoRoot, relativePath));
}

const vbaContract = JSON.parse(await fs.readFile(path.join(repoRoot, "contracts/vba.contract.json"), "utf8"));
for (const component of vbaContract.components ?? []) {
  await fs.access(path.join(repoRoot, component.source));
}

if (typeof ExcelJS.Workbook !== "function" || typeof JSZip.loadAsync !== "function") {
  throw new Error("Portable XLSX dependencies failed to load");
}

console.log(`ENVIRONMENT OK node=${process.versions.node} platform=${process.platform} arch=${process.arch}`);
