import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

import ExcelJS from "exceljs";
import JSZip from "jszip";
import { deriveWorkbookLayout, loadWorkbookLimits } from "../src/config/workbook-limits.mjs";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const packageJson = JSON.parse(await fs.readFile(path.join(repoRoot, "package.json"), "utf8"));
const releaseVersion = String(process.env.RELEASE_VERSION || packageJson.version).replace(/^v/, "");
const distDir = path.join(repoRoot, "dist");
const xlsxPath = path.join(distDir, `QuarterPlan-Excel-v${releaseVersion}-no-macros.xlsx`);
const xlsmPath = path.join(distDir, `QuarterPlan-Excel-v${releaseVersion}.xlsm`);
const layout = deriveWorkbookLayout(await loadWorkbookLimits());

const forbiddenPatterns = [
  new RegExp(["gaz", "prombank"].join(""), "iu"),
  new RegExp(["jira", "dev"].join(""), "iu"),
  new RegExp(["DBO", "CORPESPLN"].join(""), "iu"),
  new RegExp(["LINE", "AUSN"].join(""), "iu"),
  new RegExp(["Fox", "trot"].join(""), "iu"),
  new RegExp(["АУСН", "\\s*-\\s*", "Разделение процесса расчета"].join(""), "iu"),
  new RegExp(["Ускорение", "\\s*-\\s*", "Выписки"].join(""), "iu"),
];
const ownerPattern = new RegExp(["Go", "guev"].join(""), "iu");
const packageExtensions = new Set([".xlsx", ".xlsm", ".zip"]);
const textExtensions = new Set([".csv", ".json", ".md", ".mjs", ".ps1", ".txt", ".yml", ".yaml", ".xml"]);

function assertNoForbidden(text, context, includeOwner = false) {
  for (const pattern of forbiddenPatterns) {
    if (pattern.test(text)) throw new Error(`Forbidden public data ${pattern} found in ${context}`);
  }
  if (includeOwner && ownerPattern.test(text)) throw new Error(`Demo owner name found in ${context}`);
}

async function scanPackage(filePath, includeOwner = true) {
  const zip = await JSZip.loadAsync(await fs.readFile(filePath));
  for (const [entryName, entry] of Object.entries(zip.files)) {
    if (entry.dir) continue;
    const bytes = await entry.async("nodebuffer");
    const combined = `${bytes.toString("utf8")}\n${bytes.toString("latin1")}\n${bytes.toString("utf16le")}`;
    assertNoForbidden(combined, `${filePath}:${entryName}`, includeOwner && entryName !== "xl/vbaProject.bin");
  }
}

function trackedFiles() {
  const result = spawnSync("git", ["-c", "core.quotepath=false", "ls-files", "-z"], {
    cwd: repoRoot,
    encoding: "utf8",
  });
  if (result.status !== 0) throw new Error(result.stderr || "git ls-files failed");
  return result.stdout.split("\0").filter(Boolean);
}

for (const relativePath of trackedFiles()) {
  const fullPath = path.join(repoRoot, relativePath);
  try {
    await fs.access(fullPath);
  } catch {
    continue;
  }
  const extension = path.extname(relativePath).toLowerCase();
  if (packageExtensions.has(extension)) {
    await scanPackage(fullPath, /^(?:assets|outputs)[\\/]/.test(relativePath));
  } else if (textExtensions.has(extension)) {
    const includeOwner = /^(?:assets|data|outputs)[\\/]/.test(relativePath);
    assertNoForbidden(await fs.readFile(fullPath, "utf8"), relativePath, includeOwner);
  } else if (extension === ".bin") {
    const bytes = await fs.readFile(fullPath);
    assertNoForbidden(`${bytes.toString("latin1")}\n${bytes.toString("utf16le")}`, relativePath, false);
  }
}

await scanPackage(xlsxPath, true);
await scanPackage(xlsmPath, true);

const xlsxZip = await JSZip.loadAsync(await fs.readFile(xlsxPath));
const xlsmZip = await JSZip.loadAsync(await fs.readFile(xlsmPath));
if (xlsxZip.file("xl/vbaProject.bin")) throw new Error("Release XLSX unexpectedly contains VBA");
if (!xlsmZip.file("xl/vbaProject.bin")) throw new Error("Release XLSM is missing VBA");

const workbook = new ExcelJS.Workbook();
await workbook.xlsx.readFile(xlsxPath);
const settings = workbook.getWorksheet("00_Настройки");
const estimates = workbook.getWorksheet("03_Оценка задач");
const plan = workbook.getWorksheet("04_Квартальный план");
if (!settings || !estimates || !plan) throw new Error("Release workbook is missing required sheets");
for (const cell of ["B4", "B5", "B6", "B7"]) {
  const value = settings.getCell(cell).value;
  if (value !== null && value !== "") throw new Error(`Release setting ${cell} is not blank: ${value}`);
}
for (let row = layout.taskEstimates.dataStartRow; row <= layout.taskEstimates.dataEndRow; row += 1) {
  for (let column = 3; column <= 12; column += 1) {
    const value = estimates.getCell(row, column).value;
    if (value !== null && value !== "") throw new Error(`Release estimate cell ${estimates.getCell(row, column).address} is not blank`);
  }
}
for (let row = layout.backlog.dataStartRow; row <= layout.backlog.dataEndRow; row += 1) {
  const value = plan.getCell(row, 8).value;
  if (value !== null && value !== "") throw new Error(`Release backlog description H${row} is not blank`);
}

const checksumPath = path.join(distDir, "SHA256SUMS.txt");
const checksumText = await fs.readFile(checksumPath, "utf8");
for (const filePath of [xlsmPath, xlsxPath]) {
  const digest = crypto.createHash("sha256").update(await fs.readFile(filePath)).digest("hex");
  const expectedLine = `${digest}  ${path.basename(filePath)}`;
  if (!checksumText.split(/\r?\n/).includes(expectedLine)) throw new Error(`Missing checksum line: ${expectedLine}`);
}

console.log("PASS public data scan");
console.log("PASS clean release workbook profile");
console.log("PASS release package and checksums");
