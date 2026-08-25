import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const packageJson = JSON.parse(await fs.readFile(path.join(repoRoot, "package.json"), "utf8"));
const releaseVersion = String(process.env.RELEASE_VERSION || packageJson.version).replace(/^v/, "");
if (!/^\d+\.\d+\.\d+(?:[-+][0-9A-Za-z.-]+)?$/.test(releaseVersion)) {
  throw new Error(`Invalid RELEASE_VERSION: ${releaseVersion}`);
}

const distDir = path.join(repoRoot, "dist");
const dataPath = path.join(repoRoot, "data", "release_blank_quarter_planning.json");
const xlsxPath = path.join(distDir, `QuarterPlan-Excel-v${releaseVersion}-no-macros.xlsx`);
const xlsmPath = path.join(distDir, `QuarterPlan-Excel-v${releaseVersion}.xlsm`);

function run(script, env = {}, args = []) {
  const result = spawnSync(process.execPath, [path.join(repoRoot, "scripts", script), ...args], {
    cwd: repoRoot,
    env: { ...process.env, ...env },
    stdio: "inherit",
  });
  if (result.error) throw result.error;
  if (result.status !== 0) throw new Error(`${script} failed with exit code ${result.status}`);
}

await fs.mkdir(distDir, { recursive: true });
run("build_quarter_planning_step1.mjs", {
  QUARTER_PLANNING_DATA_PATH: dataPath,
  QUARTER_PLANNING_XLSX_OUTPUT: xlsxPath,
});
run("create_quarter_planning_xlsm.mjs", {
  QUARTER_PLANNING_DATA_PATH: dataPath,
  QUARTER_PLANNING_XLSX_INPUT: xlsxPath,
  QUARTER_PLANNING_XLSM_OUTPUT: xlsmPath,
});
run("sanitize_workbook_metadata.mjs", {}, [xlsxPath, xlsmPath]);

async function sha256(filePath) {
  return crypto.createHash("sha256").update(await fs.readFile(filePath)).digest("hex");
}

const checksumLines = [];
for (const filePath of [xlsmPath, xlsxPath]) {
  checksumLines.push(`${await sha256(filePath)}  ${path.basename(filePath)}`);
}
await fs.writeFile(path.join(distDir, "SHA256SUMS.txt"), `${checksumLines.join("\n")}\n`, "utf8");

console.log(`RELEASE BUILT ${releaseVersion}`);
