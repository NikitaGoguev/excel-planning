import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { loadWorkbookLimits, renderVbaLimitsModule } from "../src/config/workbook-limits.mjs";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outputPath = path.join(repoRoot, "assets", "vba", "QuarterPlanLimits_module.txt");
const expected = renderVbaLimitsModule(await loadWorkbookLimits());
const normalize = (value) => `${value.replace(/\r?\n/g, "\r\n").trimEnd()}\r\n`;

if (process.argv.includes("--check")) {
  const actual = await fs.readFile(outputPath, "utf8").catch(() => "");
  if (normalize(actual) !== normalize(expected)) {
    throw new Error("QuarterPlanLimits_module.txt is out of date. Run npm run generate:limits.");
  }
  console.log("PASS generated VBA limits");
} else {
  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await fs.writeFile(outputPath, normalize(expected), "utf8");
  console.log(`GENERATED ${outputPath}`);
}
