import fs from "node:fs/promises";
import path from "node:path";

import { canonicalXlsxFingerprint } from "../src/ooxml/fingerprint.mjs";

const repoRoot = path.resolve(import.meta.dirname, "..");
const workbookPath = path.join(repoRoot, "outputs/quarter_planning_step1.xlsx");
const snapshotPath = path.join(repoRoot, "baselines/xlsx_package_fingerprint.json");
const actual = canonicalXlsxFingerprint(workbookPath);

if (process.argv.includes("--update")) {
  await fs.writeFile(snapshotPath, `${JSON.stringify(await actual, null, 2)}\n`, "utf8");
  console.log(`UPDATED ${snapshotPath}`);
} else {
  const expected = JSON.parse(await fs.readFile(snapshotPath, "utf8"));
  const resolvedActual = await actual;
  if (JSON.stringify(resolvedActual) !== JSON.stringify(expected)) {
    const names = [...new Set([...Object.keys(expected), ...Object.keys(resolvedActual)])];
    const changed = names.filter((name) => expected[name] !== resolvedActual[name]);
    throw new Error(`Canonical XLSX fingerprint drifted in: ${changed.join(", ")}`);
  }
  console.log("PASS canonical XLSX fingerprint");
}
