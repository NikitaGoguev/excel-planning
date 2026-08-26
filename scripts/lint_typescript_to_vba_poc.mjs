import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  formatRestrictedTypeScriptDiagnostic,
  lintRestrictedTypeScript,
  relativeDiagnosticPath,
} from "./lib/restricted_typescript_lint.mjs";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const defaultSource = path.join(repoRoot, "poc", "typescript-to-vba", "src", "safe-smoke.ts");
const sourcePaths = process.argv.slice(2).length ? process.argv.slice(2) : [defaultSource];
let failed = false;

for (const sourcePath of sourcePaths) {
  const resolved = path.resolve(repoRoot, sourcePath);
  const displayPath = relativeDiagnosticPath(resolved, repoRoot);
  const source = await fs.readFile(resolved, "utf8");
  const diagnostics = lintRestrictedTypeScript(source, displayPath);
  if (diagnostics.length) {
    failed = true;
    for (const diagnostic of diagnostics) console.error(formatRestrictedTypeScriptDiagnostic(diagnostic));
  } else {
    console.log(`PASS restricted TypeScript lint: ${displayPath}`);
  }
}

if (failed) process.exitCode = 1;
