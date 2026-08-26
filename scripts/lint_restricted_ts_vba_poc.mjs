import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  analyzeRestrictedTypeScriptForVba,
  formatRestrictedEmitterDiagnostic,
} from "./lib/restricted_ts_vba_emitter.mjs";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const safePath = path.join(repoRoot, "poc", "restricted-ts-vba", "src", "planning-engine.ts");
const unsafePath = path.join(repoRoot, "poc", "restricted-ts-vba", "probes", "unsupported.ts");

const safeAnalysis = analyzeRestrictedTypeScriptForVba(await fs.readFile(safePath, "utf8"), safePath);
if (safeAnalysis.diagnostics.length) {
  throw new Error(safeAnalysis.diagnostics.map(formatRestrictedEmitterDiagnostic).join("\n"));
}

const unsafeAnalysis = analyzeRestrictedTypeScriptForVba(await fs.readFile(unsafePath, "utf8"), unsafePath);
const unsafeCodes = new Set(unsafeAnalysis.diagnostics.map(({ code }) => code));
for (const requiredCode of [
  "FUNCTION_RETURN_TYPE",
  "FUNCTION_SHAPE",
  "IMPORT_EXPORT",
  "METHOD_NOT_ALLOWED",
  "MULTIDIMENSIONAL_ARRAY",
  "PARAMETER_TYPE",
  "TOP_LEVEL_STATE",
]) {
  if (!unsafeCodes.has(requiredCode)) {
    throw new Error(`unsafe probe did not fail closed with ${requiredCode}`);
  }
}

console.log("PASS restricted emitter accepts the planning-shaped safe subset");
console.log(`PASS unsupported probe fails closed with ${unsafeAnalysis.diagnostics.length} diagnostics`);
