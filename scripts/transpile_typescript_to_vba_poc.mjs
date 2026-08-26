import path from "node:path";

import { transpilePocSource } from "./lib/typescript_to_vba_poc.mjs";

const [entryPath, outDir, outputFileName = "PlanningPoc.bas"] = process.argv.slice(2);
if (!entryPath || !outDir) {
  throw new Error("Usage: node scripts/transpile_typescript_to_vba_poc.mjs <entry.ts> <system-temp-out-dir> [output.bas]");
}

const result = transpilePocSource({
  entryPath: path.resolve(entryPath),
  outDir: path.resolve(outDir),
  outputFileName,
});

console.log(result.outputPath);
