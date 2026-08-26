import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { emitRestrictedTypeScriptFileToVba } from "./lib/restricted_ts_vba_emitter.mjs";

async function canonicalPath(targetPath) {
  const missingSegments = [];
  let existing = path.resolve(targetPath);
  while (true) {
    try {
      const real = await fs.realpath(existing);
      return path.resolve(real, ...missingSegments);
    } catch (error) {
      if (error?.code !== "ENOENT") throw error;
      const parent = path.dirname(existing);
      if (parent === existing) return path.resolve(existing, ...missingSegments);
      missingSegments.unshift(path.basename(existing));
      existing = parent;
    }
  }
}

const [sourceArgument, outputDirectoryArgument, moduleName = "QpRestrictedPlanningEngine"] = process.argv.slice(2);
if (!sourceArgument || !outputDirectoryArgument) {
  throw new Error("usage: node scripts/emit_restricted_ts_vba_poc.mjs <source.ts> <system-temp-output-dir> [module-name]");
}

const sourcePath = path.resolve(sourceArgument);
const outputDirectory = await canonicalPath(outputDirectoryArgument);
const tempRoot = await canonicalPath(os.tmpdir());
const comparableRoot = process.platform === "win32" ? tempRoot.toLowerCase() : tempRoot;
const comparableOutput = process.platform === "win32" ? outputDirectory.toLowerCase() : outputDirectory;
if (!comparableOutput.startsWith(`${comparableRoot}${path.sep}`)) {
  throw new Error(`PoC output must remain inside the system temp directory: ${outputDirectory}`);
}

const result = emitRestrictedTypeScriptFileToVba(sourcePath, { moduleName, prefix: "QPT_" });
await fs.mkdir(outputDirectory, { recursive: true });
const vbaPath = path.join(outputDirectory, `${moduleName}.bas`);
const manifestPath = path.join(outputDirectory, `${moduleName}.manifest.json`);
await fs.writeFile(vbaPath, result.vba, "ascii");
await fs.writeFile(manifestPath, `${JSON.stringify(result.manifest, null, 2)}\n`, "utf8");
console.log(vbaPath);
console.log(manifestPath);
