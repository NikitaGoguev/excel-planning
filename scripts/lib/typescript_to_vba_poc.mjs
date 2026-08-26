import os from "node:os";
import path from "node:path";
import fs from "node:fs";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";

const require = createRequire(import.meta.url);
const { transpileProject } = require("typescript-to-vba");
const { version: transpilerVersion } = require("typescript-to-vba/package.json");

if (transpilerVersion !== "1.0.1") {
  throw new Error(`PoC requires exactly typescript-to-vba 1.0.1; installed ${transpilerVersion}`);
}

export const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "..");
export const pocRoot = path.join(repoRoot, "poc", "typescript-to-vba");
export const tstvbaVersion = transpilerVersion;

function canonicalPath(targetPath) {
  const missingSegments = [];
  let existing = path.resolve(targetPath);
  while (!fs.existsSync(existing)) {
    const parent = path.dirname(existing);
    if (parent === existing) break;
    missingSegments.unshift(path.basename(existing));
    existing = parent;
  }
  const canonicalExisting = fs.existsSync(existing) ? fs.realpathSync.native(existing) : existing;
  return path.resolve(canonicalExisting, ...missingSegments);
}

export function assertSystemTempPath(targetPath) {
  const tempRoot = canonicalPath(os.tmpdir());
  const resolved = canonicalPath(targetPath);
  const comparableRoot = process.platform === "win32" ? tempRoot.toLowerCase() : tempRoot;
  const comparableTarget = process.platform === "win32" ? resolved.toLowerCase() : resolved;
  if (!comparableTarget.startsWith(`${comparableRoot}${path.sep}`)) {
    throw new Error(`PoC output must stay inside the system temp directory: ${resolved}`);
  }
  return resolved;
}

export function transpilePocSource({ entryPath, outDir, outputFileName }) {
  const resolvedEntry = path.resolve(entryPath);
  const resolvedOutDir = assertSystemTempPath(outDir);
  return transpileProject({
    projectFilePath: path.join(pocRoot, "tsconfig.json"),
    entry: resolvedEntry,
    outDir: resolvedOutDir,
    tstvbaOptions: {
      entry: resolvedEntry,
      targetApplication: "Excel",
      moduleStyle: "StandardModule",
      vbaLibraryPath: "./lib/vbalib.bas",
      namespacePrefix: "QP_",
      emitSourceMaps: false,
      bundle: true,
      outputFileName,
    },
  });
}
