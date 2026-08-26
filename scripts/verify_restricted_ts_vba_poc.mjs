import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { emitRestrictedTypeScriptFileToVba } from "./lib/restricted_ts_vba_emitter.mjs";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const sourcePath = path.join(repoRoot, "poc", "restricted-ts-vba", "src", "planning-engine.ts");
const temporaryDirectory = await fs.mkdtemp(path.join(os.tmpdir(), "restricted-ts-vba-poc-"));

try {
  const options = { moduleName: "QpRestrictedPlanningEngine", prefix: "QPT_" };
  const first = emitRestrictedTypeScriptFileToVba(sourcePath, options);
  const second = emitRestrictedTypeScriptFileToVba(sourcePath, options);
  assert.equal(first.vba, second.vba);
  assert.deepEqual(first.manifest, second.manifest);
  console.log("PASS repeated emission is byte-for-byte deterministic");

  const vbaPath = path.join(temporaryDirectory, "QpRestrictedPlanningEngine.bas");
  const manifestPath = path.join(temporaryDirectory, "QpRestrictedPlanningEngine.manifest.json");
  await fs.writeFile(vbaPath, first.vba, "ascii");
  await fs.writeFile(manifestPath, `${JSON.stringify(first.manifest, null, 2)}\n`, "utf8");

  const vba = await fs.readFile(vbaPath, "ascii");
  assert.match(vba, /QPT_plannedDuration = QPT_plannedDurationInternal\(estimate, focusFactor\)/);
  assert.match(vba, /Do While \(remaining > 0\)[\s\S]*remaining = remaining - 1[\s\S]*Loop/);
  assert.match(vba, /Public Sub QPT_scheduleDurations\(ByRef estimates As Variant, ByVal focusFactor As Double, ByRef results As Variant\)/);
  assert.match(vba, /results\(itemIndex\) = QPT_plannedDurationInternal\(estimates\(itemIndex\), focusFactor\)/);
  assert.doesNotMatch(vba, /GeneratedAt|Project:|[A-Z]:\\|\/home\//i);
  console.log("PASS return, while, helper calls, and ByRef array output are emitted explicitly");
  console.log("PASS generated VBA and manifest contain no timestamp or source path");
  console.log("GO for a restricted-emitter scheduler spike; production parity still requires full scheduler fixtures");
} finally {
  const resolved = path.resolve(temporaryDirectory);
  const tempRoot = path.resolve(os.tmpdir());
  if (!resolved.startsWith(`${tempRoot}${path.sep}`) || !path.basename(resolved).startsWith("restricted-ts-vba-poc-")) {
    throw new Error(`Refusing to remove unexpected PoC directory: ${resolved}`);
  }
  await fs.rm(resolved, { recursive: true, force: true });
}
