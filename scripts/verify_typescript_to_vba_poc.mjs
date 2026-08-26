import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { lintRestrictedTypeScript } from "./lib/restricted_typescript_lint.mjs";
import { assertSystemTempPath, pocRoot, transpilePocSource } from "./lib/typescript_to_vba_poc.mjs";

const safeEntry = path.join(pocRoot, "src", "safe-smoke.ts");
const schedulerProbe = path.join(pocRoot, "probes", "scheduler-shape.ts");
const byValProbe = path.join(pocRoot, "probes", "byval-output.ts");
const temporaryDirectory = await fs.mkdtemp(path.join(os.tmpdir(), "tstvba-poc-portable-"));

async function readSource(filePath) {
  return fs.readFile(filePath, "utf8");
}

function codes(diagnostics) {
  return new Set(diagnostics.map((diagnostic) => diagnostic.code));
}

try {
  const safeSource = await readSource(safeEntry);
  assert.deepEqual(lintRestrictedTypeScript(safeSource, "safe-smoke.ts"), []);
  console.log("PASS safe smoke is inside the restricted TypeScript subset");

  const schedulerDiagnostics = codes(
    lintRestrictedTypeScript(await readSource(schedulerProbe), "scheduler-shape.ts"),
  );
  for (const expected of [
    "FORBIDDEN_GLOBAL",
    "FUNCTION_RETURN",
    "HELPER_CALL",
    "METHOD_NOT_ALLOWED",
    "MULTIDIMENSIONAL_ARRAY",
    "UNSUPPORTED_LOOP",
    "UNSUPPORTED_RETURN",
  ]) {
    assert.ok(schedulerDiagnostics.has(expected), `scheduler probe did not report ${expected}`);
  }
  const byValDiagnostics = codes(lintRestrictedTypeScript(await readSource(byValProbe), "byval-output.ts"));
  assert.ok(byValDiagnostics.has("PARAMETER_WRITE"), "output-array mutation was not rejected");
  console.log("PASS scheduler-shaped unsafe constructs fail closed");

  const safeResult = transpilePocSource({
    entryPath: safeEntry,
    outDir: path.join(temporaryDirectory, "safe"),
    outputFileName: "PlanningPocSafe.bas",
  });
  const safeVba = await fs.readFile(safeResult.outputPath, "ascii");
  assert.match(safeVba, /Public Sub QP_safe_smoke_summarize\b/);
  assert.match(safeVba, /Public Sub TS_ConsoleLog\b/);
  assert.match(safeVba, /For itemIndex = 0 To .*UBound\(values\).*Next itemIndex/s);
  assert.match(safeVba, /Call TS_ConsoleLog\(total\)/);
  console.log("PASS unchanged package emits the minimal smoke module");

  const schedulerResult = transpilePocSource({
    entryPath: schedulerProbe,
    outDir: path.join(temporaryDirectory, "scheduler"),
    outputFileName: "PlanningPocSchedulerProbe.bas",
  });
  const schedulerVba = await fs.readFile(schedulerResult.outputPath, "ascii");
  assert.match(schedulerVba, /Public Function QP_scheduler_shape_plannedDuration\b/);
  assert.doesNotMatch(schedulerVba, /QP_scheduler_shape_plannedDuration\s*=/);
  assert.doesNotMatch(schedulerVba, /\bExit Function\b/);
  assert.doesNotMatch(schedulerVba, /\b(?:Do While|While .*Wend|Loop)\b/i);
  assert.match(schedulerVba, /remaining = remaining - 1/);
  assert.match(schedulerVba, /days = plannedDuration\(matrix\[0\]\[0\], 1\)/);
  assert.doesNotMatch(schedulerVba, /days = QP_scheduler_shape_plannedDuration\b/);
  assert.match(schedulerVba, /GeneratedAt: \d{4}-\d{2}-\d{2}T/);
  assert.ok(
    schedulerVba.includes(`Project: ${path.join(pocRoot, "tsconfig.json")}`),
    "generated metadata did not expose the absolute project path",
  );
  console.log("PASS raw output reproduces return/while/helper/matrix/metadata gaps");

  const byValResult = transpilePocSource({
    entryPath: byValProbe,
    outDir: path.join(temporaryDirectory, "byval"),
    outputFileName: "PlanningPocByVal.bas",
  });
  const byValVba = await fs.readFile(byValResult.outputPath, "ascii");
  assert.match(
    byValVba,
    /Public Sub QP_byval_output_mutateResults\(ByVal values As Variant, ByVal results As Variant\)/,
  );
  assert.match(byValVba, /results\(itemIndex\) = values\(itemIndex\) \+ 1/);
  console.log("PASS output-array probe is emitted as ByVal Variant");

  console.log("NO-GO for planning engine: unchanged typescript-to-vba 1.0.1 lacks required semantics");
} finally {
  const resolved = assertSystemTempPath(temporaryDirectory);
  if (!path.basename(resolved).startsWith("tstvba-poc-portable-")) {
    throw new Error(`Refusing to remove unexpected PoC directory: ${resolved}`);
  }
  await fs.rm(resolved, { recursive: true, force: true });
}
