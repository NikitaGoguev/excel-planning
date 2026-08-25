import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

if (process.platform !== "win32") {
  throw new Error("This command requires Windows desktop Excel and Windows PowerShell");
}

const [scriptName, ...scriptArgs] = process.argv.slice(2);
if (!scriptName) throw new Error("PowerShell script name is required");

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const scriptPath = path.join(repoRoot, "scripts", scriptName);
const result = spawnSync("powershell", [
  "-NoProfile",
  "-ExecutionPolicy",
  "Bypass",
  "-File",
  scriptPath,
  ...scriptArgs,
], {
  cwd: repoRoot,
  env: { ...process.env, QUARTER_PLANNING_NODE_EXECUTABLE: process.execPath },
  stdio: "inherit",
});

if (result.error) throw result.error;
process.exit(result.status ?? 1);
