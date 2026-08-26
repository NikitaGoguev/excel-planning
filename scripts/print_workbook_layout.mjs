import { deriveWorkbookLayout, loadWorkbookLimits } from "../src/config/workbook-limits.mjs";

process.stdout.write(JSON.stringify(deriveWorkbookLayout(await loadWorkbookLimits())));
