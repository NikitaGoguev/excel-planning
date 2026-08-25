import fs from "node:fs/promises";
import path from "node:path";

import JSZip from "jszip";

const PRODUCT_NAME = "QuarterPlan Excel";

function replaceElementText(xml, localName, value) {
  const pattern = new RegExp(`(<(?:[A-Za-z0-9]+:)?${localName}\\b[^>]*>)[\\s\\S]*?(<\\/(?:[A-Za-z0-9]+:)?${localName}>)`, "g");
  return xml.replace(pattern, `$1${value}$2`);
}

async function sanitizeWorkbookMetadata(workbookPath) {
  const absolutePath = path.resolve(workbookPath);
  const zip = await JSZip.loadAsync(await fs.readFile(absolutePath));

  const core = zip.file("docProps/core.xml");
  if (core) {
    let xml = await core.async("string");
    xml = replaceElementText(xml, "creator", PRODUCT_NAME);
    xml = replaceElementText(xml, "lastModifiedBy", PRODUCT_NAME);
    zip.file("docProps/core.xml", xml);
  }

  const app = zip.file("docProps/app.xml");
  if (app) {
    let xml = await app.async("string");
    xml = replaceElementText(xml, "Company", PRODUCT_NAME);
    xml = replaceElementText(xml, "Manager", "");
    zip.file("docProps/app.xml", xml);
  }

  const workbook = zip.file("xl/workbook.xml");
  if (workbook) {
    const xml = await workbook.async("string");
    zip.file("xl/workbook.xml", xml.replace(/(<x15ac:absPath\b[^>]*\burl=")[^"]*(")/g, "$1$2"));
  }

  await fs.writeFile(
    absolutePath,
    await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 6 } }),
  );
  console.log(`SANITIZED METADATA ${absolutePath}`);
}

const workbookPaths = process.argv.slice(2);
if (workbookPaths.length === 0) {
  console.error("Usage: node scripts/sanitize_workbook_metadata.mjs <workbook.xlsx|workbook.xlsm> [...]");
  process.exit(1);
}

for (const workbookPath of workbookPaths) await sanitizeWorkbookMetadata(workbookPath);
