import crypto from "node:crypto";
import fs from "node:fs/promises";

import JSZip from "jszip";

function normalizePart(name, bytes) {
  if (name === "docProps/core.xml") {
    return Buffer.from(bytes.toString("utf8")
      .replace(/<dcterms:created[^>]*>.*?<\/dcterms:created>/g, "<dcterms:created>__TIMESTAMP__</dcterms:created>")
      .replace(/<dcterms:modified[^>]*>.*?<\/dcterms:modified>/g, "<dcterms:modified>__TIMESTAMP__</dcterms:modified>"));
  }
  return bytes;
}

export async function canonicalXlsxFingerprint(input) {
  const bytes = Buffer.isBuffer(input) ? input : await fs.readFile(input);
  const zip = await JSZip.loadAsync(bytes);
  const parts = {};
  for (const name of Object.keys(zip.files).sort()) {
    const entry = zip.files[name];
    if (entry.dir) continue;
    const part = normalizePart(name, await entry.async("nodebuffer"));
    parts[name] = crypto.createHash("sha256").update(part).digest("hex");
  }
  return parts;
}
