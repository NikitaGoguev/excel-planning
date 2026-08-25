import fs from "node:fs/promises";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");
const sourceXlsxPath = process.env.QUARTER_PLANNING_XLSX_INPUT
  ? path.resolve(process.env.QUARTER_PLANNING_XLSX_INPUT)
  : path.join(repoRoot, "outputs", "quarter_planning_step1.xlsx");
const outputXlsmPath = process.env.QUARTER_PLANNING_XLSM_OUTPUT
  ? path.resolve(process.env.QUARTER_PLANNING_XLSM_OUTPUT)
  : path.join(repoRoot, "outputs", "quarter_planning_step2.xlsm");
const baseVbaProjectPath = path.join(repoRoot, "assets", "vba", "vbaProject.step2.bin");
const testDataPath = process.env.QUARTER_PLANNING_DATA_PATH
  ? path.resolve(process.env.QUARTER_PLANNING_DATA_PATH)
  : path.join(repoRoot, "data", "test_data_quarter_planning.json");

function excelSerialFromIso(isoDate) {
  const [year, month, day] = isoDate.split("-").map(Number);
  const utc = Date.UTC(year, month - 1, day);
  const excelEpoch = Date.UTC(1899, 11, 30);
  return Math.round((utc - excelEpoch) / 86400000);
}

function firstWeekdays(startIso, count) {
  const result = [];
  const date = new Date(`${startIso}T00:00:00Z`);
  while (result.length < count) {
    const day = date.getUTCDay();
    if (day >= 1 && day <= 5) result.push(date.toISOString().slice(0, 10));
    date.setUTCDate(date.getUTCDate() + 1);
  }
  return result;
}

function xmlPrefix(xml, rootName) {
  return xml.includes(`<x:${rootName}`) ? "x:" : "";
}

function replaceXmlCell(xml, cellRef, replacement) {
  const empty = new RegExp(`<(?:x:)?c r="${cellRef}"[^>]*/>`);
  if (empty.test(xml)) return xml.replace(empty, replacement);
  const full = new RegExp(`<(?:x:)?c r="${cellRef}"[^>]*>[\\s\\S]*?<\\/(?:x:)?c>`);
  if (full.test(xml)) return xml.replace(full, replacement);
  return xml;
}

function xmlCellStyle(xml, cellRef) {
  const match = xml.match(new RegExp(`<(?:x:)?c r="${cellRef}"[^>]*\\bs="([^"]+)"`));
  return match ? ` s="${match[1]}"` : "";
}

function prepareMacroSheetXml(sheetXml, testData) {
  const holidayCount = Number(testData.quarterSettings?.expectedWeekdayHolidaysToFill ?? 0);
  const calendarWeekdaysValue = testData.quarterSettings?.calendarWeekdays;
  const workingDaysValue = testData.quarterSettings?.workingDays;
  const hasCalendarWeekdays = calendarWeekdaysValue !== "" && Number.isFinite(Number(calendarWeekdaysValue));
  const hasWorkingDays = workingDaysValue !== "" && Number.isFinite(Number(workingDaysValue));
  const calendarWeekdays = hasCalendarWeekdays ? Number(calendarWeekdaysValue) : null;
  const workingDays = hasWorkingDays ? Number(workingDaysValue) : null;
  const startDate = testData.quarterSettings?.startDate ?? "";
  const dates = holidayCount > 0 && startDate ? firstWeekdays(startDate, holidayCount) : [];
  const p = xmlPrefix(sheetXml, "worksheet");

  let xml = sheetXml;
  const b7Style = xmlCellStyle(xml, "B7");
  const b8Style = xmlCellStyle(xml, "B8");
  const b7Cell = hasCalendarWeekdays
    ? `<${p}c r="B7"${b7Style} t="n"><${p}v>${calendarWeekdays}</${p}v></${p}c>`
    : `<${p}c r="B7"${b7Style} t="n" />`;
  const b8Cell = hasCalendarWeekdays && hasWorkingDays
    ? `<${p}c r="B8"${b8Style} t="n"><${p}v>${Math.max(0, calendarWeekdays - workingDays)}</${p}v></${p}c>`
    : `<${p}c r="B8"${b8Style} t="n" />`;
  xml = replaceXmlCell(xml, "B7", b7Cell);
  xml = replaceXmlCell(xml, "B8", b8Cell);

  for (let row = 14; row <= 33; row += 1) {
    const idx = row - 14;
    const aStyle = xmlCellStyle(xml, `A${row}`);
    const bStyle = xmlCellStyle(xml, `B${row}`);
    const cStyle = xmlCellStyle(xml, `C${row}`);
    const dStyle = xmlCellStyle(xml, `D${row}`);
    const dateCell =
      idx < dates.length
        ? `<${p}c r="A${row}"${aStyle} t="n"><${p}v>${excelSerialFromIso(dates[idx])}</${p}v></${p}c>`
        : `<${p}c r="A${row}"${aStyle} t="n" />`;
    const yesCell =
      idx < dates.length
        ? `<${p}c r="D${row}"${dStyle} t="str"><${p}v>Да</${p}v></${p}c>`
        : `<${p}c r="D${row}"${dStyle} t="str" />`;
    xml = replaceXmlCell(xml, `A${row}`, dateCell);
    xml = replaceXmlCell(xml, `B${row}`, `<${p}c r="B${row}"${bStyle} t="str" />`);
    xml = replaceXmlCell(xml, `C${row}`, `<${p}c r="C${row}"${cStyle} t="str" />`);
    xml = replaceXmlCell(xml, `D${row}`, yesCell);
  }
  return xml;
}

function updateContentTypes(xml) {
  let out = xml.replace(
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
  );
  const vbaDefault = '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>';
  if (out.includes('Default Extension="bin"')) {
    out = out.replace(/<Default Extension="bin" ContentType="[^"]+"\s*\/>/, vbaDefault);
  } else {
    out = out.replace("</Types>", `${vbaDefault}</Types>`);
  }
  return out;
}

function updateWorkbookRels(xml) {
  if (xml.includes("relationships/vbaProject")) return xml;
  const ids = Array.from(xml.matchAll(/Id="rId(\d+)"/g), (match) => Number(match[1]));
  const nextId = Math.max(0, ...ids) + 1;
  return xml.replace(
    "</Relationships>",
    `<Relationship Id="rId${nextId}" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>`,
  );
}

function updateWorkbookXml(xml) {
  let out = xml;
  out = out.replace(/<fileVersion\b(?![^>]*\bcodeName=)/, '<fileVersion codeName="{00000000-0000-0000-0000-000000000000}"');
  if (out.includes("codeName=\"ThisWorkbook\"")) return out;
  const workbookPr = xml.match(/<((?:[A-Za-z0-9]+:)?)workbookPr\b/);
  if (workbookPr) {
    return out.replace(workbookPr[0], `${workbookPr[0]} codeName="ThisWorkbook"`);
  }
  const sheets = out.match(/<((?:[A-Za-z0-9]+:)?)sheets\b[^>]*>/);
  if (!sheets) return out;
  const p = sheets[1] ?? "";
  return out.replace(sheets[0], `<${p}workbookPr codeName="ThisWorkbook" />${sheets[0]}`);
}

function removePrinterSettingsRelationships(xml) {
  return xml
    .replace(
      /<Relationship\b[^>]*Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/printerSettings"[^>]*\/>/g,
      "",
    )
    .replace(/<Relationships xmlns="http:\/\/schemas\.openxmlformats\.org\/package\/2006\/relationships">\s*<\/Relationships>/, "");
}

function removeCalcChainRelationship(xml) {
  return xml.replace(
    /<Relationship\b[^>]*Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/calcChain"[^>]*\/>/g,
    "",
  );
}

function removeCalcChainContentType(xml) {
  return xml.replace(
    /<Override PartName="\/xl\/calcChain\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.calcChain\+xml"\s*\/>/g,
    "",
  );
}

function updateHolidayTableXml(xml, holidayCount) {
  const activeRows = Math.max(1, Number(holidayCount) || 0);
  const ref = `A13:D${13 + activeRows}`;
  return xml
    .replace(/ref="A13:D\d+"/g, `ref="${ref}"`)
    .replace(/<(x:)?autoFilter ref="A13:D\d+"\s*\/>/, (_match, prefix = "") => `<${prefix}autoFilter ref="${ref}" />`);
}

function updateWorksheetCodeName(xml, codeName) {
  if (xml.includes("codeName=")) return xml;
  const sheetPr = xml.match(/<((?:[A-Za-z0-9]+:)?)sheetPr\b/);
  if (sheetPr) {
    return xml.replace(sheetPr[0], `${sheetPr[0]} codeName="${codeName}"`);
  }
  const worksheetOpen = xml.match(/<((?:[A-Za-z0-9]+:)?)worksheet\b[^>]*>/)?.[0];
  if (!worksheetOpen) return xml;
  const p = worksheetOpen.startsWith("<x:") ? "x:" : "";
  return xml.replace(worksheetOpen, `${worksheetOpen}<${p}sheetPr codeName="${codeName}" />`);
}

function normalizeWorkbookDesign(workbookPath) {
  const result = spawnSync(process.execPath, [path.join(repoRoot, "scripts", "normalize_workbook_design.mjs"), workbookPath], {
    cwd: repoRoot,
    encoding: "utf8",
  });
  if (result.status !== 0) {
    throw new Error(`Workbook design normalization failed: ${result.stderr || result.stdout}`);
  }
}

const testData = JSON.parse(await fs.readFile(testDataPath, "utf8"));
const sourceBytes = await fs.readFile(sourceXlsxPath);
const zip = await JSZip.loadAsync(sourceBytes);
const worksheetIndexes = Object.keys(zip.files)
  .map((name) => name.match(/^xl\/worksheets\/sheet(\d+)\.xml$/)?.[1])
  .filter(Boolean)
  .map(Number)
  .sort((a, b) => a - b);

zip.remove("xl/printerSettings/printerSettings1.bin");
zip.remove("xl/calcChain.xml");
for (const relEntry of Object.keys(zip.files).filter((name) => /^xl\/worksheets\/_rels\/sheet\d+\.xml\.rels$/.test(name))) {
  const updatedRels = removePrinterSettingsRelationships(await zip.file(relEntry).async("string"));
  if (updatedRels.trim()) {
    zip.file(relEntry, updatedRels);
  } else {
    zip.remove(relEntry);
  }
}

zip.file("xl/vbaProject.bin", await fs.readFile(baseVbaProjectPath));

zip.file("[Content_Types].xml", removeCalcChainContentType(updateContentTypes(await zip.file("[Content_Types].xml").async("string"))));
zip.file("xl/workbook.xml", updateWorkbookXml(await zip.file("xl/workbook.xml").async("string")));
zip.file("xl/_rels/workbook.xml.rels", updateWorkbookRels(removeCalcChainRelationship(await zip.file("xl/_rels/workbook.xml.rels").async("string"))));

for (const sheetIndex of worksheetIndexes) {
  const sheetPath = `xl/worksheets/sheet${sheetIndex}.xml`;
  zip.file(sheetPath, updateWorksheetCodeName(await zip.file(sheetPath).async("string"), `Sheet${sheetIndex}`));
}

zip.file("xl/worksheets/sheet2.xml", prepareMacroSheetXml(await zip.file("xl/worksheets/sheet2.xml").async("string"), testData));
const tableEntries = Object.keys(zip.files).filter((name) => /^xl\/tables\/table\d+\.xml$/.test(name));
let holidaysTablePath = null;
for (const tablePath of tableEntries) {
  const tableXml = await zip.file(tablePath).async("string");
  if (/\bname="tblHolidays"/.test(tableXml)) {
    holidaysTablePath = tablePath;
    zip.file(
      tablePath,
      updateHolidayTableXml(tableXml, testData.quarterSettings?.expectedWeekdayHolidaysToFill),
    );
    break;
  }
}
if (!holidaysTablePath) throw new Error("tblHolidays was not found in the XLSX package");

await fs.writeFile(
  outputXlsmPath,
  await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  }),
);

normalizeWorkbookDesign(outputXlsmPath);

const verifyZip = await JSZip.loadAsync(await fs.readFile(outputXlsmPath));
if (!verifyZip.file("xl/vbaProject.bin")) throw new Error("vbaProject.bin was not embedded");
const workbookXml = await verifyZip.file("xl/workbook.xml").async("string");
if (!workbookXml.includes('codeName="ThisWorkbook"')) throw new Error("Workbook codeName was not set");
for (const sheetIndex of worksheetIndexes) {
  const sheetXml = await verifyZip.file(`xl/worksheets/sheet${sheetIndex}.xml`).async("string");
  if (!sheetXml.includes(`codeName="Sheet${sheetIndex}"`)) {
    throw new Error(`Worksheet Sheet${sheetIndex} codeName was not set`);
  }
}

console.log(`EXPORTED ${outputXlsmPath}`);
