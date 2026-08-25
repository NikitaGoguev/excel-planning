import fs from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";

const DESIGN_FONT_NAME = "Calibri";
const DESIGN_FONT_SIZE = 11;
const HIDDEN_FILTER_BUTTONS = {
  tblTeamComposition: [1, 2],
  tblTeamMembers: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12],
  tblHolidays: [1, 2, 3, 4],
  tblTaskEstimates: [5, 6, 7, 8, 9, 10, 11],
  tblPlanActive: [1, 2, 3, 4, 5, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19],
  tblPlanGrey: [1, 2, 3, 4, 5, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19],
  tblPlanBacklog: [1, 2, 3, 4, 5, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19],
};

function xmlAttr(xml, attrName) {
  const match = xml.match(new RegExp(`\\b${attrName}="([^"]*)"`));
  return match ? match[1] : null;
}

function setHiddenFilterButtonsInTableXml(xml, hiddenColumns) {
  const ref = xmlAttr(xml, "ref");
  const hiddenColIds = new Set(hiddenColumns.map((columnNumber) => Number(columnNumber) - 1));
  const filterColumnsXml = Array.from(hiddenColIds)
    .sort((a, b) => a - b)
    .map((colId) => `<filterColumn colId="${colId}" hiddenButton="1"/>`)
    .join("");

  let updatedXml = xml
    .replace(/totalsRowShown="1"/g, 'totalsRowShown="0"')
    .replace(/\s(?:totalsRowLabel|totalsRowFunction)="[^"]*"/g, "")
    .replace(/(<tableColumn\b[^>]*?)\sfilterButton="[^"]*"/g, "$1");
  updatedXml = updatedXml.replace(/(<\/tableColumn>)\sfilterButton="[^"]*"/g, "$1");

  const expandedAutoFilter = updatedXml.match(/<autoFilter\b[^>]*>[\s\S]*?<\/autoFilter>/);
  if (expandedAutoFilter) {
    const autoFilterXml = expandedAutoFilter[0];
    const openTag = autoFilterXml.match(/^<autoFilter\b[^>]*>/)?.[0] ?? `<autoFilter ref="${ref}">`;
    const keptChildren = autoFilterXml
      .slice(openTag.length, -"</autoFilter>".length)
      .replace(/<filterColumn\b[^>]*(?:\/>|>[\s\S]*?<\/filterColumn>)/g, (filterColumnXml) => {
        const colId = Number(xmlAttr(filterColumnXml, "colId"));
        return hiddenColIds.has(colId) || /\bhiddenButton="1"/.test(filterColumnXml) ? "" : filterColumnXml;
      });
    return updatedXml.replace(autoFilterXml, `${openTag}${keptChildren}${filterColumnsXml}</autoFilter>`);
  }

  const selfClosingAutoFilter = updatedXml.match(/<autoFilter\b[^>]*\/>/);
  if (selfClosingAutoFilter) {
    const autoFilterXml = selfClosingAutoFilter[0];
    const attrs = autoFilterXml
      .replace(/^<autoFilter\b/, "")
      .replace(/\/>$/, "")
      .trim();
    return updatedXml.replace(autoFilterXml, `<autoFilter${attrs ? ` ${attrs}` : ""}>${filterColumnsXml}</autoFilter>`);
  }

  if (!ref) return updatedXml;
  return updatedXml.replace(/<tableColumns\b/, `<autoFilter ref="${ref}">${filterColumnsXml}</autoFilter><tableColumns`);
}

async function normalizeWorkbookDesign(workbookPath) {
  const bytes = await fs.readFile(workbookPath);
  const zip = await JSZip.loadAsync(bytes);

  const stylesEntry = zip.file("xl/styles.xml");
  if (stylesEntry) {
    const stylesXml = await stylesEntry.async("string");
    const updatedStylesXml = stylesXml.replace(/<font>[\s\S]*?<\/font>/g, (fontXml) => {
      let updated = fontXml;
      if (/<name\b/.test(updated)) {
        updated = updated.replace(/<name\b[^>]*\/>/g, `<name val="${DESIGN_FONT_NAME}"/>`);
      } else if (/<sz\b/.test(updated)) {
        updated = updated.replace("</font>", `<name val="${DESIGN_FONT_NAME}"/></font>`);
      }

      if (/<sz\b/.test(updated)) {
        updated = updated.replace(/<sz\b[^>]*\/>/g, `<sz val="${DESIGN_FONT_SIZE}"/>`);
      }
      return updated;
    });
    zip.file("xl/styles.xml", updatedStylesXml);
  }

  for (const entryName of Object.keys(zip.files).filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/.test(name))) {
    const entry = zip.file(entryName);
    if (!entry) continue;
    const sheetXml = await entry.async("string");
    const updatedSheetXml = sheetXml.replace(/<row\b[^>]*>/g, (rowTag) =>
      rowTag
        .replace(/\sht="[^"]*"/g, "")
        .replace(/\scustomHeight="[^"]*"/g, ""),
    );
    if (updatedSheetXml !== sheetXml) zip.file(entryName, updatedSheetXml);
  }

  for (const entryName of Object.keys(zip.files).filter((name) => /^xl\/tables\/table\d+\.xml$/.test(name))) {
    const entry = zip.file(entryName);
    if (!entry) continue;
    const xml = await entry.async("string");
    const tableName = xml.match(/\bname="([^"]*)"/)?.[1];
    const hiddenColumns = HIDDEN_FILTER_BUTTONS[tableName];
    if (!hiddenColumns) continue;

    const updatedXml = setHiddenFilterButtonsInTableXml(xml, hiddenColumns);
    if (updatedXml !== xml) zip.file(entryName, updatedXml);
  }

  await fs.writeFile(
    workbookPath,
    await zip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    }),
  );
}

const workbookPaths = process.argv.slice(2);
if (workbookPaths.length === 0) {
  console.error("Usage: node scripts/normalize_workbook_design.mjs <workbook.xlsx|workbook.xlsm> [...]");
  process.exit(1);
}

for (const workbookPath of workbookPaths) {
  await normalizeWorkbookDesign(path.resolve(workbookPath));
  console.log(`NORMALIZED ${path.resolve(workbookPath)}`);
}
