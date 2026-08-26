import fs from "node:fs/promises";

import JSZip from "jszip";

import { DESIGN_FONT_NAME, DESIGN_FONT_SIZE } from "../workbook/design.mjs";

export const HIDDEN_FILTER_BUTTONS = Object.freeze({
  tblTeamComposition: [1, 2],
  tblTeamMembers: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12],
  tblHolidays: [1, 2, 3, 4],
  tblTaskEstimates: [5, 6, 7, 8, 9, 10, 11],
  tblPlanActive: [1, 2, 3, 4, 5, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19],
  tblPlanGrey: [1, 2, 3, 4, 5, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19],
  tblPlanBacklog: [1, 2, 3, 4, 5, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19],
});

async function rewritePackage(xlsxPath, update) {
  const zip = await JSZip.loadAsync(await fs.readFile(xlsxPath));
  await update(zip);
  await fs.writeFile(xlsxPath, await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  }));
}

export async function allowBlankInDataValidations(xlsxPath) {
  await rewritePackage(xlsxPath, async (zip) => {
    for (const entryName of Object.keys(zip.files).filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/.test(name))) {
      const entry = zip.file(entryName);
      if (!entry) continue;
      const xml = await entry.async("string");
      if (!xml.includes("dataValidation")) continue;
      const updatedXml = xml.replace(/<([A-Za-z0-9_]+:)?dataValidation\b(?![^>]*\ballowBlank=)/g, (_match, prefix = "") => `<${prefix}dataValidation allowBlank="1"`);
      if (updatedXml !== xml) zip.file(entryName, updatedXml);
    }
  });
}

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
    const attrs = autoFilterXml.replace(/^<autoFilter\b/, "").replace(/\/>$/, "").trim();
    return updatedXml.replace(autoFilterXml, `<autoFilter${attrs ? ` ${attrs}` : ""}>${filterColumnsXml}</autoFilter>`);
  }

  if (!ref) return updatedXml;
  return updatedXml.replace(/<tableColumns\b/, `<autoFilter ref="${ref}">${filterColumnsXml}</autoFilter><tableColumns`);
}

export async function applyHiddenFilterButtons(xlsxPath) {
  await rewritePackage(xlsxPath, async (zip) => {
    for (const entryName of Object.keys(zip.files).filter((name) => /^xl\/tables\/table\d+\.xml$/.test(name))) {
      const entry = zip.file(entryName);
      if (!entry) continue;
      const xml = await entry.async("string");
      const hiddenColumns = HIDDEN_FILTER_BUTTONS[xml.match(/\bname="([^"]*)"/)?.[1]];
      if (!hiddenColumns) continue;
      const updatedXml = setHiddenFilterButtonsInTableXml(xml, hiddenColumns);
      if (updatedXml !== xml) zip.file(entryName, updatedXml);
    }
  });
}

export async function normalizeWorkbookDesign(xlsxPath) {
  await rewritePackage(xlsxPath, async (zip) => {
    const stylesEntry = zip.file("xl/styles.xml");
    if (stylesEntry) {
      const stylesXml = await stylesEntry.async("string");
      zip.file("xl/styles.xml", stylesXml.replace(/<font>[\s\S]*?<\/font>/g, (fontXml) => {
        let updated = fontXml;
        if (/<name\b/.test(updated)) updated = updated.replace(/<name\b[^>]*\/>/g, `<name val="${DESIGN_FONT_NAME}"/>`);
        else if (/<sz\b/.test(updated)) updated = updated.replace("</font>", `<name val="${DESIGN_FONT_NAME}"/></font>`);
        if (/<sz\b/.test(updated)) updated = updated.replace(/<sz\b[^>]*\/>/g, `<sz val="${DESIGN_FONT_SIZE}"/>`);
        return updated;
      }));
    }

    for (const entryName of Object.keys(zip.files).filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/.test(name))) {
      const entry = zip.file(entryName);
      if (!entry) continue;
      const sheetXml = await entry.async("string");
      const updatedSheetXml = sheetXml.replace(/<row\b[^>]*>/g, (rowTag) => rowTag.replace(/\sht="[^"]*"/g, "").replace(/\scustomHeight="[^"]*"/g, ""));
      if (updatedSheetXml !== sheetXml) zip.file(entryName, updatedSheetXml);
    }
  });
}
