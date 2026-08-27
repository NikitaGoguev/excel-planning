import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";
import {
  deriveWorkbookLayout,
  loadWorkbookLimits,
  taskEstimateCsvCaption,
} from "../src/config/workbook-limits.mjs";

const __filename = fileURLToPath(import.meta.url);
const repoRoot = path.resolve(path.dirname(__filename), "..");

const results = [];

function layoutValue(layout, key) {
  return key.split(".").reduce((value, segment) => value?.[segment], layout);
}

function resolveContractValue(value, layout) {
  if (typeof value === "string") return value;
  if (value?.layout) return layoutValue(layout, value.layout);
  if (value?.template === "taskEstimateCsvCaption") return taskEstimateCsvCaption(layout.limits);
  throw new Error(`Unsupported symbolic contract value: ${JSON.stringify(value)}`);
}

function requiredCells(contract, layout) {
  const cells = { ...contract.requiredCells };
  for (const [layoutKey, expected] of Object.entries(contract.requiredLayoutCells ?? {})) {
    cells[layoutValue(layout, layoutKey)] = expected;
  }
  return cells;
}

function pass(name) {
  results.push({ ok: true, name });
  console.log(`PASS ${name}`);
}

function fail(name, message) {
  results.push({ ok: false, name, message });
  console.error(`FAIL ${name}: ${message}`);
}

async function readJson(relativePath) {
  return JSON.parse(await fs.readFile(path.join(repoRoot, relativePath), "utf8"));
}

async function loadZip(relativePath) {
  const fullPath = path.join(repoRoot, relativePath);
  return JSZip.loadAsync(await fs.readFile(fullPath));
}

async function zipText(zip, name) {
  const file = zip.file(name);
  return file ? file.async("text") : null;
}

function xmlAttr(xml, attrName) {
  const match = xml.match(new RegExp(`\\b${attrName}="([^"]*)"`));
  return match ? match[1] : null;
}

function decodeXml(value) {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function parseSharedStrings(xml) {
  if (!xml) return [];
  return Array.from(xml.matchAll(/<si\b[\s\S]*?<\/si>/g), (si) => {
    const text = Array.from(si[0].matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g), (t) => decodeXml(t[1])).join("");
    return text;
  });
}

function cellXml(sheetXml, ref) {
  return sheetXml.match(new RegExp(`<[^:>]*(?::)?c\\b[^>]*\\br="${ref}"[^>]*(?:/>|>[\\s\\S]*?<\\/[^:>]*(?::)?c>)`))?.[0] ?? null;
}

function cellValue(sheetXml, ref, sharedStrings) {
  const xml = cellXml(sheetXml, ref);
  if (!xml) return "";
  const type = xmlAttr(xml, "t");
  if (type === "inlineStr") {
    return Array.from(xml.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g), (t) => decodeXml(t[1])).join("");
  }
  const value = xml.match(/<[^:>]*(?::)?v>([\s\S]*?)<\/[^:>]*(?::)?v>/)?.[1] ?? "";
  if (type === "s") return sharedStrings[Number(value)] ?? "";
  return decodeXml(value);
}

function colNameToNumber(name) {
  return Array.from(name).reduce((sum, ch) => sum * 26 + ch.charCodeAt(0) - 64, 0);
}

function hiddenColumns(sheetXml) {
  const hidden = new Set();
  for (const match of sheetXml.matchAll(/<col\b[^>]*>/g)) {
    const colXml = match[0];
    if (!/\bhidden="1"/.test(colXml)) continue;
    const min = Number(xmlAttr(colXml, "min"));
    const max = Number(xmlAttr(colXml, "max"));
    for (let col = min; col <= max; col += 1) hidden.add(col);
  }
  return hidden;
}

function tagAttr(tagXml, attrName) {
  const match = tagXml.match(new RegExp(`\\b${attrName}="([^"]*)"`));
  return match ? decodeXml(match[1]) : null;
}

async function tableMap(zip) {
  const tables = new Map();
  for (const name of Object.keys(zip.files).filter((file) => /^xl\/tables\/table\d+\.xml$/.test(file))) {
    const xml = await zip.file(name).async("text");
    const tableName = xmlAttr(xml, "name");
    const ref = xmlAttr(xml, "ref");
    if (tableName) tables.set(tableName, { name: tableName, ref, path: name });
  }

  for (const sheetPath of Object.keys(zip.files).filter((file) => /^xl\/worksheets\/sheet\d+\.xml$/.test(file))) {
    const relPath = sheetPath.replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels";
    const relXml = await zipText(zip, relPath);
    if (!relXml) continue;
    for (const rel of relXml.matchAll(/<Relationship\b[^>]*>/g)) {
      const relTag = rel[0];
      if (!/relationships\/table"/.test(relTag)) continue;
      const target = xmlAttr(relTag, "Target")?.replace(/^\.\.\//, "xl/");
      const table = Array.from(tables.values()).find((entry) => entry.path === target);
      if (table) table.sheetPath = sheetPath;
    }
  }
  return tables;
}

async function assertDesignContracts(buildContract) {
  const design = buildContract.design;
  if (!design) return;

  for (const relativePath of [buildContract.outputs.xlsx, buildContract.outputs.xlsm]) {
    const zip = await loadZip(relativePath);
    const stylesXml = await zipText(zip, "xl/styles.xml");
    if (!stylesXml) throw new Error(`${relativePath} misses xl/styles.xml`);

    const fonts = Array.from(stylesXml.matchAll(/<font>[\s\S]*?<\/font>/g), (match) => match[0]);
    for (const [index, fontXml] of fonts.entries()) {
      const nameTag = fontXml.match(/<name\b[^>]*\/>/)?.[0] ?? null;
      const sizeTag = fontXml.match(/<sz\b[^>]*\/>/)?.[0] ?? null;
      if (!nameTag && !sizeTag && design.allowFontRecordsWithoutNameAndSize) continue;

      const actualName = nameTag ? tagAttr(nameTag, "val") : null;
      if (actualName !== design.fontName) {
        throw new Error(`${relativePath} font ${index} name is ${JSON.stringify(actualName)}, expected ${design.fontName}`);
      }

      const actualSize = sizeTag ? Number(tagAttr(sizeTag, "val")) : NaN;
      if (actualSize !== Number(design.fontSize)) {
        throw new Error(`${relativePath} font ${index} size is ${actualSize}, expected ${design.fontSize}`);
      }
    }
  }
  pass("workbook fonts are Calibri 11");

  if (design.disallowCustomRowHeights) {
    for (const relativePath of [buildContract.outputs.xlsx, buildContract.outputs.xlsm]) {
      const zip = await loadZip(relativePath);
      for (const sheetPath of Object.keys(zip.files).filter((file) => /^xl\/worksheets\/sheet\d+\.xml$/.test(file))) {
        const sheetXml = await zipText(zip, sheetPath);
        const customRow = sheetXml.match(/<row\b[^>]*(?:\bht=|\bcustomHeight=)[^>]*>/);
        if (customRow) {
          throw new Error(`${relativePath} ${sheetPath} has custom row height: ${customRow[0]}`);
        }
      }
    }
    pass("workbook rows have no custom heights");
  }
}

function tableColumnTags(tableXml) {
  return Array.from(tableXml.matchAll(/<tableColumn\b[^>]*(?:\/>|>[\s\S]*?<\/tableColumn>)/g), (match) => match[0]);
}

function autoFilterXml(tableXml) {
  return tableXml.match(/<autoFilter\b[^>]*(?:\/>|>[\s\S]*?<\/autoFilter>)/)?.[0] ?? null;
}

function hiddenFilterColIds(tableXml) {
  const autoFilter = autoFilterXml(tableXml);
  if (!autoFilter) return new Set();

  const hidden = new Set();
  for (const match of autoFilter.matchAll(/<filterColumn\b[^>]*(?:\/>|>[\s\S]*?<\/filterColumn>)/g)) {
    const filterColumnXml = match[0];
    if (xmlAttr(filterColumnXml, "hiddenButton") === "1") {
      hidden.add(Number(xmlAttr(filterColumnXml, "colId")));
    }
  }
  return hidden;
}

async function assertFilterButtonContracts(buildContract) {
  const hiddenFilterButtons = buildContract.hiddenFilterButtons;
  if (!hiddenFilterButtons) return;

  for (const relativePath of [buildContract.outputs.xlsx, buildContract.outputs.xlsm]) {
    const zip = await loadZip(relativePath);
    const tables = await tableMap(zip);
    for (const [tableName, hiddenColumns] of Object.entries(hiddenFilterButtons)) {
      const table = tables.get(tableName);
      if (!table) throw new Error(`${relativePath} table ${tableName} not found`);
      const tableXml = await zipText(zip, table.path);
      const columns = tableColumnTags(tableXml);
      const actualHiddenColIds = hiddenFilterColIds(tableXml);
      const hiddenSet = new Set(hiddenColumns.map(Number));

      for (const [index, columnXml] of columns.entries()) {
        const columnNumber = index + 1;
        const colId = index;
        const legacyFilterButton = xmlAttr(columnXml, "filterButton");
        if (legacyFilterButton !== null) {
          throw new Error(`${relativePath} ${tableName} column ${columnNumber} uses unsupported tableColumn filterButton=${JSON.stringify(legacyFilterButton)}`);
        }
        if (hiddenSet.has(columnNumber)) {
          if (!actualHiddenColIds.has(colId)) {
            throw new Error(`${relativePath} ${tableName} column ${columnNumber} filter dropdown is visible, expected hiddenButton=1`);
          }
        } else if (actualHiddenColIds.has(colId)) {
          throw new Error(`${relativePath} ${tableName} column ${columnNumber} filter dropdown is hidden unexpectedly`);
        }
      }
    }
  }
  pass("table filter button contracts");
}

function stableJson(value) {
  if (Array.isArray(value)) return value.map(stableJson);
  if (value && typeof value === "object") {
    return Object.fromEntries(Object.keys(value).sort().map((key) => [key, stableJson(value[key])]));
  }
  return value;
}

async function normalizedSnapshot(xlsmZip, buildContract, sheet03Contract, sheet04Contract, vbaContract, layout) {
  const tables = await tableMap(xlsmZip);
  const sharedStrings = parseSharedStrings(await zipText(xlsmZip, "xl/sharedStrings.xml"));
  const tableRefs = {};
  for (const [name, table] of tables.entries()) {
    tableRefs[name] = table.ref;
  }

  const cells = {};
  const taskTable = tables.get("tblTaskEstimates");
  if (taskTable?.sheetPath) {
    const taskSheetXml = await zipText(xlsmZip, taskTable.sheetPath);
    for (const cellRef of Object.keys(requiredCells(sheet03Contract, layout))) {
      cells[`tblTaskEstimates:${cellRef}`] = cellValue(taskSheetXml, cellRef, sharedStrings);
    }
  }
  const backlogTable = tables.get("tblPlanBacklog");
  if (backlogTable?.sheetPath) {
    const planSheetXml = await zipText(xlsmZip, backlogTable.sheetPath);
    for (const cellRef of Object.keys(requiredCells(sheet04Contract, layout))) {
      cells[`tblPlanBacklog:${cellRef}`] = cellValue(planSheetXml, cellRef, sharedStrings);
    }
  }

  return stableJson({
    workbook: {
      workbookCodeName: buildContract.package.workbookCodeName,
      worksheetCodeNames: buildContract.package.worksheetCodeNames,
    },
    tables: tableRefs,
    actionCells: cells,
    vbaPublicApi: [...vbaContract.requiredPublicMacros].sort(),
  });
}

async function assertSnapshot(workbookContract, snapshot) {
  const snapshotPath = workbookContract.snapshots?.structure;
  if (!snapshotPath) return;
  const fullPath = path.join(repoRoot, snapshotPath);
  if (process.env.UPDATE_WORKBOOK_SNAPSHOT === "1") {
    await fs.mkdir(path.dirname(fullPath), { recursive: true });
    await fs.writeFile(fullPath, `${JSON.stringify(snapshot, null, 2)}\n`, "utf8");
    pass("normalized workbook snapshot updated");
    return;
  }
  const expected = JSON.parse(await fs.readFile(fullPath, "utf8"));
  const expectedText = `${JSON.stringify(stableJson(expected), null, 2)}\n`;
  const actualText = `${JSON.stringify(snapshot, null, 2)}\n`;
  if (actualText !== expectedText) {
    throw new Error(`normalized workbook snapshot drifted: ${snapshotPath}`);
  }
  pass("normalized workbook snapshot");
}

async function assertPackageContracts(buildContract) {
  const xlsxZip = await loadZip(buildContract.outputs.xlsx);
  const xlsmZip = await loadZip(buildContract.outputs.xlsm);

  for (const part of buildContract.package.xlsxMustNotContain) {
    if (xlsxZip.file(part)) throw new Error(`${buildContract.outputs.xlsx} contains forbidden part ${part}`);
  }
  pass("xlsx has no vbaProject.bin");

  for (const part of buildContract.package.xlsmMustContain) {
    if (!xlsmZip.file(part)) throw new Error(`${buildContract.outputs.xlsm} misses required part ${part}`);
  }
  pass("xlsm has vbaProject.bin");

  const contentTypes = await zipText(xlsmZip, "[Content_Types].xml");
  if (!contentTypes?.includes(buildContract.package.xlsmContentType)) {
    throw new Error("macro-enabled workbook content type not found");
  }
  pass("xlsm content type is macro-enabled");

  const workbookXml = await zipText(xlsmZip, "xl/workbook.xml");
  if (!workbookXml?.includes(`codeName="${buildContract.package.workbookCodeName}"`)) {
    throw new Error(`workbook codeName ${buildContract.package.workbookCodeName} not found`);
  }
  pass("workbook codeName is present");

  for (const [index, codeName] of buildContract.package.worksheetCodeNames.entries()) {
    const sheetXml = await zipText(xlsmZip, `xl/worksheets/sheet${index + 1}.xml`);
    if (!sheetXml?.includes(`codeName="${codeName}"`)) {
      throw new Error(`worksheet sheet${index + 1} codeName ${codeName} not found`);
    }
  }
  pass("worksheet codeNames are present");

  return { xlsxZip, xlsmZip };
}

async function assertSheetContracts(xlsmZip, sheet03Contract, sheet04Contract, sheet101Contract, layout) {
  const tables = await tableMap(xlsmZip);
  for (const [name, expectedValue] of Object.entries(sheet03Contract.tables)) {
    const expectedRef = resolveContractValue(expectedValue, layout);
    const table = tables.get(name);
    if (!table) throw new Error(`table ${name} not found`);
    if (table.ref !== expectedRef) throw new Error(`table ${name} ref ${table.ref}, expected ${expectedRef}`);
  }
  pass("sheet 03 table contracts");

  const sharedStrings = parseSharedStrings(await zipText(xlsmZip, "xl/sharedStrings.xml"));
  const taskTable = tables.get("tblTaskEstimates");
  const taskSheetXml = await zipText(xlsmZip, taskTable.sheetPath);
  for (const [cellRef, expectedValue] of Object.entries(requiredCells(sheet03Contract, layout))) {
    const expected = resolveContractValue(expectedValue, layout);
    const actual = cellValue(taskSheetXml, cellRef, sharedStrings);
    if (actual !== expected) throw new Error(`${cellRef} is ${JSON.stringify(actual)}, expected ${JSON.stringify(expected)}`);
  }
  pass("sheet 03 action cells");

  for (const [name, expectedValue] of Object.entries(sheet04Contract.tables)) {
    const expectedRef = resolveContractValue(expectedValue, layout);
    const table = tables.get(name);
    if (!table) throw new Error(`table ${name} not found`);
    if (table.ref !== expectedRef) throw new Error(`table ${name} ref ${table.ref}, expected ${expectedRef}`);
  }
  pass("sheet 04 table contracts");

  const backlogTable = tables.get("tblPlanBacklog");
  const planSheetXml = await zipText(xlsmZip, backlogTable.sheetPath);
  for (const [cellRef, expectedValue] of Object.entries(requiredCells(sheet04Contract, layout))) {
    const expected = resolveContractValue(expectedValue, layout);
    const actual = cellValue(planSheetXml, cellRef, sharedStrings);
    if (actual !== expected) throw new Error(`${cellRef} is ${JSON.stringify(actual)}, expected ${JSON.stringify(expected)}`);
  }
  pass("sheet 04 action cells");

  const hidden = hiddenColumns(planSheetXml);
  const hiddenRequired = sheet04Contract.visibleColumns.filter((col) => hidden.has(colNameToNumber(col)));
  if (hiddenRequired.length > 0) throw new Error(`sheet 04 columns hidden: ${hiddenRequired.join(", ")}`);
  pass("sheet 04 technical columns are visible");

  if (!planSheetXml.includes("<conditionalFormatting")) {
    throw new Error("sheet 04 conditional formatting is missing");
  }
  pass("sheet 04 status conditional formatting is present");

  for (const [name, expectedValue] of Object.entries(sheet101Contract.tables)) {
    const table = tables.get(name);
    if (!table) throw new Error(`table ${name} not found`);
    if (table.ref !== expectedValue) throw new Error(`table ${name} ref ${table.ref}, expected ${expectedValue}`);
  }
  const listTable = tables.get("tblTaskCommentArtifacts");
  const listSheetXml = await zipText(xlsmZip, listTable.sheetPath);
  for (const [cellRef, expectedValue] of Object.entries(sheet101Contract.requiredCells)) {
    const actual = cellValue(listSheetXml, cellRef, sharedStrings);
    if (actual !== expectedValue) throw new Error(`${cellRef} is ${JSON.stringify(actual)}, expected ${JSON.stringify(expectedValue)}`);
  }
  pass("sheet 101 list contracts");
}

async function executableFiles(relativePath) {
  const fullPath = path.join(repoRoot, relativePath);
  const stat = await fs.stat(fullPath);
  if (stat.isFile()) return [relativePath];
  const entries = await fs.readdir(fullPath, { withFileTypes: true });
  const nested = await Promise.all(entries
    .filter((entry) => entry.isDirectory() || /\.(mjs|ps1|json)$/.test(entry.name))
    .map((entry) => executableFiles(path.join(relativePath, entry.name))));
  return nested.flat();
}

async function assertPortableSources() {
  const files = (await Promise.all([
    executableFiles("scripts"),
    executableFiles("package.json"),
    executableFiles("Makefile"),
  ])).flat();
  const forbidden = [
    { name: "private artifact package", value: "@oai/" + "artifact-tool" },
    { name: "personal Codex runtime", value: ["codex", "-runtimes"].join("") },
    { name: "internal pnpm store", value: [".", "pnpm"].join("") },
    { name: "personal Windows profile", value: ["C:", "\\Users\\", "nikita", "goguev"].join("") },
  ];
  for (const relativePath of files) {
    const source = await fs.readFile(path.join(repoRoot, relativePath), "utf8");
    for (const pattern of forbidden) {
      if (source.toLowerCase().includes(pattern.value.toLowerCase())) {
        throw new Error(`${pattern.name} found in ${relativePath}`);
      }
    }
  }
  pass("portable source dependency contract");
}

async function assertVbaContracts(vbaContract) {
  const components = vbaContract.components ?? vbaContract.sources.map((source) => ({ source }));
  const sourceText = (
    await Promise.all(components.map(({ source }) => fs.readFile(path.join(repoRoot, source), "utf8")))
  ).join("\n");

  const componentNames = components.map(({ name }) => name).filter(Boolean);
  const duplicates = componentNames.filter((name, index) => componentNames.indexOf(name) !== index);
  if (duplicates.length) throw new Error(`duplicate VBA components: ${[...new Set(duplicates)].join(", ")}`);
  pass("VBA component manifest");

  for (const component of components.filter(({ type }) => type === "form")) {
    if (!component.source.endsWith(".frm")) throw new Error(`VBA form source must be .frm: ${component.source}`);
    const companionPath = path.join(repoRoot, component.source.replace(/\.frm$/i, ".frx"));
    await fs.access(companionPath);
  }
  pass("VBA form source companions");

  for (const macro of vbaContract.requiredPublicMacros) {
    const pattern = new RegExp(`\\bPublic\\s+Sub\\s+${macro}\\b`, "i");
    if (!pattern.test(sourceText)) throw new Error(`required public macro not found: ${macro}`);
  }
  pass("VBA public macros are present");

  for (const forbidden of vbaContract.forbiddenPatterns) {
    const pattern = new RegExp(forbidden.pattern, "i");
    if (pattern.test(sourceText)) throw new Error(`forbidden VBA pattern found: ${forbidden.name}`);
  }
  pass("VBA forbidden patterns are absent");

  if (vbaContract.scheduler) {
    const schedulerComponent = components.find(({ name }) => name === vbaContract.scheduler.component);
    if (!schedulerComponent) throw new Error(`scheduler component is missing: ${vbaContract.scheduler.component}`);
    const schedulerSource = await fs.readFile(path.join(repoRoot, schedulerComponent.source), "utf8");
    if (!new RegExp(`\\bPublic\\s+Function\\s+${vbaContract.scheduler.entrypoint}\\b`, "i").test(schedulerSource)) {
      throw new Error(`scheduler entrypoint is missing: ${vbaContract.scheduler.entrypoint}`);
    }
    for (const sourcePattern of vbaContract.scheduler.forbiddenPatterns) {
      if (new RegExp(sourcePattern, "i").test(schedulerSource)) throw new Error(`scheduler is not domain-pure: ${sourcePattern}`);
    }
    const testSource = await fs.readFile(path.join(repoRoot, vbaContract.scheduler.testSource), "utf8");
    const scenarioMatch = testSource.match(/Private\s+Const\s+QPS_TEST_SCENARIO_COUNT\s+As\s+Long\s*=\s*(\d+)/i);
    if (!scenarioMatch || Number(scenarioMatch[1]) !== vbaContract.scheduler.scenarioCount) {
      throw new Error(`scheduler scenario count must equal ${vbaContract.scheduler.scenarioCount}`);
    }
    pass("VBA scheduler domain boundary and scenario manifest");
  }

  const handwrittenText = (
    await Promise.all(components.filter(({ generated }) => !generated).map(({ source }) => fs.readFile(path.join(repoRoot, source), "utf8")))
  ).join("\n");
  for (const sourcePattern of vbaContract.handwrittenLimitForbiddenPatterns ?? []) {
    if (new RegExp(sourcePattern, "i").test(handwrittenText)) throw new Error(`handwritten VBA duplicates a generated limit: ${sourcePattern}`);
  }
  pass("VBA structural limits use generated constants");
}

async function assertWorkbookContract(workbookContract, xlsmZip) {
  const workbookXml = await zipText(xlsmZip, "xl/workbook.xml");
  const sheets = workbookXml.match(/<sheet\b/g) ?? [];
  if (sheets.length !== workbookContract.workbook.sheetCount) {
    throw new Error(`workbook sheet count ${sheets.length}, expected ${workbookContract.workbook.sheetCount}`);
  }
  pass("workbook sheet count");

  const sheetNames = Array.from(workbookXml.matchAll(/<sheet\b[^>]*\bname="([^"]*)"/g), (match) => decodeXml(match[1]));
  const expectedNames = workbookContract.workbook.sheetNames ?? [];
  if (JSON.stringify(sheetNames) !== JSON.stringify(expectedNames)) {
    throw new Error(`workbook sheet order ${JSON.stringify(sheetNames)}, expected ${JSON.stringify(expectedNames)}`);
  }
  pass("workbook sheet order");
}

try {
  const layout = deriveWorkbookLayout(await loadWorkbookLimits());
  const buildContract = await readJson("contracts/build.contract.json");
  if (process.env.QUARTER_PLANNING_XLSX_OUTPUT) buildContract.outputs.xlsx = process.env.QUARTER_PLANNING_XLSX_OUTPUT;
  if (process.env.QUARTER_PLANNING_XLSM_OUTPUT) buildContract.outputs.xlsm = process.env.QUARTER_PLANNING_XLSM_OUTPUT;
  const sheet03Contract = await readJson("contracts/sheet03.contract.json");
  const sheet04Contract = await readJson("contracts/sheet04.contract.json");
  const sheet101Contract = await readJson("contracts/sheet101.contract.json");
  const vbaContract = await readJson("contracts/vba.contract.json");
  const workbookContract = await readJson("contracts/workbook.contract.json");

  const { xlsmZip } = await assertPackageContracts(buildContract);
  await assertPortableSources();
  await assertDesignContracts(buildContract);
  await assertFilterButtonContracts(buildContract);
  await assertWorkbookContract(workbookContract, xlsmZip);
  await assertSheetContracts(xlsmZip, sheet03Contract, sheet04Contract, sheet101Contract, layout);
  await assertVbaContracts(vbaContract);
  await assertSnapshot(
    workbookContract,
    await normalizedSnapshot(xlsmZip, buildContract, sheet03Contract, sheet04Contract, vbaContract, layout),
  );
} catch (error) {
  fail("static contracts", error.message);
}

if (results.some((result) => !result.ok)) {
  process.exit(1);
}
