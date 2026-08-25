import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";

import { WorkbookAdapter, appendWorksheetFromFile } from "./lib/exceljs_compat.mjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");
const outputDir = path.join(repoRoot, "outputs");
const outputPath = process.env.QUARTER_PLANNING_XLSX_OUTPUT
  ? path.resolve(process.env.QUARTER_PLANNING_XLSX_OUTPUT)
  : path.join(outputDir, "quarter_planning_step1.xlsx");
const testDataPath = process.env.QUARTER_PLANNING_DATA_PATH
  ? path.resolve(process.env.QUARTER_PLANNING_DATA_PATH)
  : path.join(repoRoot, "data", "test_data_quarter_planning.json");
const expressEstimateTemplatePath = path.join(repoRoot, "assets", "Шаблон Экспресс оценки.xlsx");

let testData = {};
try {
  testData = JSON.parse(await fs.readFile(testDataPath, "utf8"));
} catch {
  testData = {};
}

const SHEET_SETTINGS = "00_\u041d\u0430\u0441\u0442\u0440\u043e\u0439\u043a\u0438";
const SHEET_QUARTER = "01_\u041d\u0430\u0441\u0442\u0440\u043e\u0439\u043a\u0438 \u043a\u0432\u0430\u0440\u0442\u0430\u043b\u0430";
const SHEET_CAPACITY = "02_Capacity";
const SHEET_ESTIMATES = "03_\u041e\u0446\u0435\u043d\u043a\u0430 \u0437\u0430\u0434\u0430\u0447";
const SHEET_PLAN = "04_\u041a\u0432\u0430\u0440\u0442\u0430\u043b\u044c\u043d\u044b\u0439 \u043f\u043b\u0430\u043d";
const SHEET_EXPRESS_TEMPLATE = "100_\u0428\u0430\u0431\u043b\u043e\u043d \u044d\u043a\u0441\u043f\u0440\u0435\u0441\u0441 \u043e\u0446\u0435\u043d\u043a\u0438";
const SHEET_REFS = "99_\u0421\u043f\u0440\u0430\u0432\u043e\u0447\u043d\u0438\u043a\u0438";
const EXCEL_DATE_FORMAT = "dd-mm-yyyy";
const DESIGN_FONT_NAME = "Calibri";
const DESIGN_FONT_SIZE = 11;
const TEAM_MEMBER_LIMIT = 20;

function excelSheetName(name) {
  return Array.from(name).slice(0, 31).join("");
}

function sheetRef(name) {
  return `'${excelSheetName(name).replace(/'/g, "''")}'`;
}

function dateFromIso(value) {
  return value ? new Date(`${value}T00:00:00Z`) : "";
}

function teamCompositionCount(index) {
  const item = testData.teamComposition?.[index];
  return item?.count ?? item?.people ?? item?.fte ?? "";
}

function teamMemberRows() {
  const rows = [];
  for (const item of testData.teamComposition ?? []) {
    const role = item.role ?? "";
    const count = Math.max(0, Math.min(TEAM_MEMBER_LIMIT, Number(item.count ?? item.people ?? item.fte ?? 0) || 0));
    for (let roleIndex = 1; roleIndex <= count && rows.length < TEAM_MEMBER_LIMIT; roleIndex += 1) {
      rows.push([role, `${role} ${roleIndex}`, 1, "", "", "", "", "", "", "", "", ""]);
    }
  }
  while (rows.length < TEAM_MEMBER_LIMIT) rows.push(["", "", "", "", "", "", "", "", "", "", "", ""]);
  return rows;
}

function teamMemberVacationDayFormulas() {
  return Array.from({ length: TEAM_MEMBER_LIMIT }, (_, index) => {
    const row = 24 + index;
    return [
      teamMemberVacationDaysFormula(`D${row}`, `E${row}`),
      teamMemberVacationDaysFormula(`G${row}`, `H${row}`),
      teamMemberVacationDaysFormula(`J${row}`, `K${row}`),
    ];
  });
}

function teamMemberVacationDaysFormula(startCell, endCell) {
  return `=IF(OR(${startCell}="",${endCell}=""),"",IFERROR(MAX(0,NETWORKDAYS(${startCell},${endCell})-SUMPRODUCT((tblHolidays[Учитывать]="Да")*(tblHolidays[Дата]>=${startCell})*(tblHolidays[Дата]<=${endCell})*(WEEKDAY(tblHolidays[Дата],2)<=5))),""))`;
}

function teamMemberAllocationSumFormula(roleName) {
  return `=SUMIFS(tblTeamMembers[Аллокация],tblTeamMembers[Роль],"${roleName}")`;
}

function teamMemberVacationSumFormula(roleName) {
  const criteria = `(tblTeamMembers[Роль]="${roleName}")*tblTeamMembers[Аллокация]`;
  return `=IFERROR(SUMPRODUCT(${criteria},tblTeamMembers[Отпуск 1 дней])+SUMPRODUCT(${criteria},tblTeamMembers[Отпуск 2 дней])+SUMPRODUCT(${criteria},tblTeamMembers[Отпуск 3 дней]),0)`;
}

const colors = {
  title: "#1F4E79",
  section: "#9DC3E6",
  header: "#D9EAF7",
  input: "#FFF2CC",
  calculated: "#E2F0D9",
  technical: "#F3F3F3",
  border: "#808080",
  text: "#1F1F1F",
  mutedText: "#808080",
  white: "#FFFFFF",
};

function fontStyle(overrides = {}) {
  return { name: DESIGN_FONT_NAME, size: DESIGN_FONT_SIZE, ...overrides };
}

function applyTitle(range) {
  range.format = {
    fill: colors.title,
    font: fontStyle({ bold: true, color: colors.white }),
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };
}

function applyHeader(range) {
  range.format = {
    fill: colors.header,
    font: fontStyle({ bold: true, color: colors.text }),
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: {
      top: { style: "continuous", color: colors.border },
      bottom: { style: "continuous", color: colors.border },
      left: { style: "continuous", color: colors.border },
      right: { style: "continuous", color: colors.border },
    },
  };
}

function applyInput(range) {
  range.format = {
    fill: colors.input,
    font: fontStyle({ color: colors.text }),
    verticalAlignment: "center",
    wrapText: true,
    borders: {
      top: { style: "continuous", color: colors.border },
      bottom: { style: "continuous", color: colors.border },
      left: { style: "continuous", color: colors.border },
      right: { style: "continuous", color: colors.border },
    },
  };
}

function applyCalculated(range) {
  range.format = {
    fill: colors.calculated,
    font: fontStyle({ color: colors.text }),
    verticalAlignment: "center",
    wrapText: true,
    borders: {
      top: { style: "continuous", color: colors.border },
      bottom: { style: "continuous", color: colors.border },
      left: { style: "continuous", color: colors.border },
      right: { style: "continuous", color: colors.border },
    },
  };
}

function applyPlain(range) {
  range.format = {
    font: fontStyle({ color: colors.text }),
    verticalAlignment: "top",
    wrapText: true,
    borders: {
      top: { style: "continuous", color: colors.border },
      bottom: { style: "continuous", color: colors.border },
      left: { style: "continuous", color: colors.border },
      right: { style: "continuous", color: colors.border },
    },
  };
}

function applyAction(range) {
  range.format = {
    fill: colors.header,
    font: fontStyle({ bold: true, color: colors.title }),
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: {
      top: { style: "continuous", color: colors.border },
      bottom: { style: "continuous", color: colors.border },
      left: { style: "continuous", color: colors.border },
      right: { style: "continuous", color: colors.border },
    },
  };
}

function setWidths(sheet, widths) {
  for (const [col, widthPx] of Object.entries(widths)) {
    sheet.getRange(`${col}:${col}`).format.columnWidthPx = widthPx;
  }
}

function applyCapacityRangeStyle(range, fill, horizontalAlignment = "left", bold = false) {
  range.format = {
    fill,
    font: fontStyle({ bold, color: colors.text }),
    horizontalAlignment,
    verticalAlignment: "center",
    wrapText: true,
    borders: {
      top: { style: "continuous", color: "#000000" },
      bottom: { style: "continuous", color: "#000000" },
      left: { style: "continuous", color: "#000000" },
      right: { style: "continuous", color: "#000000" },
    },
  };
}

function addTable(sheet, range, name) {
  const table = sheet.tables.add(range, true, name);
  table.style = "TableStyleMedium2";
  table.showFilterButton = true;
  return table;
}

function addListValidation(range, values) {
  range.dataValidation = {
    rule: { type: "list", values },
    prompt: { showPrompt: true },
    errorAlert: { showAlert: true },
  };
}

function addWholeValidation(range, min, max) {
  range.dataValidation = {
    rule: {
      type: "whole",
      operator: "between",
      formula1: min,
      formula2: max,
    },
    errorAlert: { showAlert: true },
  };
}

function addDecimalValidation(range, min, max) {
  range.dataValidation = {
    rule: {
      type: "decimal",
      operator: "between",
      formula1: min,
      formula2: max,
    },
    errorAlert: { showAlert: true },
  };
}

function addNonNegativeValidation(range) {
  range.dataValidation = {
    rule: {
      type: "decimal",
      operator: "between",
      formula1: 0,
      formula2: 999999,
    },
    ignoreBlank: true,
    errorAlert: { showAlert: true },
  };
}

async function allowBlankInDataValidations(xlsxPath) {
  const bytes = await fs.readFile(xlsxPath);
  const zip = await JSZip.loadAsync(bytes);

  for (const entryName of Object.keys(zip.files).filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/.test(name))) {
    const entry = zip.file(entryName);
    if (!entry) continue;

    const xml = await entry.async("string");
    if (!xml.includes("dataValidation")) continue;

    const updatedXml = xml.replace(/<([A-Za-z0-9_]+:)?dataValidation\b(?![^>]*\ballowBlank=)/g, (_match, prefix = "") => {
      return `<${prefix}dataValidation allowBlank="1"`;
    });
    if (updatedXml !== xml) {
      zip.file(entryName, updatedXml);
    }
  }

  await fs.writeFile(
    xlsxPath,
    await zip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    }),
  );
}

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

async function applyHiddenFilterButtons(xlsxPath) {
  const bytes = await fs.readFile(xlsxPath);
  const zip = await JSZip.loadAsync(bytes);

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
    xlsxPath,
    await zip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    }),
  );
}

async function normalizeWorkbookDesign(xlsxPath) {
  const bytes = await fs.readFile(xlsxPath);
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

  await fs.writeFile(
    xlsxPath,
    await zip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    }),
  );
}

const workbook = WorkbookAdapter.create();

const settings = workbook.worksheets.add(SHEET_SETTINGS);
const quarter = workbook.worksheets.add(SHEET_QUARTER);
const capacity = workbook.worksheets.add(SHEET_CAPACITY);
const estimates = workbook.worksheets.add(SHEET_ESTIMATES);
const plan = workbook.worksheets.add(SHEET_PLAN);
const refs = workbook.worksheets.add(SHEET_REFS);

for (const sheet of [settings, quarter, refs, capacity, estimates, plan]) {
  sheet.showGridLines = false;
}

settings.getRange("A1:C1").merge();
settings.getRange("A1").values = [["Настройки квартального планирования"]];
applyTitle(settings.getRange("A1:C1"));
settings.getRange("A3:C8").values = [
  ["Параметр", "Значение", "Комментарий"],
  ["Команда", testData.settings?.team ?? "", "Название команды или потока"],
  ["Руководитель / ИТ-лид", testData.settings?.projectLead ?? "", "Ответственный за планирование"],
  ["Год", testData.settings?.year ?? "", "Календарный год квартала"],
  ["Квартал", testData.settings?.quarter ?? "", "Номер квартала: 1, 2, 3 или 4"],
  ["Часов в рабочем дне", testData.settings?.hoursPerWorkingDay ?? 8, "Используется на шаге 2 для перевода ч/д в ч/ч"],
];
for (let row = 3; row <= 8; row += 1) {
  settings.getRange(`C${row}:G${row}`).merge();
}
applyHeader(settings.getRange("A3:G3"));
applyPlain(settings.getRange("A4:A8"));
applyInput(settings.getRange("B4:B7"));
applyInput(settings.getRange("B8"));
applyPlain(settings.getRange("C4:G8"));
settings.getRange("B6:B8").setNumberFormat("0");
addWholeValidation(settings.getRange("B6"), 2020, 2100);
addWholeValidation(settings.getRange("B7"), 1, 4);
addDecimalValidation(settings.getRange("B8"), 1, 24);

settings.getRange("A11:M11").merge();
settings.getRange("A11").values = [["Настройка команды"]];
settings.getRange("A11:M11").format = {
  fill: colors.section,
  font: fontStyle({ bold: true, color: colors.text }),
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
settings.getRange("A13:B19").values = [
  ["Роль", "Количество людей"],
  ["Тим-лид / техлид", teamCompositionCount(0)],
  ["Лид по аналитике", teamCompositionCount(1)],
  ["Аналитик", teamCompositionCount(2)],
  ["Разработчик бэкенд", teamCompositionCount(3)],
  ["Разработчик фронтенд", teamCompositionCount(4)],
  ["Тестировщик", teamCompositionCount(5)],
];
applyHeader(settings.getRange("A13:B13"));
applyPlain(settings.getRange("A14:A19"));
applyInput(settings.getRange("B14:B19"));
settings.getRange("B14:B19").setNumberFormat("0");
addWholeValidation(settings.getRange("B14:B19"), 0, TEAM_MEMBER_LIMIT);
addTable(settings, "A13:B19", "tblTeamComposition");
settings.getRange("C13:G19").values = [
  ["Комментарий", null, null, null, null],
  [testData.teamComposition?.[0]?.comment ?? "", null, null, null, null],
  [testData.teamComposition?.[1]?.comment ?? "", null, null, null, null],
  [testData.teamComposition?.[2]?.comment ?? "", null, null, null, null],
  [testData.teamComposition?.[3]?.comment ?? "", null, null, null, null],
  [testData.teamComposition?.[4]?.comment ?? "", null, null, null, null],
  [testData.teamComposition?.[5]?.comment ?? "", null, null, null, null],
];
for (let row = 13; row <= 19; row += 1) {
  settings.getRange(`C${row}:G${row}`).merge();
}
applyHeader(settings.getRange("C13:G13"));
applyInput(settings.getRange("C14:G19"));
settings.getRange("H13:M19").format = {
  font: fontStyle({ color: colors.text }),
  verticalAlignment: "center",
  wrapText: true,
};
settings.getRange("H13").formulas = [[`=IF(SUM(B14:B19)>${TEAM_MEMBER_LIMIT},"Превышен лимит ${TEAM_MEMBER_LIMIT} сотрудников","")`]];
settings.getRange("H13:M13").merge();
settings.getRange("H13:M13").format = {
  fill: colors.input,
  font: fontStyle({ bold: true, color: "#9C0006" }),
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
  borders: {
    top: { style: "continuous", color: colors.border },
    bottom: { style: "continuous", color: colors.border },
    left: { style: "continuous", color: colors.border },
    right: { style: "continuous", color: colors.border },
  },
};
settings.getRange("A22:M22").merge();
settings.getRange("A22").values = [["Сотрудники команды"]];
settings.getRange("A22:M22").format = {
  fill: colors.section,
  font: fontStyle({ bold: true, color: colors.text }),
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
settings.getRange("A23:L43").values = [
  [
    "Роль",
    "Сотрудник",
    "Аллокация",
    "Отпуск 1 с",
    "Отпуск 1 по",
    "Отпуск 1 дней",
    "Отпуск 2 с",
    "Отпуск 2 по",
    "Отпуск 2 дней",
    "Отпуск 3 с",
    "Отпуск 3 по",
    "Отпуск 3 дней",
  ],
  ...teamMemberRows(),
];
settings.getRange("F24:F43").formulas = teamMemberVacationDayFormulas().map((row) => [row[0]]);
settings.getRange("I24:I43").formulas = teamMemberVacationDayFormulas().map((row) => [row[1]]);
settings.getRange("L24:L43").formulas = teamMemberVacationDayFormulas().map((row) => [row[2]]);
applyHeader(settings.getRange("A23:L23"));
applyCalculated(settings.getRange("A24:A43"));
applyInput(settings.getRange("B24:L43"));
applyCalculated(settings.getRange("F24:F43"));
applyCalculated(settings.getRange("I24:I43"));
applyCalculated(settings.getRange("L24:L43"));
settings.getRange("C24:C43").setNumberFormat("0%");
settings.getRange("D24:E43").setNumberFormat(EXCEL_DATE_FORMAT);
settings.getRange("G24:H43").setNumberFormat(EXCEL_DATE_FORMAT);
settings.getRange("J24:K43").setNumberFormat(EXCEL_DATE_FORMAT);
settings.getRange("F24:F43").setNumberFormat("0");
settings.getRange("I24:I43").setNumberFormat("0");
settings.getRange("L24:L43").setNumberFormat("0");
addDecimalValidation(settings.getRange("C24:C43"), 0, 1);
addTable(settings, "A23:L43", "tblTeamMembers");
settings.freezePanes.freezeRows(3);
setWidths(settings, {
  A: 230,
  B: 196,
  C: 130,
  D: 105,
  E: 105,
  F: 105,
  G: 105,
  H: 105,
  I: 105,
  J: 105,
  K: 105,
  L: 105,
});

quarter.getRange("A1:D1").merge();
quarter.getRange("A1").values = [["Настройки квартала"]];
applyTitle(quarter.getRange("A1:D1"));
quarter.getRange("A3:C8").values = [
  ["Параметр", "Значение", "Комментарий"],
  ["Дата начала", dateFromIso(testData.quarterSettings?.startDate), "Первый календарный день планируемого квартала"],
  ["Дата завершения", dateFromIso(testData.quarterSettings?.endDate), "Последний календарный день планируемого квартала"],
  ["Количество рабочих дней", testData.quarterSettings?.workingDays ?? "", "Рабочие дни по производственному календарю; заполняется вручную"],
  ["Будние дни по календарю", "", "Автоматически: количество дней Пн-Пт между датой начала и датой завершения"],
  ["Праздничных дней к заполнению", "", "Автоматически: будние дни по календарю минус введенное количество рабочих дней"],
];
applyHeader(quarter.getRange("A3:E3"));
applyPlain(quarter.getRange("A4:A8"));
applyInput(quarter.getRange("B4:B6"));
applyCalculated(quarter.getRange("B7:B8"));
applyPlain(quarter.getRange("C4:C8"));
quarter.getRange("B4:B5").setNumberFormat(EXCEL_DATE_FORMAT);
quarter.getRange("B6:B8").setNumberFormat("0");
addWholeValidation(quarter.getRange("B6"), 0, 100);
quarter.getRange("B7").formulas = [[`=IF(OR(B4="",B5=""),"",NETWORKDAYS(B4,B5))`]];
quarter.getRange("B8").formulas = [[`=IF(OR(B4="",B5="",B6=""),"",MAX(0,B7-B6))`]];

quarter.getRange("A11:D11").merge();
quarter.getRange("A11").values = [["Список праздничных дней"]];
quarter.getRange("A11:D11").format = {
  fill: colors.section,
  font: fontStyle({ bold: true, color: colors.text }),
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
quarter.getRange("A13:D33").values = [
  ["Дата", "Название", "Комментарий", "Учитывать"],
  ...Array.from({ length: 20 }, () => ["", "", "", ""]),
];
applyHeader(quarter.getRange("A13:D13"));
applyCalculated(quarter.getRange("A14:A33"));
applyInput(quarter.getRange("B14:C33"));
applyCalculated(quarter.getRange("D14:D33"));
quarter.getRange("A14:A33").setNumberFormat(EXCEL_DATE_FORMAT);
quarter.getRange("A14").formulas = [[`=IF(OR($B$4="",ROW(A1)>$B$8),"",WORKDAY($B$4-1,ROW(A1)))`]];
quarter.getRange("A14:A33").fillDown();
quarter.getRange("D14").formulas = [[`=IF($A14="","","Да")`]];
quarter.getRange("D14:D33").fillDown();
addListValidation(quarter.getRange("D14:D33"), ["Да", "Нет"]);
addTable(quarter, "A13:D33", "tblHolidays");
quarter.freezePanes.freezeRows(13);
setWidths(quarter, { A: 190, B: 240, C: 460, D: 120 });

refs.getRange("A1:D1").merge();
refs.getRange("A1").values = [["Справочники"]];
applyTitle(refs.getRange("A1:D1"));
refs.getRange("A3:C7").values = [
  ["Экспертиза", "Название", "Waterfall порядок"],
  ["AN", "Аналитика", 1],
  ["BE", "Backend-разработка", 2],
  ["FE", "Frontend-разработка", 3],
  ["QA", "Тестирование", 4],
];
applyHeader(refs.getRange("A3:E3"));
applyPlain(refs.getRange("A4:C7"));
addTable(refs, "A3:C7", "tblExpertise");
const quarterPlanStatuses = [
  "Готова аналитика",
  "Готова разработка",
  "Готова разработка (бэк)",
  "Готова разработка (фронт)",
  "Готово к релизу",
  "ПРОМ",
  "Отложено",
];

refs.getRange("E3:E10").values = [
  ["Статусы"],
  ...quarterPlanStatuses.map((status) => [status]),
];
applyHeader(refs.getRange("E3:E3"));
applyPlain(refs.getRange("E4:E10"));
addTable(refs, "E3:E10", "tblStatuses");
refs.getRange("G3:G4").values = [["Да/Нет"], ["Да"]];
refs.getRange("G5").values = [["Нет"]];
applyHeader(refs.getRange("G3:G3"));
applyPlain(refs.getRange("G4:G5"));
addTable(refs, "G3:G5", "tblYesNo");
refs.getRange("A11:G11").merge();
refs.getRange("A11").values = [["Правила квартального планирования"]];
refs.getRange("A11:G11").format = {
  fill: colors.section,
  font: fontStyle({ bold: true, color: colors.text }),
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
refs.getRange("A13:G27").values = [
  ["Правило", "Описание", null, null, null, null, null],
  ["Единица оценки", "Оценки задач указаны в человеко-днях.", null, null, null, null, null],
  ["Источник задач", "Первичный источник задач: лист 03_Оценка задач.", null, null, null, null, null],
  ["Стартовое состояние", "При первичной сборке все задачи попадают в бэклог.", null, null, null, null, null],
  ["Переходы", "Бэклог -> План; План -> Серая зона; План -> Бэклог; Серая зона -> Бэклог.", null, null, null, null, null],
  ["Формула длительности", "Расчетная длительность этапа = CEILING(оценка / фокус-фактор основной команды, 1).", null, null, null, null, null],
  ["Пустые оценки", "Пустые AN/BE/FE/QA считаются нулем.", null, null, null, null, null],
  ["Ресурсы AN", "AN выполняют аналитики и лид по аналитике.", null, null, null, null, null],
  ["Ресурсы BE", "BE выполняют BE-разработчики и тим-лид / техлид.", null, null, null, null, null],
  ["Ресурсы FE/QA", "FE выполняют FE-разработчики; QA выполняют тестировщики.", null, null, null, null, null],
  ["Порядок расчета", "Задачи в квартальном плане рассчитываются сверху вниз по строкам секции.", null, null, null, null, null],
  ["Даты готовности", "Даты считаются только для секции Квартальный план; для серой зоны и бэклога очищаются.", null, null, null, null, null],
  ["AN/BE overlap", "Если AN больше настроенного лага, BE может стартовать после заданного числа рабочих дней AN и после выполнения заданной доли AN; если AN завершилась раньше, BE стартует после AN.", null, null, null, null, null],
  ["FE overlap", "Если BE больше 5 ч/д, FE может стартовать после 5 рабочих дней BE и после выполнения 50% BE; если BE завершился раньше, FE стартует после BE.", null, null, null, null, null],
  ["QA overlap", "\u041f\u043e\u043b\u043e\u0432\u0438\u043d\u0430 QA \u043c\u043e\u0436\u0435\u0442 \u0441\u0442\u0430\u0440\u0442\u043e\u0432\u0430\u0442\u044c \u043f\u043e\u0441\u043b\u0435 \u0437\u0430\u0432\u0435\u0440\u0448\u0435\u043d\u0438\u044f \u0430\u043d\u0430\u043b\u0438\u0442\u0438\u043a\u0438; \u0432\u0442\u043e\u0440\u0430\u044f \u043f\u043e\u043b\u043e\u0432\u0438\u043d\u0430 QA \u0441\u0442\u0430\u0440\u0442\u0443\u0435\u0442 \u043f\u043e\u0441\u043b\u0435 \u0437\u0430\u0432\u0435\u0440\u0448\u0435\u043d\u0438\u044f BE \u0438 FE.", null, null, null, null, null],
];
refs.getRange("B13:G27").merge(true);
applyHeader(refs.getRange("A13:G13"));
applyPlain(refs.getRange("A14:G27"));
refs.getRange("A29:E35").values = [
  ["Код", "Правило", "Значение", "Ед.", "Описание"],
  ["AN_BE_MIN_WORKDAYS", "AN -> BE: минимальный лаг", 5, "раб. дни", "Минимум рабочих дней от старта AN до возможного старта BE для длинной аналитики."],
  ["AN_BE_COMPLETION_PERCENT", "AN -> BE: готовность AN", 0.5, "%", "Доля расчетной длительности AN, после которой BE может стартовать для длинной аналитики."],
  ["BE_FE_MIN_WORKDAYS", "BE -> FE: минимальный лаг", 5, "раб. дни", "Минимум рабочих дней от старта BE до возможного старта FE для длинной backend-разработки."],
  ["BE_FE_COMPLETION_PERCENT", "BE -> FE: готовность BE", 0.5, "%", "Доля расчетной длительности BE, после которой FE может стартовать для длинной backend-разработки."],
  ["QA_FIRST_PART_PERCENT", "QA: первая часть", 0.5, "%", "Доля QA, которая может выполняться после завершения аналитики."],
  ["RESOURCE_BALANCE_RISK_PERCENT", "Баланс ресурсов: риск", 0.2, "%", "Риск, добавляемый к оценкам активного плана при расчете баланса ресурсов."],
];
applyHeader(refs.getRange("A29:E29"));
applyPlain(refs.getRange("A30:E35"));
addTable(refs, "A29:E35", "tblPlanningRuleSettings");
addWholeValidation(refs.getRange("C30:C30"), 0, 1000);
addWholeValidation(refs.getRange("C32:C32"), 0, 1000);
addDecimalValidation(refs.getRange("C31:C31"), 0, 1);
addDecimalValidation(refs.getRange("C33:C35"), 0, 1);
refs.getRange("C30:C30").setNumberFormat("0");
refs.getRange("C32:C32").setNumberFormat("0");
refs.getRange("C31:C31").setNumberFormat("0%");
refs.getRange("C33:C35").setNumberFormat("0%");
refs.freezePanes.freezeRows(3);
setWidths(refs, { A: 190, B: 270, C: 95, D: 85, E: 350, F: 40, G: 110 });

capacity.getRange("C1:J1").merge();
capacity.getRange("C1").values = [["Capacity команды"]];
applyTitle(capacity.getRange("C1:J1"));
capacity.getRange("C1:J1").format = {
  fill: colors.title,
  font: fontStyle({ bold: true, color: colors.white }),
  horizontalAlignment: "center",
  verticalAlignment: "center",
  borders: {
    top: { style: "continuous", color: "#000000" },
    bottom: { style: "continuous", color: "#000000" },
    left: { style: "continuous", color: "#000000" },
    right: { style: "continuous", color: "#000000" },
  },
};

capacity.getRange("C2:D4").values = [
  ["Команда", null],
  ["Квартал", null],
  ["ИТ-лид команды", null],
];
capacity.getRange("C2:D4").merge(true);
capacity.getRange("E2:J4").values = [
  ["", null, null, null, null, null],
  ["", null, null, null, null, null],
  ["", null, null, null, null, null],
];
capacity.getRange("E2:J4").merge(true);
capacity.getRange("E2").formulas = [[`=${sheetRef(SHEET_SETTINGS)}!B4`]];
capacity.getRange("E3").formulas = [[`=${sheetRef(SHEET_SETTINGS)}!B7&" квартал "&${sheetRef(SHEET_SETTINGS)}!B6`]];
capacity.getRange("E4").formulas = [[`=${sheetRef(SHEET_SETTINGS)}!B5`]];
capacity.getRange("C2:D4").format = {
  fill: colors.section,
  font: fontStyle({ color: colors.text }),
  verticalAlignment: "center",
  borders: {
    top: { style: "continuous", color: "#000000" },
    bottom: { style: "continuous", color: "#000000" },
    left: { style: "continuous", color: "#000000" },
    right: { style: "continuous", color: "#000000" },
  },
};
capacity.getRange("E2:J4").format = {
  fill: colors.white,
  font: fontStyle({ color: colors.text }),
  horizontalAlignment: "center",
  verticalAlignment: "center",
  borders: {
    top: { style: "continuous", color: "#000000" },
    bottom: { style: "continuous", color: "#000000" },
    left: { style: "continuous", color: "#000000" },
    right: { style: "continuous", color: "#000000" },
  },
};

capacity.getRange("C6:G22").values = [
  ["Показатель", null, "Значение", "Авто 00", "Переопределение"],
  ["Количество рабочих дней в квартале", null, "", "", ""],
  ["Количество аналитиков в команде", null, "", "", ""],
  ["Количество разработчиков BE в команде", null, "", "", ""],
  ["Количество разработчиков FE в команде", null, "", "", ""],
  ["Количество тестировщиков в команде", null, "", "", ""],
  ["Тим лид команды (разработка BE/FE, тестирование)", null, "", "", ""],
  ["Лид по аналитике", null, "", "", ""],
  ["Отпуска аналитиков", null, "", "", ""],
  ["Отпуска разработчиков BE", null, "", "", ""],
  ["Отпуска разработчиков FE", null, "", "", ""],
  ["Отпуска тестировщиков", null, "", "", ""],
  ["Отпуска тим лида команды (разработка BE, тестирование)", null, "", "", ""],
  ["Отпуска лида по аналитике", null, "", "", ""],
  ["Фокус-фактор основной команды", null, "", "", ""],
  ["Фокус-фактор тим лида", null, "", "", ""],
  ["Фокус-фактор лида по аналитике", null, "", "", ""],
];
capacity.getRange("C6:D22").merge(true);
applyHeader(capacity.getRange("C6:G6"));
applyPlain(capacity.getRange("C7:D19"));
capacity.getRange("E7:G19").format = {
  fill: colors.white,
  font: fontStyle({ color: colors.text }),
  horizontalAlignment: "right",
  verticalAlignment: "center",
  borders: {
    top: { style: "continuous", color: "#000000" },
    bottom: { style: "continuous", color: "#000000" },
    left: { style: "continuous", color: "#000000" },
    right: { style: "continuous", color: "#000000" },
  },
};
capacity.getRange("C20:E20").format = {
  fill: "#92D050",
  font: fontStyle({ color: colors.text }),
  verticalAlignment: "center",
  wrapText: true,
  borders: {
    top: { style: "continuous", color: "#000000" },
    bottom: { style: "continuous", color: "#000000" },
    left: { style: "continuous", color: "#000000" },
    right: { style: "continuous", color: "#000000" },
  },
};
applyPlain(capacity.getRange("C21:D22"));
capacity.getRange("E21:E22").format = {
  fill: colors.white,
  font: fontStyle({ color: colors.text }),
  horizontalAlignment: "right",
  verticalAlignment: "center",
  borders: {
    top: { style: "continuous", color: "#000000" },
    bottom: { style: "continuous", color: "#000000" },
    left: { style: "continuous", color: "#000000" },
    right: { style: "continuous", color: "#000000" },
  },
};
capacity.getRange("E7:G19").setNumberFormat("0.00");
capacity.getRange("E20:E22").setNumberFormat("0%");
capacity.getRange("E7:E19").formulas = [
  [`=${sheetRef(SHEET_QUARTER)}!B6`],
  [`=IF(G8="",F8,G8)`],
  [`=IF(G9="",F9,G9)`],
  [`=IF(G10="",F10,G10)`],
  [`=IF(G11="",F11,G11)`],
  [`=IF(G12="",F12,G12)`],
  [`=IF(G13="",F13,G13)`],
  [`=IF(G14="",F14,G14)`],
  [`=IF(G15="",F15,G15)`],
  [`=IF(G16="",F16,G16)`],
  [`=IF(G17="",F17,G17)`],
  [`=IF(G18="",F18,G18)`],
  [`=IF(G19="",F19,G19)`],
];
capacity.getRange("F8:F13").formulas = [
  [teamMemberAllocationSumFormula("Аналитик")],
  [teamMemberAllocationSumFormula("Разработчик бэкенд")],
  [teamMemberAllocationSumFormula("Разработчик фронтенд")],
  [teamMemberAllocationSumFormula("Тестировщик")],
  [teamMemberAllocationSumFormula("Тим-лид / техлид")],
  [teamMemberAllocationSumFormula("Лид по аналитике")],
];
capacity.getRange("F14:F19").formulas = [
  [teamMemberVacationSumFormula("Аналитик")],
  [teamMemberVacationSumFormula("Разработчик бэкенд")],
  [teamMemberVacationSumFormula("Разработчик фронтенд")],
  [teamMemberVacationSumFormula("Тестировщик")],
  [teamMemberVacationSumFormula("Тим-лид / техлид")],
  [teamMemberVacationSumFormula("Лид по аналитике")],
];
capacity.getRange("F6:G19").format = {
  fill: colors.technical,
  font: fontStyle({ color: colors.mutedText }),
  horizontalAlignment: "right",
  verticalAlignment: "center",
  wrapText: true,
  borders: {
    top: { style: "continuous", color: "#D9D9D9" },
    bottom: { style: "continuous", color: "#D9D9D9" },
    left: { style: "continuous", color: "#D9D9D9" },
    right: { style: "continuous", color: "#D9D9D9" },
  },
};
capacity.getRange("F6:G6").format = {
  fill: colors.technical,
  font: fontStyle({ bold: true, color: colors.mutedText }),
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
  borders: {
    top: { style: "continuous", color: "#D9D9D9" },
    bottom: { style: "continuous", color: "#D9D9D9" },
    left: { style: "continuous", color: "#D9D9D9" },
    right: { style: "continuous", color: "#D9D9D9" },
  },
};
applyInput(capacity.getRange("G8:G19"));
capacity.getRange("G8:G19").format = {
  fill: colors.technical,
  font: fontStyle({ color: colors.mutedText }),
  horizontalAlignment: "right",
  verticalAlignment: "center",
  borders: {
    top: { style: "continuous", color: "#D9D9D9" },
    bottom: { style: "continuous", color: "#D9D9D9" },
    left: { style: "continuous", color: "#D9D9D9" },
    right: { style: "continuous", color: "#D9D9D9" },
  },
};
capacity.getRange("E20:E22").values = [
  [testData.capacity?.coreTeamFocusFactor ?? ""],
  [testData.capacity?.teamLeadFocusFactor ?? ""],
  [testData.capacity?.analysisLeadFocusFactor ?? ""],
];
addNonNegativeValidation(capacity.getRange("E7:E19"));
addNonNegativeValidation(capacity.getRange("G8:G19"));
addDecimalValidation(capacity.getRange("E20:E22"), 0, 1);

capacity.getRange("F20:M22").values = [
  ["- заполняется вручную для каждой команды, рекомендуемое значение - 70%", null, null, null, null, null, null, null],
  ["- заполняется вручную для каждой команды", null, null, null, null, null, null, null],
  ["- заполняется вручную для каждой команды (ограничение - не менее 50%)", null, null, null, null, null, null, null],
];
capacity.getRange("F20:M22").merge(true);
capacity.getRange("F20:M22").format = {
  font: fontStyle({ color: colors.text }),
  verticalAlignment: "center",
  wrapText: true,
};

capacity.getRange("E26:F26").values = [["ч/д", "ч/ч"]];
applyHeader(capacity.getRange("E26:F26"));
capacity.getRange("C26:D26").format = {
  fill: colors.white,
  borders: {
    top: { style: "continuous", color: "#000000" },
    bottom: { style: "continuous", color: "#000000" },
    left: { style: "continuous", color: "#000000" },
    right: { style: "continuous", color: "#000000" },
  },
};
capacity.getRange("C27:F30").values = [
  ["Capacity по аналитикам", null, "", ""],
  ["Capacity по разработчикам BE", null, "", ""],
  ["Capacity по разработчикам FE", null, "", ""],
  ["Capacity по тестировщикам", null, "", ""],
];
capacity.getRange("C27:D30").merge(true);
capacity.getRange("C27:D30").format = {
  fill: colors.section,
  font: fontStyle({ color: colors.text }),
  verticalAlignment: "center",
  wrapText: true,
  borders: {
    top: { style: "continuous", color: "#000000" },
    bottom: { style: "continuous", color: "#000000" },
    left: { style: "continuous", color: "#000000" },
    right: { style: "continuous", color: "#000000" },
  },
};
applyCalculated(capacity.getRange("E27:F30"));
capacity.getRange("C32:F32").values = [["Общее Capacity", null, "", ""]];
capacity.getRange("C32:D32").merge();
capacity.getRange("C32:D32").format = {
  fill: colors.section,
  font: fontStyle({ color: colors.text }),
  verticalAlignment: "center",
  wrapText: true,
  borders: {
    top: { style: "continuous", color: "#000000" },
    bottom: { style: "continuous", color: "#000000" },
    left: { style: "continuous", color: "#000000" },
    right: { style: "continuous", color: "#000000" },
  },
};
applyCalculated(capacity.getRange("E32:F32"));
capacity.getRange("E27:F32").setNumberFormat("0.00");
capacity.getRange("E27:F30").formulas = [
  [
    `=IFERROR(ROUNDDOWN(((E7*E8-E14)*E20)+((E7*E13-E19)*E22),0),0)`,
    `=IFERROR(E27*${sheetRef(SHEET_SETTINGS)}!B$8,0)`,
  ],
  [
    `=IFERROR(ROUNDDOWN(((E7*E9-E15)*E20)+((E7*E12-E18)*E21),0),0)`,
    `=IFERROR(E28*${sheetRef(SHEET_SETTINGS)}!B$8,0)`,
  ],
  [
    `=IFERROR(ROUNDDOWN((E7*E10-E16)*E20,0),0)`,
    `=IFERROR(E29*${sheetRef(SHEET_SETTINGS)}!B$8,0)`,
  ],
  [
    `=IFERROR(ROUNDDOWN((E7*E11-E17)*E20,0),0)`,
    `=IFERROR(E30*${sheetRef(SHEET_SETTINGS)}!B$8,0)`,
  ],
];
capacity.getRange("E32:F32").formulas = [[`=SUM(E27:E30)`, `=SUM(F27:F30)`]];
for (let row = 2; row <= 4; row += 1) {
  applyCapacityRangeStyle(capacity.getRange(`C${row}:D${row}`), colors.section);
  applyCapacityRangeStyle(capacity.getRange(`E${row}:J${row}`), colors.white, "center");
}
for (let row = 7; row <= 19; row += 1) {
  applyCapacityRangeStyle(capacity.getRange(`C${row}:D${row}`), colors.white);
  applyCapacityRangeStyle(capacity.getRange(`E${row}:E${row}`), colors.white, "right");
}
applyCapacityRangeStyle(capacity.getRange("C20:E20"), "#92D050");
for (let row = 21; row <= 22; row += 1) {
  applyCapacityRangeStyle(capacity.getRange(`C${row}:D${row}`), colors.white);
  applyCapacityRangeStyle(capacity.getRange(`E${row}:E${row}`), colors.white);
}
applyCapacityRangeStyle(capacity.getRange("C26:D26"), colors.white);
applyCapacityRangeStyle(capacity.getRange("E26:F26"), colors.header, "center", true);
for (let row = 27; row <= 30; row += 1) {
  applyCapacityRangeStyle(capacity.getRange(`C${row}:D${row}`), colors.section);
  applyCapacityRangeStyle(capacity.getRange(`E${row}:F${row}`), colors.calculated, "right");
}
applyCapacityRangeStyle(capacity.getRange("C32:D32"), colors.section);
applyCapacityRangeStyle(capacity.getRange("E32:F32"), colors.calculated, "right");
capacity.freezePanes.freezeRows(6);
setWidths(capacity, {
  A: 40,
  B: 50,
  C: 250,
  D: 340,
  E: 130,
  F: 130,
  G: 120,
  H: 120,
  I: 120,
  J: 120,
  K: 120,
  L: 120,
  M: 120,
});

estimates.getRange("E1").values = [["Оценка задач"]];
applyTitle(estimates.getRange("E1"));
estimates.getRange("A2:B2").merge();
estimates.getRange("A2").values = [["Сбросить"]];
applyAction(estimates.getRange("A2:B2"));
estimates.getRange("C2").values = [["Обновить"]];
applyAction(estimates.getRange("C2"));
estimates.getRange("C1").values = [["Экспорт"]];
applyAction(estimates.getRange("C1"));
estimates.getRange("D1").values = [["Импорт"]];
applyAction(estimates.getRange("D1"));
estimates.getRange("D2").values = [["Импорт CSV (до 100)"]];
applyAction(estimates.getRange("D2"));
estimates.getRange("H1:K1").merge();
estimates.getRange("H1").values = [["Статистика"]];
applyTitle(estimates.getRange("H1:K1"));
estimates.getRange("H2:K2").formulas = [[
  "SUM(H4:H103)",
  "SUM(I4:I103)",
  "SUM(J4:J103)",
  "SUM(K4:K103)",
]];
applyCalculated(estimates.getRange("H2:K2"));
estimates.getRange("H2:K2").setNumberFormat("0.00");
estimates.getRange("A3:M103").values = [
  [">", "x", "Приоритет", "Направление", "Описание", "ЗНИ/Jira", "Примечание", "AN", "BE", "FE", "QA", "Комментарий", "Экспорт"],
  ...Array.from({ length: 100 }, (_, index) => {
    const task = testData.taskEstimates?.[index];
    return task
      ? [
          ">",
          "",
          task.priority ?? "",
          task.direction ?? "",
          task.description ?? "",
          task.ticket ?? "",
          task.note ?? "",
          task.AN ?? "",
          task.BE ?? "",
          task.FE ?? "",
          task.QA ?? "",
          task.comment ?? "",
          "",
        ]
      : ["", "", "", "", "", "", "", "", "", "", "", "", ""];
  }),
];
applyHeader(estimates.getRange("A3:M3"));
applyAction(estimates.getRange("A4:B103"));
applyInput(estimates.getRange("C4:L103"));
applyAction(estimates.getRange("M4:M103"));
estimates.getRange("C4:C103").setNumberFormat("0");
estimates.getRange("H4:K103").setNumberFormat("0.00");
addWholeValidation(estimates.getRange("C4:C103"), 1, 999);
addNonNegativeValidation(estimates.getRange("H4:K103"));
addTable(estimates, "C3:M103", "tblTaskEstimates");
estimates.freezePanes.freezeRows(3);
setWidths(estimates, {
  A: 46,
  B: 46,
  C: 101,
  D: 172,
  E: 479,
  F: 172,
  G: 220,
  H: 56,
  I: 56,
  J: 56,
  K: 56,
  L: 519,
  M: 140,
});

const planHeaders = [
  "+",
  "\\",
  "-",
  "\u2191",
  "\u2193",
  "Приоритет",
  "Направление",
  "Описание",
  "ЗНИ/Jira",
  "Примечание",
  "AN",
  "BE",
  "FE",
  "QA",
  "Трудозатраты",
  "Завершение (план)",
  "Релиз",
  "Статус на конец квартала",
  "Комментарий",
];

const planTechHeaders = [
  "AN дни",
  "BE дни",
  "FE дни",
  "QA дни",
  "AN старт",
  "AN финиш",
  "BE старт",
  "BE финиш",
  "FE старт",
  "FE финиш",
  "QA старт",
  "QA финиш",
  "Диагностика",
];

function plannedEffortFormula(row) {
  const factor = "'02_Capacity'!$E$20";
  return `=IF(COUNTA(G${row}:N${row})=0,"",IFERROR(IF(K${row}="",0,CEILING(K${row}/${factor},1))+IF(L${row}="",0,CEILING(L${row}/${factor},1))+IF(M${row}="",0,CEILING(M${row}/${factor},1))+IF(N${row}="",0,CEILING(N${row}/${factor},1)),""))`;
}

function emptyPlanRow() {
  return Array.from({ length: planHeaders.length }, () => "");
}

function taskToPlanRow(task, action) {
  return [
    action === "toPlan" ? "+" : "",
    action === "toGrey" ? "\\" : "",
    action === "toBacklog" ? "-" : "",
    "",
    "",
    task.priority ?? "",
    task.direction ?? "",
    task.description ?? "",
    task.ticket ?? "",
    task.note ?? "",
    task.AN ?? "",
    task.BE ?? "",
    task.FE ?? "",
    task.QA ?? "",
    "",
    "",
    "",
    "",
    task.comment ?? "",
  ];
}

function makePlanRows(count, tasks = [], action = "") {
  return Array.from({ length: count }, (_, index) => {
    const task = tasks[index];
    return task ? taskToPlanRow(task, action) : emptyPlanRow();
  });
}

function applyPlanSection(sectionTitle, titleRange, tableRange, headerRange, dataStartRow, dataEndRow, tableName, rows) {
  plan.getRange(titleRange).merge();
  plan.getRange(titleRange.split(":")[0]).values = [[sectionTitle]];
  plan.getRange(titleRange).format = {
    fill: sectionTitle === "Серая зона" ? "#A6A6A6" : colors.title,
    font: fontStyle({ bold: true, color: colors.white }),
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };
  plan.getRange(tableRange).values = [planHeaders, ...rows];
  applyHeader(plan.getRange(headerRange));
  applyAction(plan.getRange(`A${dataStartRow}:E${dataEndRow}`));
  applyInput(plan.getRange(`F${dataStartRow}:J${dataEndRow}`));
  applyInput(plan.getRange(`K${dataStartRow}:N${dataEndRow}`));
  applyCalculated(plan.getRange(`O${dataStartRow}:P${dataEndRow}`));
  applyInput(plan.getRange(`Q${dataStartRow}:S${dataEndRow}`));
  plan.getRange(`A${dataStartRow}:AF${dataEndRow}`).format.wrapText = false;
  plan.getRange(`F${dataStartRow}:F${dataEndRow}`).setNumberFormat("0");
  plan.getRange(`K${dataStartRow}:O${dataEndRow}`).setNumberFormat("0.00");
  plan.getRange(`P${dataStartRow}:P${dataEndRow}`).setNumberFormat(EXCEL_DATE_FORMAT);
  plan.getRange(`O${dataStartRow}:O${dataEndRow}`).formulas = Array.from(
    { length: dataEndRow - dataStartRow + 1 },
    (_, index) => {
      const row = dataStartRow + index;
      return [plannedEffortFormula(row)];
    },
  );
  plan.getRange(`T${dataStartRow - 1}:AF${dataStartRow - 1}`).values = [planTechHeaders];
  applyHeader(plan.getRange(`T${dataStartRow - 1}:AF${dataStartRow - 1}`));
  applyCalculated(plan.getRange(`T${dataStartRow}:AF${dataEndRow}`));
  plan.getRange(`T${dataStartRow}:W${dataEndRow}`).setNumberFormat("0.00");
  plan.getRange(`X${dataStartRow}:AE${dataEndRow}`).setNumberFormat(EXCEL_DATE_FORMAT);
  addWholeValidation(plan.getRange(`F${dataStartRow}:F${dataEndRow}`), 1, 999);
  addNonNegativeValidation(plan.getRange(`K${dataStartRow}:N${dataEndRow}`));
  addListValidation(plan.getRange(`R${dataStartRow}:R${dataEndRow}`), quarterPlanStatuses);
  addTable(plan, tableRange, tableName);
}

applyPlanSection("Квартальный план", "A1:S1", "A5:S25", "A5:S5", 6, 25, "tblPlanActive", makePlanRows(20));
applyPlanSection("Серая зона", "A28:S28", "A30:S50", "A30:S30", 31, 50, "tblPlanGrey", makePlanRows(20));
applyPlanSection(
  "Бэклог",
  "A53:S53",
  "A55:S155",
  "A55:S55",
  56,
  155,
  "tblPlanBacklog",
  makePlanRows(100, testData.taskEstimates ?? [], "toPlan"),
);

function addExcludedEstimateFormat(reference, formula) {
  plan.raw.addConditionalFormatting({
    ref: reference,
    rules: [{
      type: "expression",
      formulae: [formula],
      style: {
        fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } },
        font: { color: { argb: "FF808080" }, strike: true },
      },
    }],
  });
}

for (const [startRow, endRow] of [[6, 25], [31, 50], [56, 155]]) {
  addExcludedEstimateFormat(`L${startRow}:N${endRow}`, `$R${startRow}="Готова аналитика"`);
  addExcludedEstimateFormat(`N${startRow}:N${endRow}`, `$R${startRow}="Готова разработка"`);
  addExcludedEstimateFormat(`M${startRow}:N${endRow}`, `$R${startRow}="Готова разработка (бэк)"`);
  addExcludedEstimateFormat(`L${startRow}:L${endRow}`, `$R${startRow}="Готова разработка (фронт)"`);
  addExcludedEstimateFormat(`N${startRow}:N${endRow}`, `$R${startRow}="Готова разработка (фронт)"`);
  addExcludedEstimateFormat(`K${startRow}:N${endRow}`, `$R${startRow}="Отложено"`);
}

plan.getRange("A5:E5").format.horizontalAlignment = "left";
plan.getRange("A30:E30").format.horizontalAlignment = "left";
plan.getRange("A55:E55").format.horizontalAlignment = "left";
plan.freezePanes.freezeRows(5);
setWidths(plan, {
  A: 35,
  B: 35,
  C: 35,
  D: 35,
  E: 35,
  F: 86,
  G: 150,
  H: 360,
  I: 166,
  J: 193,
  K: 52,
  L: 52,
  M: 52,
  N: 52,
  O: 120,
  P: 113,
  Q: 130,
  R: 230,
  S: 420,
});
setWidths(plan, {
  T: 70,
  U: 70,
  V: 70,
  W: 70,
  X: 100,
  Y: 100,
  Z: 100,
});
setWidths(plan, {
  AA: 100,
  AB: 100,
  AC: 100,
  AD: 100,
  AE: 100,
  AF: 260,
});
plan.getRange("T:AF").format.columnHidden = false;

const refsSheetRef = sheetRef(SHEET_REFS);
const resourceRiskFormula =
  `IFERROR(IF(AND(ISNUMBER(${refsSheetRef}!$C$35),${refsSheetRef}!$C$35>=0,${refsSheetRef}!$C$35<=1),${refsSheetRef}!$C$35,0.2),0.2)`;

function plannedWithRiskExpression(estimateExpression) {
  const riskFactor = `(1+${resourceRiskFormula})`;
  return `(${estimateExpression})*${riskFactor}`;
}

function resourceCapacityFormula(capacityRef) {
  return `=ROUND(${capacityRef},1)`;
}

function resourcePlannedWithRiskFormula(estimateExpression) {
  return `=ROUND(${plannedWithRiskExpression(estimateExpression)},1)`;
}

function resourceBalanceFormula(capacityRef, estimateExpression) {
  return `=ROUND(${capacityRef}-${plannedWithRiskExpression(estimateExpression)},1)`;
}

function activePlanEstimateExpression(expertise) {
  const statusColumn = "tblPlanActive[Статус на конец квартала]";
  if (expertise === "AN") return `SUMIFS(tblPlanActive[AN],${statusColumn},"<>Отложено")`;
  if (expertise === "BE") {
    return `SUMIFS(tblPlanActive[BE],${statusColumn},"<>Готова аналитика",${statusColumn},"<>Готова разработка (фронт)",${statusColumn},"<>Отложено")`;
  }
  if (expertise === "FE") {
    return `SUMIFS(tblPlanActive[FE],${statusColumn},"<>Готова аналитика",${statusColumn},"<>Готова разработка (бэк)",${statusColumn},"<>Отложено")`;
  }
  if (expertise === "QA") {
    return `SUMIFS(tblPlanActive[QA],${statusColumn},"<>Готова аналитика",${statusColumn},"<>Готова разработка",${statusColumn},"<>Готова разработка (бэк)",${statusColumn},"<>Готова разработка (фронт)",${statusColumn},"<>Отложено")`;
  }
  throw new Error(`Unsupported expertise for balance: ${expertise}`);
}

plan.getRange("J2:N4").format = {
  fill: colors.calculated,
  font: fontStyle({ bold: true, color: colors.text }),
  horizontalAlignment: "center",
  verticalAlignment: "center",
  borders: {
    top: { style: "continuous", color: colors.border },
    bottom: { style: "continuous", color: colors.border },
    left: { style: "continuous", color: colors.border },
    right: { style: "continuous", color: colors.border },
  },
};
plan.getRange("J2:J4").values = [["Емкость"], ["План+риск"], ["Баланс"]];
for (const [column, capacityRef, expertise] of [
  ["K", "'02_Capacity'!$E$27", "AN"],
  ["L", "'02_Capacity'!$E$28", "BE"],
  ["M", "'02_Capacity'!$E$29", "FE"],
  ["N", "'02_Capacity'!$E$30", "QA"],
]) {
  const estimateExpression = activePlanEstimateExpression(expertise);
  plan.getRange(`${column}2`).formulas = [[resourceCapacityFormula(capacityRef)]];
  plan.getRange(`${column}3`).formulas = [[resourcePlannedWithRiskFormula(estimateExpression)]];
  plan.getRange(`${column}4`).formulas = [[resourceBalanceFormula(capacityRef, estimateExpression)]];
}
plan.getRange("K2:N4").setNumberFormat("0.0");
plan.getRange("A54").values = [["++"]];
plan.getRange("A54").format = {
  fill: colors.title,
  font: fontStyle({ bold: true, color: colors.white }),
  horizontalAlignment: "center",
  verticalAlignment: "center",
};

plan.getRange("G2:G4").merge();
plan.getRange("G2").values = [["\u0417\u0430\u0433\u0440\u0443\u0437\u0438\u0442\u044c 03 \u0432 \u0431\u044d\u043a\u043b\u043e\u0433"]];
plan.getRange("G2:G4").format = {
  fill: colors.title,
  font: fontStyle({ bold: true, color: colors.white }),
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
plan.getRange("P2:P4").merge();
plan.getRange("P2").values = [["\u041f\u043e\u0441\u0447\u0438\u0442\u0430\u0442\u044c \u043f\u043b\u0430\u043d"]];
plan.getRange("P2:P4").format = {
  fill: colors.title,
  font: fontStyle({ bold: true, color: colors.white }),
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
plan.getRange("Q2:Q4").merge();
plan.getRange("Q2").values = [["Экспорт плана"]];
plan.getRange("Q2:Q4").format = {
  fill: colors.title,
  font: fontStyle({ bold: true, color: colors.white }),
  horizontalAlignment: "center",
  verticalAlignment: "center",
};

await fs.mkdir(path.dirname(outputPath), { recursive: true });
await appendWorksheetFromFile(workbook.raw, expressEstimateTemplatePath, SHEET_EXPRESS_TEMPLATE);
await workbook.raw.xlsx.writeFile(outputPath);
await allowBlankInDataValidations(outputPath);
await applyHiddenFilterButtons(outputPath);
await normalizeWorkbookDesign(outputPath);
console.log(`EXPORTED ${outputPath}`);
