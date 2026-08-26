import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { WorkbookAdapter, appendWorksheetFromFile } from "../../scripts/lib/exceljs_compat.mjs";
import { deriveWorkbookLayout, loadWorkbookLimits, taskEstimateCsvCaption } from "../config/workbook-limits.mjs";
import { allowBlankInDataValidations, applyHiddenFilterButtons, normalizeWorkbookDesign } from "../ooxml/postprocess.mjs";
import {
  EXCEL_DATE_FORMAT,
  SHEET_CAPACITY,
  SHEET_ESTIMATES,
  SHEET_EXPRESS_TEMPLATE,
  SHEET_PLAN,
  SHEET_QUARTER,
  SHEET_REFS,
  SHEET_SETTINGS,
  sheetRef,
} from "./constants.mjs";
import {
  applyAction,
  applyCalculated,
  applyCapacityRangeStyle,
  applyHeader,
  applyInput,
  applyPlain,
  applyTitle,
  colors,
  fontStyle,
  setWidths,
} from "./design.mjs";
import {
  addDecimalValidation,
  addListValidation,
  addNonNegativeValidation,
  addTable,
  addWholeValidation,
} from "./validations.mjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "../..");
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
const limits = await loadWorkbookLimits(
  process.env.QUARTER_PLANNING_LIMITS_PATH
    ? path.resolve(process.env.QUARTER_PLANNING_LIMITS_PATH)
    : undefined,
);
const layout = deriveWorkbookLayout(limits);

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
    const count = Math.max(0, Math.min(limits.teamMembers, Number(item.count ?? item.people ?? item.fte ?? 0) || 0));
    for (let roleIndex = 1; roleIndex <= count && rows.length < limits.teamMembers; roleIndex += 1) {
      rows.push([role, `${role} ${roleIndex}`, 1, "", "", "", "", "", "", "", "", ""]);
    }
  }
  while (rows.length < limits.teamMembers) rows.push(["", "", "", "", "", "", "", "", "", "", "", ""]);
  return rows;
}

function teamMemberVacationDayFormulas() {
  return Array.from({ length: limits.teamMembers }, (_, index) => {
    const row = layout.teamMembers.dataStartRow + index;
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

import { buildSettingsSheet } from "../sheets/settings.mjs";
import { buildQuarterSheet } from "../sheets/quarter.mjs";
import { buildReferencesSheet } from "../sheets/references.mjs";
import { buildCapacitySheet } from "../sheets/capacity.mjs";
import { buildEstimatesSheet } from "../sheets/estimates.mjs";
import { buildPlanSheet } from "../sheets/plan.mjs";

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

const sheetContext = {
  addDecimalValidation, addListValidation, addNonNegativeValidation, addTable, addWholeValidation,
  applyAction, applyCalculated, applyCapacityRangeStyle, applyHeader, applyInput, applyPlain, applyTitle,
  colors, dateFromIso, EXCEL_DATE_FORMAT, fontStyle, layout, limits, setWidths, sheetRef, testData,
  teamCompositionCount, teamMemberAllocationSumFormula, teamMemberRows, teamMemberVacationDayFormulas,
  teamMemberVacationSumFormula, taskEstimateCsvCaption,
  SHEET_QUARTER, SHEET_REFS, SHEET_SETTINGS,
};

buildSettingsSheet({ ...sheetContext, settings });
buildQuarterSheet({ ...sheetContext, quarter });
const { quarterPlanStatuses } = buildReferencesSheet({ ...sheetContext, refs });
buildCapacitySheet({ ...sheetContext, capacity });
buildEstimatesSheet({ ...sheetContext, estimates });
buildPlanSheet({ ...sheetContext, plan, quarterPlanStatuses });

await fs.mkdir(path.dirname(outputPath), { recursive: true });
await appendWorksheetFromFile(workbook.raw, expressEstimateTemplatePath, SHEET_EXPRESS_TEMPLATE);
await workbook.raw.xlsx.writeFile(outputPath);
await allowBlankInDataValidations(outputPath);
await applyHiddenFilterButtons(outputPath);
await normalizeWorkbookDesign(outputPath);
console.log(`EXPORTED ${outputPath}`);
