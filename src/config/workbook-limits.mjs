import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const moduleDir = path.dirname(fileURLToPath(import.meta.url));
export const DEFAULT_WORKBOOK_LIMITS_PATH = path.resolve(moduleDir, "../../config/workbook-limits.json");

export const WORKBOOK_LIMIT_KEYS = Object.freeze([
  "teamMembers",
  "holidayRows",
  "taskRows",
  "activePlanRows",
  "greyZoneRows",
]);

const SAFE_MAXIMUMS = Object.freeze({
  teamMembers: 500,
  holidayRows: 366,
  taskRows: 5000,
  activePlanRows: 1000,
  greyZoneRows: 1000,
});

export function validateWorkbookLimits(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new Error("Workbook limits must be a JSON object.");
  }
  const unexpected = Object.keys(value).filter((key) => !WORKBOOK_LIMIT_KEYS.includes(key));
  if (unexpected.length) throw new Error(`Unexpected workbook limit keys: ${unexpected.join(", ")}`);

  const limits = {};
  for (const key of WORKBOOK_LIMIT_KEYS) {
    const number = value[key];
    if (!Number.isSafeInteger(number) || number <= 0) {
      throw new Error(`Workbook limit ${key} must be a positive safe integer.`);
    }
    if (number > SAFE_MAXIMUMS[key]) {
      throw new Error(`Workbook limit ${key} exceeds the safe maximum ${SAFE_MAXIMUMS[key]}.`);
    }
    limits[key] = number;
  }
  return Object.freeze(limits);
}

export async function loadWorkbookLimits(limitsPath = DEFAULT_WORKBOOK_LIMITS_PATH) {
  return validateWorkbookLimits(JSON.parse(await fs.readFile(limitsPath, "utf8")));
}

export function deriveWorkbookLayout(input) {
  const limits = validateWorkbookLimits(input);
  const teamMembers = { headerRow: 23, dataStartRow: 24 };
  teamMembers.dataEndRow = teamMembers.headerRow + limits.teamMembers;
  teamMembers.tableRange = `A${teamMembers.headerRow}:L${teamMembers.dataEndRow}`;

  const holidays = { headerRow: 13, dataStartRow: 14 };
  holidays.dataEndRow = holidays.headerRow + limits.holidayRows;
  holidays.tableRange = `A${holidays.headerRow}:D${holidays.dataEndRow}`;

  const taskEstimates = { headerRow: 3, dataStartRow: 4, jqlActionCell: "F2" };
  taskEstimates.dataEndRow = taskEstimates.headerRow + limits.taskRows;
  taskEstimates.tableRange = `C${taskEstimates.headerRow}:M${taskEstimates.dataEndRow}`;

  const activePlan = { titleRow: 1, headerRow: 5, dataStartRow: 6 };
  activePlan.dataEndRow = activePlan.headerRow + limits.activePlanRows;
  activePlan.jqlActionCell = `I${activePlan.headerRow - 1}`;
  activePlan.titleRange = `A${activePlan.titleRow}:T${activePlan.titleRow}`;
  activePlan.tableRange = `A${activePlan.headerRow}:T${activePlan.dataEndRow}`;

  const greyZone = {
    titleRow: activePlan.dataEndRow + 3,
    headerRow: activePlan.dataEndRow + 5,
  };
  greyZone.dataStartRow = greyZone.headerRow + 1;
  greyZone.dataEndRow = greyZone.headerRow + limits.greyZoneRows;
  greyZone.jqlActionCell = `I${greyZone.headerRow - 1}`;
  greyZone.titleRange = `A${greyZone.titleRow}:T${greyZone.titleRow}`;
  greyZone.tableRange = `A${greyZone.headerRow}:T${greyZone.dataEndRow}`;

  const backlog = {
    actionRow: greyZone.dataEndRow + 4,
    headerRow: greyZone.dataEndRow + 5,
  };
  backlog.titleRow = backlog.headerRow - 2;
  backlog.dataStartRow = backlog.headerRow + 1;
  backlog.dataEndRow = backlog.headerRow + limits.taskRows;
  backlog.actionCell = `A${backlog.actionRow}`;
  backlog.jqlActionCell = `I${backlog.actionRow}`;
  backlog.titleRange = `A${backlog.titleRow}:T${backlog.titleRow}`;
  backlog.tableRange = `A${backlog.headerRow}:T${backlog.dataEndRow}`;

  const references = { jqlClipboardCell: "Z1" };

  return Object.freeze({ limits, teamMembers, holidays, taskEstimates, activePlan, greyZone, backlog, references });
}

export function taskEstimateCsvCaption(limits) {
  return `Импорт CSV (до ${validateWorkbookLimits(limits).taskRows})`;
}

export function renderVbaLimitsModule(input) {
  const limits = validateWorkbookLimits(input);
  const layout = deriveWorkbookLayout(limits);
  return [
    "' AUTO-GENERATED from config/workbook-limits.json. DO NOT EDIT.",
    "Option Explicit",
    "",
    `Public Const QP_TEAM_MEMBER_ROWS As Long = ${limits.teamMembers}`,
    `Public Const QP_HOLIDAY_ROWS As Long = ${limits.holidayRows}`,
    `Public Const QP_TASK_ROWS As Long = ${limits.taskRows}`,
    `Public Const QP_ACTIVE_PLAN_ROWS As Long = ${limits.activePlanRows}`,
    `Public Const QP_GREY_ZONE_ROWS As Long = ${limits.greyZoneRows}`,
    "",
    `Public Const QP_TEAM_MEMBER_HEADER_ROW As Long = ${layout.teamMembers.headerRow}`,
    `Public Const QP_TEAM_MEMBER_FIRST_ROW As Long = ${layout.teamMembers.dataStartRow}`,
    `Public Const QP_TEAM_MEMBER_LAST_ROW As Long = ${layout.teamMembers.dataEndRow}`,
    `Public Const QP_TEAM_MEMBER_TABLE_RANGE As String = "${layout.teamMembers.tableRange}"`,
    `Public Const QP_HOLIDAY_HEADER_ROW As Long = ${layout.holidays.headerRow}`,
    `Public Const QP_HOLIDAY_FIRST_ROW As Long = ${layout.holidays.dataStartRow}`,
    `Public Const QP_HOLIDAY_LAST_ROW As Long = ${layout.holidays.dataEndRow}`,
    `Public Const QP_HOLIDAY_TABLE_RANGE As String = "${layout.holidays.tableRange}"`,
    `Public Const QP_TASK_HEADER_ROW As Long = ${layout.taskEstimates.headerRow}`,
    `Public Const QP_TASK_FIRST_ROW As Long = ${layout.taskEstimates.dataStartRow}`,
    `Public Const QP_TASK_LAST_ROW As Long = ${layout.taskEstimates.dataEndRow}`,
    `Public Const QP_TASK_TABLE_RANGE As String = "${layout.taskEstimates.tableRange}"`,
    `Public Const QP_TASK_JQL_ACTION_CELL As String = "${layout.taskEstimates.jqlActionCell}"`,
    `Public Const QP_ACTIVE_PLAN_HEADER_ROW As Long = ${layout.activePlan.headerRow}`,
    `Public Const QP_ACTIVE_PLAN_FIRST_ROW As Long = ${layout.activePlan.dataStartRow}`,
    `Public Const QP_ACTIVE_PLAN_LAST_ROW As Long = ${layout.activePlan.dataEndRow}`,
    `Public Const QP_ACTIVE_PLAN_TABLE_RANGE As String = "${layout.activePlan.tableRange}"`,
    `Public Const QP_ACTIVE_PLAN_JQL_ACTION_CELL As String = "${layout.activePlan.jqlActionCell}"`,
    `Public Const QP_GREY_ZONE_HEADER_ROW As Long = ${layout.greyZone.headerRow}`,
    `Public Const QP_GREY_ZONE_FIRST_ROW As Long = ${layout.greyZone.dataStartRow}`,
    `Public Const QP_GREY_ZONE_LAST_ROW As Long = ${layout.greyZone.dataEndRow}`,
    `Public Const QP_GREY_ZONE_TABLE_RANGE As String = "${layout.greyZone.tableRange}"`,
    `Public Const QP_GREY_ZONE_JQL_ACTION_CELL As String = "${layout.greyZone.jqlActionCell}"`,
    `Public Const QP_BACKLOG_HEADER_ROW As Long = ${layout.backlog.headerRow}`,
    `Public Const QP_BACKLOG_ACTION_CELL As String = "${layout.backlog.actionCell}"`,
    `Public Const QP_BACKLOG_JQL_ACTION_CELL As String = "${layout.backlog.jqlActionCell}"`,
    `Public Const QP_BACKLOG_FIRST_ROW As Long = ${layout.backlog.dataStartRow}`,
    `Public Const QP_BACKLOG_LAST_ROW As Long = ${layout.backlog.dataEndRow}`,
    `Public Const QP_BACKLOG_TABLE_RANGE As String = "${layout.backlog.tableRange}"`,
    `Public Const QP_JQL_CLIPBOARD_CELL As String = "${layout.references.jqlClipboardCell}"`,
    "",
  ].join("\r\n");
}
