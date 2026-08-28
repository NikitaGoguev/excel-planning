import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";
import test from "node:test";

import ExcelJS from "exceljs";

import {
  DEFAULT_WORKBOOK_LIMITS_PATH,
  deriveWorkbookLayout,
  loadWorkbookLimits,
  renderVbaLimitsModule,
  taskEstimateCsvCaption,
  validateWorkbookLimits,
} from "../src/config/workbook-limits.mjs";

const defaults = {
  teamMembers: 20,
  holidayRows: 20,
  taskRows: 100,
  activePlanRows: 20,
  greyZoneRows: 20,
};

test("default limits reproduce the v1.0.0 layout", () => {
  const layout = deriveWorkbookLayout(defaults);
  assert.equal(layout.teamMembers.tableRange, "A23:L43");
  assert.equal(layout.holidays.tableRange, "A13:D33");
  assert.equal(layout.taskEstimates.tableRange, "C3:M103");
  assert.equal(layout.taskEstimates.jqlActionCell, "F2");
  assert.equal(layout.activePlan.tableRange, "A5:T25");
  assert.equal(layout.activePlan.jqlActionCell, "I4");
  assert.equal(layout.greyZone.tableRange, "A30:T50");
  assert.equal(layout.greyZone.jqlActionCell, "I29");
  assert.equal(layout.backlog.tableRange, "A55:T155");
  assert.equal(layout.backlog.actionCell, "A54");
  assert.equal(layout.backlog.jqlActionCell, "I54");
  assert.equal(layout.references.jqlClipboardCell, "Z1");
  assert.equal(taskEstimateCsvCaption(defaults), "Импорт CSV (до 100)");
});

test("alternate limits derive every structural section", () => {
  const layout = deriveWorkbookLayout({ teamMembers: 4, holidayRows: 5, taskRows: 8, activePlanRows: 3, greyZoneRows: 2 });
  assert.equal(layout.teamMembers.tableRange, "A23:L27");
  assert.equal(layout.holidays.tableRange, "A13:D18");
  assert.equal(layout.taskEstimates.tableRange, "C3:M11");
  assert.equal(layout.taskEstimates.jqlActionCell, "F2");
  assert.equal(layout.activePlan.tableRange, "A5:T8");
  assert.equal(layout.activePlan.jqlActionCell, "I4");
  assert.equal(layout.greyZone.tableRange, "A13:T15");
  assert.equal(layout.greyZone.jqlActionCell, "I12");
  assert.equal(layout.backlog.actionCell, "A19");
  assert.equal(layout.backlog.jqlActionCell, "I19");
  assert.equal(layout.backlog.tableRange, "A20:T28");
});

test("invalid limits fail closed", () => {
  assert.throws(() => validateWorkbookLimits({ ...defaults, taskRows: 0 }), /positive safe integer/);
  assert.throws(() => validateWorkbookLimits({ ...defaults, extra: 1 }), /Unexpected/);
  assert.throws(() => validateWorkbookLimits({ ...defaults, holidayRows: 367 }), /safe maximum/);
});

test("tracked VBA limits are generated from the JSON config", async () => {
  const limits = await loadWorkbookLimits(DEFAULT_WORKBOOK_LIMITS_PATH);
  const actual = await fs.readFile(new URL("../assets/vba/QuarterPlanLimits_module.txt", import.meta.url), "utf8");
  const normalize = (value) => `${value.replace(/\r?\n/g, "\r\n").trimEnd()}\r\n`;
  assert.equal(normalize(actual), normalize(renderVbaLimitsModule(limits)));
});

test("alternate limits build resizes all structural sections", async () => {
  const temporaryDirectory = await fs.mkdtemp(path.join(os.tmpdir(), "quarterplan-limits-"));
  try {
    const limitsPath = path.join(temporaryDirectory, "limits.json");
    const outputPath = path.join(temporaryDirectory, "alternate.xlsx");
    const alternate = { teamMembers: 4, holidayRows: 5, taskRows: 8, activePlanRows: 3, greyZoneRows: 2 };
    await fs.writeFile(limitsPath, JSON.stringify(alternate), "utf8");
    const build = spawnSync(process.execPath, ["scripts/build_quarter_planning_step1.mjs"], {
      cwd: path.resolve(import.meta.dirname, ".."),
      env: {
        ...process.env,
        QUARTER_PLANNING_LIMITS_PATH: limitsPath,
        QUARTER_PLANNING_XLSX_OUTPUT: outputPath,
      },
      encoding: "utf8",
    });
    assert.equal(build.status, 0, build.stderr || build.stdout);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(outputPath);
    assert.equal(workbook.getWorksheet("00_Настройки").getTable("tblTeamMembers").table.tableRef, "A23:L27");
    assert.equal(workbook.getWorksheet("01_Настройки квартала").getTable("tblHolidays").table.tableRef, "A13:D18");
    assert.equal(workbook.getWorksheet("03_Оценка задач").getTable("tblTaskEstimates").table.tableRef, "C3:M11");
    assert.equal(workbook.getWorksheet("04_Квартальный план").getTable("tblPlanActive").table.tableRef, "A5:T8");
    assert.equal(workbook.getWorksheet("04_Квартальный план").getTable("tblPlanGrey").table.tableRef, "A13:T15");
    assert.equal(workbook.getWorksheet("04_Квартальный план").getTable("tblPlanBacklog").table.tableRef, "A20:T28");
    assert.equal(workbook.getWorksheet("03_Оценка задач").getCell("D2").value, "Импорт CSV (до 8)");
    assert.equal(workbook.getWorksheet("03_Оценка задач").getCell("F2").value, "JQL");
    assert.equal(workbook.getWorksheet("04_Квартальный план").getCell("A19").value, "++");
    assert.equal(workbook.getWorksheet("04_Квартальный план").getCell("I4").value, "JQL");
    assert.equal(workbook.getWorksheet("04_Квартальный план").getCell("I12").value, "JQL");
    assert.equal(workbook.getWorksheet("04_Квартальный план").getCell("I19").value, "JQL");
  } finally {
    const tempRoot = path.resolve(os.tmpdir());
    const resolved = path.resolve(temporaryDirectory);
    if (!resolved.startsWith(`${tempRoot}${path.sep}`) || !path.basename(resolved).startsWith("quarterplan-limits-")) {
      throw new Error(`Refusing to remove unexpected temporary directory: ${resolved}`);
    }
    await fs.rm(resolved, { recursive: true, force: true });
  }
});
