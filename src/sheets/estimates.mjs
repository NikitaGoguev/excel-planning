export function buildEstimatesSheet(context) {
  const {
    addDecimalValidation, addListValidation, addNonNegativeValidation, addTable, addWholeValidation,
    applyAction, applyCalculated, applyCapacityRangeStyle, applyHeader, applyInput, applyPlain, applyTitle,
    colors, dateFromIso, EXCEL_DATE_FORMAT, fontStyle, layout, limits, setWidths, sheetRef, testData,
    teamCompositionCount, teamMemberAllocationSumFormula, teamMemberRows, teamMemberVacationDayFormulas,
    teamMemberVacationSumFormula, taskEstimateCsvCaption,
    SHEET_QUARTER, SHEET_REFS, SHEET_SETTINGS,
  } = context;
  const { estimates } = context;
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
  estimates.getRange("D2").values = [[taskEstimateCsvCaption(limits)]];
  applyAction(estimates.getRange("D2"));
  estimates.getRange(layout.taskEstimates.jqlActionCell).values = [["JQL"]];
  applyAction(estimates.getRange(layout.taskEstimates.jqlActionCell));
  estimates.getRange("H1:L1").merge();
  estimates.getRange("H1").values = [["Статистика"]];
  applyTitle(estimates.getRange("H1:L1"));
  estimates.getRange("H2:L2").formulas = [[
    `SUM(H${layout.taskEstimates.dataStartRow}:H${layout.taskEstimates.dataEndRow})`,
    `SUM(I${layout.taskEstimates.dataStartRow}:I${layout.taskEstimates.dataEndRow})`,
    `SUM(J${layout.taskEstimates.dataStartRow}:J${layout.taskEstimates.dataEndRow})`,
    `SUM(K${layout.taskEstimates.dataStartRow}:K${layout.taskEstimates.dataEndRow})`,
    `SUM(L${layout.taskEstimates.dataStartRow}:L${layout.taskEstimates.dataEndRow})`,
  ]];
  applyCalculated(estimates.getRange("H2:L2"));
  estimates.getRange("H2:L2").setNumberFormat("0.00");
  estimates.getRange(`A${layout.taskEstimates.headerRow}:M${layout.taskEstimates.dataEndRow}`).values = [
    [">", "x", "Приоритет", "Направление", "Описание", "ЗНИ/Jira", "Примечание", "DE", "AN", "BE", "FE", "QA", "Комментарий"],
    ...Array.from({ length: limits.taskRows }, (_, index) => {
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
            task.DE ?? "",
            task.AN ?? "",
            task.BE ?? "",
            task.FE ?? "",
            task.QA ?? "",
            task.comment ?? "",
          ]
        : ["", "", "", "", "", "", "", "", "", "", "", "", ""];
    }),
  ];
  estimates.getRange(`N${layout.taskEstimates.headerRow}:N${layout.taskEstimates.dataEndRow}`).values = [
    ["+"],
    ...Array.from({ length: limits.taskRows }, (_, index) => [testData.taskEstimates?.[index] ? "+" : ""]),
  ];
  estimates.getRange(`O${layout.taskEstimates.headerRow}:O${layout.taskEstimates.dataEndRow}`).values = [
    ["Экспорт"],
    ...Array.from({ length: limits.taskRows }, () => [""]),
  ];
  applyHeader(estimates.getRange("A3:M3"));
  applyHeader(estimates.getRange("N3"));
  applyHeader(estimates.getRange("O3"));
  applyAction(estimates.getRange(`A${layout.taskEstimates.dataStartRow}:B${layout.taskEstimates.dataEndRow}`));
  applyInput(estimates.getRange(`C${layout.taskEstimates.dataStartRow}:M${layout.taskEstimates.dataEndRow}`));
  applyAction(estimates.getRange(`N${layout.taskEstimates.dataStartRow}:N${layout.taskEstimates.dataEndRow}`));
  applyAction(estimates.getRange(`O${layout.taskEstimates.dataStartRow}:O${layout.taskEstimates.dataEndRow}`));
  estimates.getRange(`C${layout.taskEstimates.dataStartRow}:C${layout.taskEstimates.dataEndRow}`).setNumberFormat("0");
  estimates.getRange(`H${layout.taskEstimates.dataStartRow}:L${layout.taskEstimates.dataEndRow}`).setNumberFormat("0.00");
  addWholeValidation(estimates.getRange(`C${layout.taskEstimates.dataStartRow}:C${layout.taskEstimates.dataEndRow}`), 1, 999);
  addNonNegativeValidation(estimates.getRange(`H${layout.taskEstimates.dataStartRow}:L${layout.taskEstimates.dataEndRow}`));
  addTable(estimates, layout.taskEstimates.tableRange, "tblTaskEstimates");
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
    L: 56,
    M: 519,
    N: 46,
    O: 140,
  });
}
