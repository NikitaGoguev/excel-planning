export function buildPlanSheet(context) {
  const {
    addDecimalValidation, addListValidation, addNonNegativeValidation, addTable, addWholeValidation,
    applyAction, applyCalculated, applyCapacityRangeStyle, applyHeader, applyInput, applyPlain, applyTitle,
    colors, dateFromIso, EXCEL_DATE_FORMAT, fontStyle, layout, limits, setWidths, sheetRef, testData,
    teamCompositionCount, teamMemberAllocationSumFormula, teamMemberRows, teamMemberVacationDayFormulas,
    teamMemberVacationSumFormula, taskEstimateCsvCaption,
    SHEET_QUARTER, SHEET_REFS, SHEET_SETTINGS,
  } = context;
  const { plan, quarterPlanStatuses } = context;
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

  applyPlanSection(
    "Квартальный план",
    layout.activePlan.titleRange,
    layout.activePlan.tableRange,
    `A${layout.activePlan.headerRow}:S${layout.activePlan.headerRow}`,
    layout.activePlan.dataStartRow,
    layout.activePlan.dataEndRow,
    "tblPlanActive",
    makePlanRows(limits.activePlanRows),
  );
  applyPlanSection(
    "Серая зона",
    layout.greyZone.titleRange,
    layout.greyZone.tableRange,
    `A${layout.greyZone.headerRow}:S${layout.greyZone.headerRow}`,
    layout.greyZone.dataStartRow,
    layout.greyZone.dataEndRow,
    "tblPlanGrey",
    makePlanRows(limits.greyZoneRows),
  );
  applyPlanSection(
    "Бэклог",
    layout.backlog.titleRange,
    layout.backlog.tableRange,
    `A${layout.backlog.headerRow}:S${layout.backlog.headerRow}`,
    layout.backlog.dataStartRow,
    layout.backlog.dataEndRow,
    "tblPlanBacklog",
    makePlanRows(limits.taskRows, testData.taskEstimates ?? [], "toPlan"),
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

  for (const [startRow, endRow] of [
    [layout.activePlan.dataStartRow, layout.activePlan.dataEndRow],
    [layout.greyZone.dataStartRow, layout.greyZone.dataEndRow],
    [layout.backlog.dataStartRow, layout.backlog.dataEndRow],
  ]) {
    addExcludedEstimateFormat(`L${startRow}:N${endRow}`, `$R${startRow}="Готова аналитика"`);
    addExcludedEstimateFormat(`N${startRow}:N${endRow}`, `$R${startRow}="Готова разработка"`);
    addExcludedEstimateFormat(`M${startRow}:N${endRow}`, `$R${startRow}="Готова разработка (бэк)"`);
    addExcludedEstimateFormat(`L${startRow}:L${endRow}`, `$R${startRow}="Готова разработка (фронт)"`);
    addExcludedEstimateFormat(`N${startRow}:N${endRow}`, `$R${startRow}="Готова разработка (фронт)"`);
    addExcludedEstimateFormat(`K${startRow}:N${endRow}`, `$R${startRow}="Отложено"`);
  }

  plan.getRange(`A${layout.activePlan.headerRow}:E${layout.activePlan.headerRow}`).format.horizontalAlignment = "left";
  plan.getRange(`A${layout.greyZone.headerRow}:E${layout.greyZone.headerRow}`).format.horizontalAlignment = "left";
  plan.getRange(`A${layout.backlog.headerRow}:E${layout.backlog.headerRow}`).format.horizontalAlignment = "left";
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
  plan.getRange(layout.backlog.actionCell).values = [["++"]];
  plan.getRange(layout.backlog.actionCell).format = {
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
}
