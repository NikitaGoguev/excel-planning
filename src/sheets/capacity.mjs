export function buildCapacitySheet(context) {
  const {
    addDecimalValidation, addListValidation, addNonNegativeValidation, addTable, addWholeValidation,
    applyAction, applyCalculated, applyCapacityRangeStyle, applyHeader, applyInput, applyPlain, applyTitle,
    colors, dateFromIso, EXCEL_DATE_FORMAT, fontStyle, layout, limits, setWidths, sheetRef, testData,
    teamCompositionCount, teamMemberAllocationSumFormula, teamMemberRows, teamMemberVacationDayFormulas,
    teamMemberVacationSumFormula, taskEstimateCsvCaption,
    SHEET_QUARTER, SHEET_REFS, SHEET_SETTINGS,
  } = context;
  const { capacity } = context;
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
}
