export function buildSettingsSheet(context) {
  const {
    addDecimalValidation, addListValidation, addNonNegativeValidation, addTable, addWholeValidation,
    applyAction, applyCalculated, applyCapacityRangeStyle, applyHeader, applyInput, applyPlain, applyTitle,
    colors, dateFromIso, EXCEL_DATE_FORMAT, fontStyle, layout, limits, setWidths, sheetRef, testData,
    teamCompositionCount, teamMemberAllocationSumFormula, teamMemberRows, teamMemberVacationDayFormulas,
    teamMemberVacationSumFormula, taskEstimateCsvCaption,
    SHEET_QUARTER, SHEET_REFS, SHEET_SETTINGS,
  } = context;
  const { settings } = context;
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
  addWholeValidation(settings.getRange("B14:B19"), 0, limits.teamMembers);
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
  settings.getRange("H13").formulas = [[`=IF(SUM(B14:B19)>${limits.teamMembers},"Превышен лимит ${limits.teamMembers} сотрудников","")`]];
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
  settings.getRange(layout.teamMembers.tableRange).values = [
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
  settings.getRange(`F${layout.teamMembers.dataStartRow}:F${layout.teamMembers.dataEndRow}`).formulas = teamMemberVacationDayFormulas().map((row) => [row[0]]);
  settings.getRange(`I${layout.teamMembers.dataStartRow}:I${layout.teamMembers.dataEndRow}`).formulas = teamMemberVacationDayFormulas().map((row) => [row[1]]);
  settings.getRange(`L${layout.teamMembers.dataStartRow}:L${layout.teamMembers.dataEndRow}`).formulas = teamMemberVacationDayFormulas().map((row) => [row[2]]);
  applyHeader(settings.getRange("A23:L23"));
  applyCalculated(settings.getRange(`A${layout.teamMembers.dataStartRow}:A${layout.teamMembers.dataEndRow}`));
  applyInput(settings.getRange(`B${layout.teamMembers.dataStartRow}:L${layout.teamMembers.dataEndRow}`));
  applyCalculated(settings.getRange(`F${layout.teamMembers.dataStartRow}:F${layout.teamMembers.dataEndRow}`));
  applyCalculated(settings.getRange(`I${layout.teamMembers.dataStartRow}:I${layout.teamMembers.dataEndRow}`));
  applyCalculated(settings.getRange(`L${layout.teamMembers.dataStartRow}:L${layout.teamMembers.dataEndRow}`));
  settings.getRange(`C${layout.teamMembers.dataStartRow}:C${layout.teamMembers.dataEndRow}`).setNumberFormat("0%");
  settings.getRange(`D${layout.teamMembers.dataStartRow}:E${layout.teamMembers.dataEndRow}`).setNumberFormat(EXCEL_DATE_FORMAT);
  settings.getRange(`G${layout.teamMembers.dataStartRow}:H${layout.teamMembers.dataEndRow}`).setNumberFormat(EXCEL_DATE_FORMAT);
  settings.getRange(`J${layout.teamMembers.dataStartRow}:K${layout.teamMembers.dataEndRow}`).setNumberFormat(EXCEL_DATE_FORMAT);
  settings.getRange(`F${layout.teamMembers.dataStartRow}:F${layout.teamMembers.dataEndRow}`).setNumberFormat("0");
  settings.getRange(`I${layout.teamMembers.dataStartRow}:I${layout.teamMembers.dataEndRow}`).setNumberFormat("0");
  settings.getRange(`L${layout.teamMembers.dataStartRow}:L${layout.teamMembers.dataEndRow}`).setNumberFormat("0");
  addDecimalValidation(settings.getRange(`C${layout.teamMembers.dataStartRow}:C${layout.teamMembers.dataEndRow}`), 0, 1);
  addTable(settings, layout.teamMembers.tableRange, "tblTeamMembers");
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
}
