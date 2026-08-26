export function buildQuarterSheet(context) {
  const {
    addDecimalValidation, addListValidation, addNonNegativeValidation, addTable, addWholeValidation,
    applyAction, applyCalculated, applyCapacityRangeStyle, applyHeader, applyInput, applyPlain, applyTitle,
    colors, dateFromIso, EXCEL_DATE_FORMAT, fontStyle, layout, limits, setWidths, sheetRef, testData,
    teamCompositionCount, teamMemberAllocationSumFormula, teamMemberRows, teamMemberVacationDayFormulas,
    teamMemberVacationSumFormula, taskEstimateCsvCaption,
    SHEET_QUARTER, SHEET_REFS, SHEET_SETTINGS,
  } = context;
  const { quarter } = context;
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
  quarter.getRange(layout.holidays.tableRange).values = [
    ["Дата", "Название", "Комментарий", "Учитывать"],
    ...Array.from({ length: limits.holidayRows }, () => ["", "", "", ""]),
  ];
  applyHeader(quarter.getRange("A13:D13"));
  applyCalculated(quarter.getRange(`A${layout.holidays.dataStartRow}:A${layout.holidays.dataEndRow}`));
  applyInput(quarter.getRange(`B${layout.holidays.dataStartRow}:C${layout.holidays.dataEndRow}`));
  applyCalculated(quarter.getRange(`D${layout.holidays.dataStartRow}:D${layout.holidays.dataEndRow}`));
  quarter.getRange(`A${layout.holidays.dataStartRow}:A${layout.holidays.dataEndRow}`).setNumberFormat(EXCEL_DATE_FORMAT);
  quarter.getRange("A14").formulas = [[`=IF(OR($B$4="",ROW(A1)>$B$8),"",WORKDAY($B$4-1,ROW(A1)))`]];
  quarter.getRange(`A${layout.holidays.dataStartRow}:A${layout.holidays.dataEndRow}`).fillDown();
  quarter.getRange("D14").formulas = [[`=IF($A14="","","Да")`]];
  quarter.getRange(`D${layout.holidays.dataStartRow}:D${layout.holidays.dataEndRow}`).fillDown();
  addListValidation(quarter.getRange(`D${layout.holidays.dataStartRow}:D${layout.holidays.dataEndRow}`), ["Да", "Нет"]);
  addTable(quarter, layout.holidays.tableRange, "tblHolidays");
  quarter.freezePanes.freezeRows(13);
  setWidths(quarter, { A: 190, B: 240, C: 460, D: 120 });
}
