export function buildReferencesSheet(context) {
  const {
    addDecimalValidation, addListValidation, addNonNegativeValidation, addTable, addWholeValidation,
    applyAction, applyCalculated, applyCapacityRangeStyle, applyHeader, applyInput, applyPlain, applyTitle,
    colors, dateFromIso, EXCEL_DATE_FORMAT, fontStyle, layout, limits, setWidths, sheetRef, testData,
    teamCompositionCount, teamMemberAllocationSumFormula, teamMemberRows, teamMemberVacationDayFormulas,
    teamMemberVacationSumFormula, taskEstimateCsvCaption,
    SHEET_QUARTER, SHEET_REFS, SHEET_SETTINGS,
  } = context;
  const { refs } = context;
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
  return { quarterPlanStatuses };
}
