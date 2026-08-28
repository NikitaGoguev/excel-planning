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
  refs.getRange("A1:G1").merge();
  refs.getRange("A1").values = [["Правила планирования"]];
  applyTitle(refs.getRange("A1:G1"));
  refs.getRange("A3:C8").values = [
    ["Экспертиза", "Название", "Waterfall порядок"],
    ["AN", "Аналитика", 1],
    ["DE", "Дизайн", 1],
    ["BE", "Backend-разработка", 2],
    ["FE", "Frontend-разработка", 3],
    ["QA", "Тестирование", 4],
  ];
  applyHeader(refs.getRange("A3:C3"));
  applyPlain(refs.getRange("A4:C8"));
  addTable(refs, "A3:C8", "tblExpertise");
  const quarterPlanStatuses = [
    "Готова аналитика",
    "Готова разработка",
    "Готова разработка (бэк)",
    "Готова разработка (фронт)",
    "Готово к релизу",
    "ПРОМ",
    "Отложено",
  ];
  refs.getRange("A11:G11").merge();
  refs.getRange("A11").values = [["Правила квартального планирования"]];
  refs.getRange("A11:G11").format = {
    fill: colors.section,
    font: fontStyle({ bold: true, color: colors.text }),
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };
  refs.getRange("A13:G33").values = [
    ["Правило", "Описание", null, null, null, null, null],
    ["Единица оценки", "Оценки задач указаны в человеко-днях.", null, null, null, null, null],
    ["Источник задач", "Первичный источник задач: лист 03_Оценка задач.", null, null, null, null, null],
    ["Стартовое состояние", "При первичной сборке все задачи попадают в бэклог.", null, null, null, null, null],
    ["Переходы", "Бэклог -> План; План -> Серая зона; План -> Бэклог; Серая зона -> Бэклог.", null, null, null, null, null],
    ["Трудозатраты", "Трудозатраты показывают нормализованные человеко-дни: DE = CEILING(DE / фокус-фактор дизайнеров, 1), остальные этапы = CEILING(оценка / фокус-фактор основной команды, 1).", null, null, null, null, null],
    ["Календарная длительность", "Календарные длительности рассчитываются по дням доступности сотрудников и могут быть больше трудозатрат из-за аллокации и отпусков.", null, null, null, null, null],
    ["Пустые оценки", "Пустые DE/AN/BE/FE/QA считаются нулем.", null, null, null, null, null],
    ["Фокус-фактор DE", "Для DE используется отдельный фокус-фактор дизайнеров с листа 02_Capacity; он применяется и к трудозатратам, и к дневному прогрессу.", null, null, null, null, null],
    ["Ресурсы AN", "AN выполняют аналитики и лид по аналитике.", null, null, null, null, null],
    ["Ресурсы DE", "DE выполняют сотрудники с ролью Дизайнер.", null, null, null, null, null],
    ["Ресурсы BE", "BE выполняют BE-разработчики и тим-лид / техлид.", null, null, null, null, null],
    ["Ресурсы FE/QA", "FE выполняют FE-разработчики; QA выполняют тестировщики.", null, null, null, null, null],
    ["Порядок расчета", "Задачи в квартальном плане рассчитываются сверху вниз по строкам секции.", null, null, null, null, null],
    ["Greedy ресурсов", "В каждый рабочий день для этапа выбирается доступный сотрудник нужной роли с максимальной аллокацией; при равной аллокации выбирается первый сверху в tblTeamMembers.", null, null, null, null, null],
    ["Аллокации и отпуска", "Дневной прогресс равен аллокации выбранного сотрудника, умноженной на фокус-фактор этапа. Сотрудник недоступен во все даты своего отпуска; при наличии другого доступного сотрудника работа продолжается им.", null, null, null, null, null],
    ["Даты готовности", "Даты считаются только для секции Квартальный план; для серой зоны и бэклога очищаются.", null, null, null, null, null],
    ["DE параллельно AN", "DE стартует одновременно с AN, а при пустом AN — с первой доступной даты задачи. DE резервирует дизайнеров между задачами, но не задает старт BE/FE/QA и не влияет на Завершение (план).", null, null, null, null, null],
    ["AN/BE overlap", "Если AN больше настроенного лага, BE может стартовать после заданного числа рабочих дней AN и после выполнения заданной доли AN; если AN завершилась раньше, BE стартует после AN.", null, null, null, null, null],
    ["FE overlap", "Если BE больше 5 ч/д, FE может стартовать после 5 рабочих дней BE и после выполнения 50% BE; если BE завершился раньше, FE стартует после BE.", null, null, null, null, null],
    ["QA overlap", "\u041f\u043e\u043b\u043e\u0432\u0438\u043d\u0430 QA \u043c\u043e\u0436\u0435\u0442 \u0441\u0442\u0430\u0440\u0442\u043e\u0432\u0430\u0442\u044c \u043f\u043e\u0441\u043b\u0435 \u0437\u0430\u0432\u0435\u0440\u0448\u0435\u043d\u0438\u044f \u0430\u043d\u0430\u043b\u0438\u0442\u0438\u043a\u0438; \u0432\u0442\u043e\u0440\u0430\u044f \u043f\u043e\u043b\u043e\u0432\u0438\u043d\u0430 QA \u0441\u0442\u0430\u0440\u0442\u0443\u0435\u0442 \u043f\u043e\u0441\u043b\u0435 \u0437\u0430\u0432\u0435\u0440\u0448\u0435\u043d\u0438\u044f BE \u0438 FE.", null, null, null, null, null],
  ];
  refs.getRange("B13:G33").merge(true);
  applyHeader(refs.getRange("A13:G13"));
  applyPlain(refs.getRange("A14:G33"));
  refs.getRange("A35:E41").values = [
    ["Код", "Правило", "Значение", "Ед.", "Описание"],
    ["AN_BE_MIN_WORKDAYS", "AN -> BE: минимальный лаг", 5, "раб. дни", "Минимум рабочих дней от старта AN до возможного старта BE для длинной аналитики."],
    ["AN_BE_COMPLETION_PERCENT", "AN -> BE: готовность AN", 0.5, "%", "Доля расчетной длительности AN, после которой BE может стартовать для длинной аналитики."],
    ["BE_FE_MIN_WORKDAYS", "BE -> FE: минимальный лаг", 5, "раб. дни", "Минимум рабочих дней от старта BE до возможного старта FE для длинной backend-разработки."],
    ["BE_FE_COMPLETION_PERCENT", "BE -> FE: готовность BE", 0.5, "%", "Доля расчетной длительности BE, после которой FE может стартовать для длинной backend-разработки."],
    ["QA_FIRST_PART_PERCENT", "QA: первая часть", 0.5, "%", "Доля QA, которая может выполняться после завершения аналитики."],
    ["RESOURCE_BALANCE_RISK_PERCENT", "Баланс ресурсов: риск", 0.2, "%", "Риск, добавляемый к оценкам активного плана при расчете баланса ресурсов."],
  ];
  applyHeader(refs.getRange("A35:E35"));
  applyPlain(refs.getRange("A36:E41"));
  addTable(refs, "A35:E41", "tblPlanningRuleSettings");
  addWholeValidation(refs.getRange("C36:C36"), 0, 1000);
  addWholeValidation(refs.getRange("C38:C38"), 0, 1000);
  addDecimalValidation(refs.getRange("C37:C37"), 0, 1);
  addDecimalValidation(refs.getRange("C39:C41"), 0, 1);
  refs.getRange("C36:C36").setNumberFormat("0");
  refs.getRange("C38:C38").setNumberFormat("0");
  refs.getRange("C37:C37").setNumberFormat("0%");
  refs.getRange("C39:C41").setNumberFormat("0%");
  refs.freezePanes.freezeRows(3);
  setWidths(refs, { A: 190, B: 270, C: 95, D: 85, E: 350, F: 40, G: 110 });
  return { quarterPlanStatuses };
}
