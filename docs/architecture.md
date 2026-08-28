# Architecture

QuarterPlan Excel использует модель **workbook as code**.

```text
config/workbook-limits.json + JSON profile
    ↓
layout resolver + sheet builders
    ↓
base XLSX
    ↓ + validated Excel-saved VBA project
macro-enabled XLSM
    ↓
static contracts + Excel COM acceptance
```

## Слои

- `src/workbook/build.mjs` оркестрирует сборку; `src/sheets/*.mjs` строят отдельные листы, а `src/workbook/design.mjs`, `validations.mjs` и `src/ooxml/*` обслуживают оформление, validation и OOXML post-processing.
- XLSM packager добавляет сохранённый Excel VBA project целиком и выставляет OOXML code names/content types.
- `ThisWorkbook` содержит только events и compatibility wrappers. Domain/UI-код разделён между `QuarterPlanCommon`, `QuarterPlanUi`, `QuarterPlanTeamCapacity`, `QuarterPlanTaskEstimates`, `QuarterPlanImportExport` и `QuarterPlanPlanActions`.
- `QuarterPlanScheduler` — чистый domain engine: получает Variant matrices задач, ресурсов, праздников и правил и возвращает effort, ready date и `U:AJ`, не обращаясь к Excel Object Model. DE планируется параллельно AN с отдельным фокус-фактором и не входит в зависимость готовности.
- Contracts проверяют структуру пакета, диапазоны таблиц, public macros, оформление и запрещённые VBA-паттерны.
- Excel COM acceptance запускает пользовательские сценарии в настоящем desktop Excel.

## Build profiles

- `data/test_data_quarter_planning.json` — обезличенный demo для разработки, COM smoke и скриншотов.
- `data/release_blank_quarter_planning.json` — чистый пользовательский релиз без команды, руководителя и задач.

Release orchestration передаёт data/input/output paths через environment interface и не меняет canonical demo outputs.

## Источники истины

- Structural capacity: `config/workbook-limits.json`; `src/config/workbook-limits.mjs` валидирует лимиты и выводит все диапазоны.
- VBA limits: auto-generated `assets/vba/QuarterPlanLimits_module.txt`; drift проверяется автоматически, ручное редактирование запрещено.
- VBA component manifest: `contracts/vba.contract.json`; исходники модулей находятся в `assets/vba/*_module.txt`, события — в `ThisWorkbook_holiday_macro.txt`.
- Embedded VBA: Excel-saved `vbaProject.step2.bin`; бинарное редактирование запрещено.
- Workbook behavior/layout: builder source и `contracts/*.json`.
- Test fixtures: `data/test_data_quarter_planning.json` и `assets/import1.csv`.
- Scheduler scenarios: test-only `tests/vba/QuarterPlanSchedulerTests_module.txt`, временно импортируемый только в копию XLSM.
