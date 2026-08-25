# Architecture

QuarterPlan Excel использует модель **workbook as code**.

```text
JSON profile
    ↓
portable ExcelJS builder
    ↓
base XLSX
    ↓ + validated Excel-saved VBA project
macro-enabled XLSM
    ↓
static contracts + Excel COM acceptance
```

## Слои

- Builder создаёт листы, таблицы, формулы, validations и оформление.
- XLSM packager добавляет сохранённый Excel VBA project целиком и выставляет OOXML code names/content types.
- VBA обслуживает UI actions, import/export, декомпозицию, employee scheduler и shape buttons.
- Contracts проверяют структуру пакета, диапазоны таблиц, public macros, оформление и запрещённые VBA-паттерны.
- Excel COM acceptance запускает пользовательские сценарии в настоящем desktop Excel.

## Build profiles

- `data/test_data_quarter_planning.json` — обезличенный demo для разработки, COM smoke и скриншотов.
- `data/release_blank_quarter_planning.json` — чистый пользовательский релиз без команды, руководителя и задач.

Release orchestration передаёт data/input/output paths через environment interface и не меняет canonical demo outputs.

## Источники истины

- VBA source: `assets/vba/ThisWorkbook_holiday_macro.txt` и `QuarterPlanActions_module.txt`.
- Embedded VBA: Excel-saved `vbaProject.step2.bin`; бинарное редактирование запрещено.
- Workbook behavior/layout: builder source и `contracts/*.json`.
- Test fixtures: `data/test_data_quarter_planning.json` и `assets/import1.csv`.
