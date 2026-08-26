---
name: excel-macro-workbook
description: Use when working on this repository's Excel workbook build pipeline, especially when generating or modifying the macro-enabled quarterly planning workbook (`.xlsm`), syncing VBA source into the embedded project, preserving Excel-saved macro integrity, rebuilding outputs, and validating that the final workbook opens and contains the expected runnable macros.
---

# Excel Macro Workbook

Use this skill for repo tasks that change workbook generation, VBA automation, `.xlsm` packaging, or Excel COM validation.

## Workflow

1. Edit source files, not generated workbooks:
   - `scripts/build_quarter_planning_step1.mjs`
   - `src/workbook/`, `src/sheets/`, and `src/ooxml/`
   - `config/workbook-limits.json`
   - `scripts/create_quarter_planning_xlsm.mjs`
   - VBA sources listed in `contracts/vba.contract.json`
   - `data/test_data_quarter_planning.json`
   - `data/release_blank_quarter_planning.json`

2. Rebuild the base workbook:
   - `npm ci`
   - `npm run build:xlsx`

3. Normalize workbook design after any Excel COM save that can rewrite row heights or font records:
   - `npm run normalize`

4. If VBA changed, resync through desktop Excel:
   - `npm run sync:vba` (Windows only)

5. Rebuild the macro workbook:
   - `npm run build:xlsm`

6. Validate the output:
   - `.xlsx` has no `xl/vbaProject.bin`
   - `.xlsm` has `xl/vbaProject.bin`
   - Excel COM opens the workbook
   - expected public macros are present in `VBProject`
   - no formula error scan matches

   The portable validation entrypoint is:
   - `npm run verify`

   For a static non-COM check, use:
   - `npm run verify:static`

   For Windows desktop Excel acceptance, use:
   - `npm run verify:excel`

   For the isolated 52-scenario exact-output domain scheduler gate, use:
   - `npm run verify:scheduler`

   For clean public artifacts and the public-data gate, use:
   - `npm run build:release`
   - `npm run verify:release`

   The machine-readable contracts live in `contracts/*.json`, and the Git pre-commit hook in `.githooks/pre-commit` runs the portable static gate.

## Constraints

- Use Node.js 26.x with the locked public `exceljs` and `jszip` dependencies. Do not add personal runtime paths or private bundled packages.
- Keep XLSX generation and XLSM packaging free of PowerShell, Excel COM, and platform-specific filesystem paths.
- Never binary-edit `vbaProject.bin`.
- Treat `config/workbook-limits.json` as the only source for structural capacities. Run `npm run generate:limits`; never hand-edit `QuarterPlanLimits_module.txt`.
- Keep `ThisWorkbook` limited to events and stable compatibility wrappers. Add/update managed modules through `contracts/vba.contract.json` so Excel sync can preserve document/sheet components.
- Keep `QuarterPlanScheduler` free of Excel Object Model and UI state; pass only Variant matrices/scalars and test changes with the test-only COM harness.
- The scheduler COM harness replaces the module in a temporary XLSM with the current `assets/vba/QuarterPlanScheduler_module.txt` source before running all 52 scenarios, so `verify:scheduler` does not depend on a prior VBA sync. Desktop COM acceptance separately requires the embedded scheduler to match that source.
- Keep demo data anonymized and build public releases from the clean release profile. Public workbook metadata must identify `QuarterPlan Excel`, not a local Office user or absolute path.
- Prefer updating VBA source text and resyncing through Excel.
- Preserve the Excel-saved template and validated `vbaProject.step2.bin`.
- Release orchestration may redirect data/input/output paths only through `QUARTER_PLANNING_DATA_PATH`, `QUARTER_PLANNING_XLSX_OUTPUT`, `QUARTER_PLANNING_XLSX_INPUT`, `QUARTER_PLANNING_XLSM_OUTPUT`, and `RELEASE_VERSION`.
- When project behavior changes, update this skill together with `AGENTS.md`.
- When adding or changing enforced workbook/VBA behavior, update `contracts/*.json` and the verification scripts in the same change.

## Project-specific notes

- Current stable Mac-safe recalculation path uses `RunQuarterPlanRecalculate`.
- Current stable backlog load path uses `RunQuarterPlanReloadBacklog`.
- Workbook design is contract-enforced: all generated sheets and action shapes use Calibri 11; only headers, section titles, and action/button cells/shapes are bold; generated worksheets must not use custom row heights to make buttons fit.
- Hidden header dropdowns are contract-enforced for selected ListObject columns: settings/team headers, sheet 00 employee headers, holiday headers, sheet 03 user estimate columns `G:M`, and sheet 04 action/status/comment/resource columns `A:E` and `J:S` across all plan sections.
- Sheet `00_Настройки` keeps `tblTeamComposition` at `A13:B19`; `tblTeamMembers` starts at `A23:L23`, and its data-row capacity comes from `teamMembers` (default range `A23:L43`).
- When synchronizing `tblTeamMembers`, write only editable columns and leave the three builder-provided vacation-day calculated columns untouched. Do not reassign `.FormulaR1C1` across those full columns: Excel 2016 can fail with error `1004`. The COM gate covers `Тестировщик` changing from `3` to `2` and verifies that the existing formula text is preserved.
- Sheet `02_Capacity!E8:E19` are effective team counts and vacation person-days for downstream capacity and resource-balance formulas; `F8:F13` auto-calculate counts from sheet 00 allocations, `F14:F19` auto-calculate weighted vacation person-days from sheet 00 vacations and allocations, and `G8:G19` are manual overrides.
- All visible workbook dates must display as `ДД-ММ-ГГГГ`. In the JS workbook builder, use `dd-mm-yyyy` so the generated XLSX style stays a real date format. In Excel COM formatting code, use `NumberFormatLocal = "ДД-ММ-ГГГГ"`. In VBA, assign the existing Unicode-safe `DateFormatLocalText()` result to `NumberFormatLocal`; a raw Cyrillic VBA format literal is corrupted when the project is opened on Mac. Russian desktop Excel treats `dd-mm-yyyy` assigned through COM/VBA as literal text. Do not introduce `yyyy-mm-dd` formats or hardcoded style ids that can display dates as serial numbers.
- Planning overlap parameters are workbook settings in `99_Справочники!tblPlanningRuleSettings`; keep VBA reading these values instead of hardcoding numeric rule constants.
- Current rule codes are `AN_BE_MIN_WORKDAYS`, `AN_BE_COMPLETION_PERCENT`, `BE_FE_MIN_WORKDAYS`, `BE_FE_COMPLETION_PERCENT`, `QA_FIRST_PART_PERCENT`, and `RESOURCE_BALANCE_RISK_PERCENT`.
- `04_Квартальный план!J2:N4` contains the formula-driven resource balance block. Keep labels in `J2:J4` (`Емкость`, `План+риск`, `Баланс`), AN/BE/FE/QA values in `K:N`, and formulas based on `tblPlanActive`, `02_Capacity!E27:E30`, and `RESOURCE_BALANCE_RISK_PERCENT`. Do not show the risk as a separate top-row cell. The user-facing ready-date header is `Завершение (план)`.
- `04_Квартальный план` technical schedule columns `T:AF` are user-visible and must not be hidden; they show AN/BE/FE/QA durations, dates, and diagnostics.
- `RunQuarterPlanRecalculate` schedules dates from `00_Настройки!tblTeamMembers`, not from aggregate/manual override capacity values. Manual overrides on `02_Capacity!G8:G19` still affect capacity/balance formulas, but the per-day calendar uses actual employees, allocation percentages, and vacation ranges.
- Sheet 04 scheduling uses one available employee per stage per working day. Daily progress is `Аллокация * 02_Capacity!E20`; full-day vacation ranges block that employee, and another employee with the same expertise may continue the remaining work.
- Default ranges remain `tblTaskEstimates=C3:M103`, `tblPlanBacklog=A55:S155`, and bulk action `A54`, but they are derived from `taskRows`; never duplicate these literals in builder, handwritten VBA, or verifiers.
- On workbook open, `04_Квартальный план` recreates shape buttons over `G2` and `P2` and binds them to those public macros.
- On workbook open, `03_Оценка задач` refreshes decomposition action cells. Shape buttons over populated action cells are recreated on sheet activation and by `RunTaskEstimateRepairActionButtons`, using the same shape-backed public-macro pattern as `04_Квартальный план`; this avoids Excel activation crashes from pre-created shapes on sheet 03.
- The sheet 03 template row count and CSV caption derive from `taskRows`; changing the config requires rebuild, Excel VBA sync and COM verification.
- `03_Оценка задач!M3` is `Экспорт`; `M4:M103` contains shape-backed per-root express-estimate export actions. `Примечание` is a visible user field in column `G`.
- The top action layout on sheet 03 is fixed, but `D2` uses dynamic caption `Импорт CSV (до <taskRows>)`.
- `03_Оценка задач!H1:K1` is a merged `Статистика` title block; `H2:K2` must contain formula-driven totals for `AN`, `BE`, `FE`, and `QA` based on `tblTaskEstimates`.
- `03_Оценка задач!D1` is a shape-backed XLSX import button with caption `Импорт`. It imports files created by the sheet 03 `Экспорт` button, restores hierarchy from description indentation, and always appends after the last task/subtree.
- `03_Оценка задач!D2` is a shape-backed CSV import button whose limit comes from `taskRows`. It imports UTF-8 comma-delimited CSV files by headers only; use `assets/import1.csv` for checks.
- `03_Оценка задач!A2:B2` is a shape-backed `Сбросить` button that confirms before clearing the task estimate table, action cells, and technical decomposition fields.
- `03_Оценка задач!A2:B2` and `C2` must restore the configured task range if worksheet rows were manually deleted.
- `03_Оценка задач!C2` is a shape-backed `Обновить` button and is the normal user path for refreshing `>`/`x`, parent formulas, and row-action shapes after manual edits.
- `03_Оценка задач!C1` is a shape-backed `Экспорт` button that exports visible user columns `C:M` to standalone `.xlsx`; parent estimates must be values, not formulas.
- Root-row `M` export actions must use the visible sheet `100_Шаблон экспресс оценки`, copied from `assets/Шаблон Экспресс оценки.xlsx` and placed after `99_Справочники`, to create `<ЗНИ/Jira> Экспресс-оценка.xlsx`. Fill `E2:K2` with the root ticket, `E3:K3` with the root description, and write root/child/grandchild rows from row 17 using description merges `A:D`, `B:D`, `C:D` and estimates `AN/BE/FE/QA` in `G/I/J/K`; leave `Design` blank.
- For workbook action buttons, follow the current `04_Квартальный план` pattern used by `Загрузить 03 в бэклог` and `Посчитать план`: keep the cell area styled in the sheet, recreate a shape button over that area on workbook open, and bind the shape to a stable public macro from a standard module.
- `04_Квартальный план!Q2` is a shape-backed `Экспорт плана` button bound to `RunQuarterPlanExportBacklog`. It exports non-empty `tblPlanActive` rows to standalone `.xlsx`, writing only visible user columns `F:S`; the default filename is `<Команда>_<Год>_Q<Квартал>_квартальный план_<yyyy-mm-dd_hh-nn>.xlsx`.
- Keep the Q2 export button caption Mac-safe: use the visible cell text for the shape caption and the existing `ChrW$` helper text for VBA-generated Cyrillic pieces instead of raw Cyrillic string literals.
- Row actions inside the plan tables are cell values with shape buttons recreated over populated action cells on workbook open/refresh. `Workbook_SheetSelectionChange` and double-click remain fallbacks.
- Current compact captions are `+` for move one task to plan, `\` for move to grey zone, `-` for move to backlog, `↑` for moving a task one row up within the same section, and `↓` for moving a task one row down within the same section.
- Sheet `04_Квартальный план` column `R` (`Статус на конец квартала`) uses the status list from `99_Справочники!tblStatuses`; the allowed values are `Готова аналитика`, `Готова разработка`, `Готова разработка (бэк)`, `Готова разработка (фронт)`, `Готово к релизу`, `ПРОМ`, and `Отложено`. Blank status is treated as `ПРОМ` for resource balance, schedule calculation, and excluded-estimate marking.
- Status controls which active-plan stages consume resources and appear in technical schedule columns: `Готова аналитика` calculates only AN, `Готова разработка` calculates AN/BE/FE, `Готова разработка (бэк)` calculates AN/BE, `Готова разработка (фронт)` calculates AN/FE, blank/`Готово к релизу`/`ПРОМ` calculates the full AN/BE/FE/QA path, and `Отложено` calculates no stages. `Трудозатраты` always remains the normalized full AN+BE+FE+QA effort and is not inflated by allocation or vacation pauses.
- Sheet `03_Оценка задач` uses compact shape-backed captions `>` to create a decomposition child, `x` to delete a decomposed row/subtree, and `Экспорт` in root rows to export the express-estimate template. Delete must show a modal choice to drop estimates, transfer up, or transfer down; transfer uses same-level siblings with parent fallback for child/grandchild when the selected direction has no sibling.
- Generate every Cyrillic string written or compared by the sheet 03 VBA flow through `TextFromCodes()`: action captions, restored table/statistics headers, export sheet names and filename parts, XLSX/CSV header matching, and import/export/reset dialogs. Raw Cyrillic VBA literals can become mojibake in Excel 2016 or across platforms. Export and import must share the same header helper array; the COM gate corrupts and rebuilds the headers, checks the default filename helper, validates the exported headers, and imports that exported file back.
- Sheet 03 CSV import must use the project pure-VBA parser and `Application.GetOpenFilename`; avoid `FileDialog`, QueryTables, Power Query, `OpenText`, ActiveX, and `CreateObject`. Append imports after the last task/subtree; replace clears the whole task estimate table and technical fields first. If CSV rows exceed available capacity, show a modal error and leave the workbook unchanged.
- The bulk backlog action `++` is a separate top-level cell derived by the layout resolver (default `A54`); VBA uses the generated constant.
- `+`, `\`, `-`, `↑`, `↓`, and `++` should use stable public macros from the standard module through shape `OnAction`, with cell-event handlers kept only as fallback for direct cell selection.
- Shape `OnAction` values should be plain public macro names, not workbook-qualified names and not macro calls with arguments, to avoid Mac Excel failures and silent row-button creation failures.
- Row action shapes should call `RunQuarterPlanCellAction`; the handler should resolve the clicked shape through `Application.Caller` and then use `TopLeftCell` for the action cell.
- Row action shape names must stay short, e.g. `qpa_b_01_1`. Excel on Windows and Mac can truncate long `Application.Caller` shape names, which breaks lookup by name.
- Task estimate action shape names must stay short, e.g. `tea_01_1`.
- `RunQuarterPlanRepairActionButtons` is the manual repair macro for the action UI; keep it available when changing the button pipeline.
- `RunTaskEstimateRepairActionButtons` is the manual repair macro for sheet 03 decomposition action cells and shape buttons.
- Sheet 03 row add/delete performance depends on the optimized refresh path: mark the earliest dirty decomposition row, refresh only that range through the used estimate rows plus one spare row, and synchronize only missing/blank shape buttons. Do not touch properties on existing buttons or return normal actions to full delete/recreate of all shape buttons; reserve full rebuild for repair flows.
- Keep Mac-critical VBA paths free of ActiveX-only helpers such as `CreateObject("Scripting.Dictionary")`; prefer plain VBA arrays/collections for lookup caches.
- Do not remove action shapes in `Workbook_BeforeClose`; cleanup-on-close can leave a still-open workbook with missing buttons and stale close-state flags if closing is cancelled.
- VBA UI entrypoints and workbook action handlers must not fail silently. Avoid bare `On Error Resume Next`, empty `SafeExit` branches, or guard-clause exits that make a clicked button do nothing with no feedback. If an action cannot run, show a modal `MsgBox` with the procedure name, error number, error description, and relevant context such as `Application.Caller`, shape name, sheet, table, or cell address.
- Standard-module public macros should delegate to `ThisWorkbook` methods. Keep sheet-name constants and task-moving logic in `ThisWorkbook` to avoid duplicated Russian sheet-name strings and silent Mac failures.
- VBA should resolve key worksheets by table names (`tblPlanActive`, `tblTaskEstimates`, `tblPlanningRuleSettings`, `tblHolidays`) instead of Russian worksheet names where possible.
- Cell-action handlers should refresh action shapes in place after successful moves and should not select `D2` or another distant cell, because that causes visible jumps on Mac.
- The `++` action must use a batched move implementation instead of row-by-row compact/refresh loops.
- Keep workbook automation conservative on Mac; avoid heavy actions from selection events, `Application.OnTime`, or old `ThisWorkbook` methods with long `ByRef` argument lists.
- `tblTaskEstimates` must start at `03_Оценка задач!C3:M3` and end at the row resolved from `taskRows`; loading to sheet 04 includes only roots and aggregates leaf descendants.
