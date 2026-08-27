# Project Notes for Agents

## Project

This project builds an Excel workbook for IT quarterly planning.

Current workbook outputs:

- `outputs/quarter_planning_step1.xlsx`
- `outputs/quarter_planning_step2.xlsm`

Current source files:

- `scripts/build_quarter_planning_step1.mjs` is the stable CLI entrypoint; `src/workbook/build.mjs` orchestrates builders from `src/sheets/`.
- `scripts/create_quarter_planning_xlsm.mjs` converts the generated `.xlsx` into the macro-enabled `.xlsm`.
- `scripts/build_release.mjs` builds clean, versioned public XLSX/XLSM artifacts in `dist/`.
- `config/workbook-limits.json` is the only source of truth for structural capacities; `src/config/workbook-limits.mjs` derives ranges and generates VBA constants.
- `scripts/verify_public_release.mjs` verifies the clean profile, checksums, packages, and public-data blocklist.
- `scripts/sanitize_workbook_metadata.mjs` removes personal Office metadata without editing VBA binary streams.
- `scripts/lib/exceljs_compat.mjs` provides the small ExcelJS range adapter used by the portable workbook builder.
- `scripts/sync_vba_from_source.ps1` updates `ThisWorkbook` through desktop Excel COM, saves the workbook through Excel, copies the validated template, and extracts `assets/vba/vbaProject.step2.bin`.
- `assets/vba/ThisWorkbook_holiday_macro.txt` is the VBA source of truth.
- `contracts/vba.contract.json` lists every managed VBA component. `ThisWorkbook` contains events/compatibility wrappers; domain and UI logic lives in the listed `QuarterPlan*` modules.
- `assets/vba/QuarterPlanLimits_module.txt` is auto-generated; never edit it manually.
- `assets/vba/vbaProject.step2.bin` is the validated Excel-saved VBA project embedded by the `.xlsm` builder.
- `assets/vba/quarter_planning_macro_template.xlsm` preserves the Excel-saved macro template.
- `data/test_data_quarter_planning.json` stores reusable test data, including sheet 03 task estimates.
- `data/release_blank_quarter_planning.json` stores the clean public release profile.
- `.codex/skills/excel-macro-workbook/SKILL.md` documents the project-specific workflow for building and validating the macro-enabled workbook.

Public product screenshots:

- `docs/images/settings.png`
- `docs/images/estimates.png`
- `docs/images/plan.png`
- `assets/import1.csv` is the sample UTF-8 comma-delimited CSV for checking sheet 03 import. It contains the required headers `Ключ запроса` and `Тема`, plus the optional `Пользовательское поле (Тип Темы (ЭПИКА))` direction header.

## Workbook Structure

The workbook currently contains:

- `00_Настройки`: base workbook settings with merged `C:G` comments, team role counts in `tblTeamComposition`, role comments in merged `C:G` fields, and a 20-row `tblTeamMembers` employee UI with allocation percent and up to 3 vacation date ranges plus calculated working-day counts per employee.
- `01_Настройки квартала`: quarter dates, manual working days, calculated calendar weekdays, calculated weekday holidays to fill, and holiday list.
- `02_Capacity`: capacity form styled to match the reference asset; includes formulas and saved test values.
- `03_Оценка задач`: task estimate input table `tblTaskEstimates`, with `>`/`x` decomposition actions in `A:B` and root express-estimate export actions in `M`; VBA maintains technical `_TaskId`, `_ParentTaskId`, `_Level` values outside the visible table.
- `04_Квартальный план`: resource balance strip, active plan, grey zone, and backlog tables. Section capacities come from `config/workbook-limits.json`.
- `100_Шаблон экспресс оценки`: visible sheet copied from `assets/Шаблон Экспресс оценки.xlsx`, placed after `99_Справочники`, and used as the formatted template for per-root express-estimate exports.
- `101_Списки`: visible editable source tables `tblTaskCommentArtifacts` and `tblTaskCommentAdjacentTeams` used by the sheet 03 comment-options modal.
- `99_Справочники`: expertise list `AN`, `BE`, `FE`, `QA`, waterfall/order rules, statuses, yes/no values, planning rule descriptions, and editable planning rule settings table `tblPlanningRuleSettings`.

## Important Design Decisions

- Keep the workbook as `.xlsx` without macros unless macro-enabled automation is explicitly needed.
- Use only expertise codes `AN`, `BE`, `FE`, `QA`.
- Base waterfall order is `AN -> BE -> FE -> QA`, with overlap rules below.
- Test data must stay in `data/test_data_quarter_planning.json`; do not hardcode it into the template except via the builder.
- Public release files must use `data/release_blank_quarter_planning.json`; demo data is reserved for development, COM acceptance, and screenshots.
- Do not add real employee names, internal URLs, production ticket keys, or organization-specific task descriptions to public fixtures or screenshots.
- All visible workbook dates must display as `ДД-ММ-ГГГГ` on every sheet. The JS builder may use `dd-mm-yyyy` for XLSX styles. Excel COM formatting must use `NumberFormatLocal = "ДД-ММ-ГГГГ"`; VBA must use the Unicode-safe `DateFormatLocalText()` helper rather than a raw Cyrillic literal, which is corrupted on Mac. Russian desktop Excel treats `dd-mm-yyyy` assigned through COM/VBA as literal text.
- All generated workbook sheets and action shapes must use Calibri 11. Only headers, section titles, and action/button cells/shapes should be bold. Do not set custom row heights to make overlay buttons fit; shape buttons must use the anchor cell or merge area dimensions.
- Some ListObject header dropdowns are intentionally hidden while the tables remain ListObjects: all headers on `00_Настройки!A13:B13` and `00_Настройки!A23:L23`; all headers on `01_Настройки квартала!A13:D13`; `03_Оценка задач!G3:M3`; and `04_Квартальный план` plan-table headers in columns `A:E` and `J:S` for active, grey, and backlog sections.
- Edit source files and rebuild outputs. Do not manually patch generated `.xlsx` / `.xlsm` files.
- Do not programmatically replace or binary-edit streams inside `vbaProject.bin`. Direct binary edits can make Excel repair or remove `xl/vbaProject.bin`.
- If the workbook build pipeline, VBA sync flow, validation steps, or output naming changes, update `.codex/skills/excel-macro-workbook/SKILL.md` in the same change.

## Current Automation

On `00_Настройки`:

- `tblTeamComposition` spans `A13:B19`; role comments are merged UI fields in `C13:G19` and are intentionally outside the ListObject.
- With the default config, `tblTeamMembers` spans `A23:L43`. Its capacity and final row are derived from `teamMembers` in `config/workbook-limits.json`.
- If existing `tblTeamMembers` rows remain in the same role after a composition change, manually edited employee names, allocation, and vacation dates are preserved.
- Team composition sync must write only editable `tblTeamMembers` columns and leave the builder-provided vacation-day formulas untouched. Do not assign an array to the full `DataBodyRange` or reassign `.FormulaR1C1` to entire calculated columns: both patterns can raise Excel error `1004` when a role count is reduced, and the formula-column rewrite specifically fails in Excel 2016.

On `02_Capacity`:

- `E8:E19` are the public effective team counts and vacation person-days used by downstream capacity and resource-balance formulas.
- `F8:F13` auto-calculate team counts from `00_Настройки!tblTeamMembers[Аллокация]` by role; `F14:F19` auto-calculate weighted vacation person-days from `tblTeamMembers` vacation day columns and allocation by role.
- `G8:G19` are manual overrides. If `G` is blank, `E` uses `F`; if `G` is filled, `E` uses `G`.
- Vacation day columns on `00_Настройки!tblTeamMembers` count weekdays in the vacation period minus included holidays from `01_Настройки квартала!tblHolidays`.
- Keep formulas on `E27:E30` and sheet 04 references pointed at the effective `E` values.

On `01_Настройки квартала`:

- `B7`: calendar weekdays between start and end dates.
- `B8`: weekday holidays to fill, calculated as calendar weekdays minus manually entered working days.
- In `.xlsx`, `A14:A33` and `D14:D33` are formula-backed.
- In `.xlsm`, formulas are removed from the holiday rows. VBA updates `B7:B8`, resizes `tblHolidays`, and fills exactly the needed rows after relevant setting changes.
- Rows `A14:D33` are kept as preformatted spare holiday rows while `tblHolidays` initially references only active rows. This avoids differently styled rows when the holiday count grows.

On `04_Квартальный план`:

- `G2` and `P2` are action areas; workbook open recreates shape buttons over them.
- `J2:N4` contains the formula-driven resource balance block: labels in `J2:J4` (`Емкость`, `План+риск`, `Баланс`) and AN/BE/FE/QA values in `K:N`; the risk value stays configurable in `99_Справочники!tblPlanningRuleSettings` and is not shown as a separate top-row cell.
- Stable public entrypoint macros are `RunQuarterPlanReloadBacklog`, `RunQuarterPlanRecalculate`, and `RunQuarterPlanExportBacklog`.
- `RunQuarterPlanReloadBacklog` asks for confirmation, clears sheet 04 plan/grey/backlog task data, and loads root tasks from `03_Оценка задач!tblTaskEstimates` into `tblPlanBacklog`.
- With the default config, `tblPlanBacklog` is `A55:S155` and the bulk action is `A54`. Both backlog and sheet 03 capacities derive from the same `taskRows` value.
- `RunQuarterPlanRecalculate` recalculates dates only for active plan rows through the Mac-safe scheduling path.
- `G2` (`Загрузить 03 в бэклог`), `P2` (`Посчитать план`), and `Q2` (`Экспорт плана`) remain shape-backed buttons bound to stable public macros.
- `RunQuarterPlanExportBacklog` exports non-empty `tblPlanActive` rows to a standalone `.xlsx`, writing only visible user columns `F:S`. The default filename is `<Команда>_<Год>_Q<Квартал>_квартальный план_<yyyy-mm-dd_hh-nn>.xlsx`, with team/year/quarter read from `00_Настройки`.
- The Q2 export caption must stay Mac-safe: keep the visible cell text as the caption source and avoid raw Cyrillic VBA literals for this button text/default filename parts where `ChrW$` helpers are already used.
- Technical schedule columns `T:AF` are visible by design and must not be hidden; they show AN/BE/FE/QA durations, dates, and diagnostics.
- `RunQuarterPlanRecalculate` schedules active-plan dates from `00_Настройки!tblTeamMembers`, not from the aggregate/manual override values on `02_Capacity!E8:E19`/`G8:G19`. Manual capacity overrides still affect capacity and balance formulas, not per-day scheduling.
- Sheet 04 scheduling uses one available employee per stage per working day. Daily progress is `Аллокация * 02_Capacity!E20`; vacation ranges on `tblTeamMembers` block that employee for the full date range. If another employee with the same expertise is available, the remaining work can continue with that employee; otherwise the stage waits until a matching employee is available again.
- `Статус на конец квартала` uses `99_Справочники!tblStatuses`; allowed values are `Готова аналитика`, `Готова разработка`, `Готова разработка (бэк)`, `Готова разработка (фронт)`, `Готово к релизу`, `ПРОМ`, and `Отложено`. Blank status is treated as `ПРОМ` for resource balance, schedule calculation, and excluded-estimate marking.
- Status controls which active-plan stages consume resources and appear in technical schedule columns: `Готова аналитика` calculates only AN, `Готова разработка` calculates AN/BE/FE, `Готова разработка (бэк)` calculates AN/BE, `Готова разработка (фронт)` calculates AN/FE, blank/`Готово к релизу`/`ПРОМ` calculates the full AN/BE/FE/QA path, and `Отложено` calculates no stages. `Трудозатраты` always remains the full AN+BE+FE+QA duration.
- Row actions inside the plan tables are cell values with shape buttons recreated over populated action cells on workbook open/refresh. `Workbook_SheetSelectionChange` and double-click remain fallbacks.
- Action captions are compact: `+` means move one task to plan, `\` means move to grey zone, `-` means move to backlog, `↑` moves the task one task row up within the same section, and `↓` moves it one task row down within the same section.
- The bulk backlog action `++` is a separate top-level cell (default `A54`), derived as `layout.backlog.actionCell`; VBA uses the generated `QP_BACKLOG_ACTION_CELL`.
- `+`, `\`, `-`, `↑`, `↓`, and `++` use stable public macros from the standard module through shape `OnAction`; keep the cell-event handlers only as fallback.
- Shape `OnAction` values must be plain public macro names such as `RunQuarterPlanCellAction`, not workbook-qualified names and not macro calls with arguments. Excel can silently reject argument-bearing `OnAction` strings on row shapes.
- Row action shapes must call `RunQuarterPlanCellAction`; the handler uses `Application.Caller` to find the clicked shape and then `TopLeftCell` to locate the action cell.
- Row action shape names must stay short, e.g. `qpa_b_01_1`. Excel on Windows and Mac can truncate long `Application.Caller` shape names, which breaks lookup by name.
- `RunQuarterPlanRepairActionButtons` is the manual repair macro for the action UI; it resets transient UI flags and recreates all shape buttons on `04_Квартальный план`.
- Do not remove action shapes in `Workbook_BeforeClose`. Removing them can leave an open workbook with missing buttons and stale close-state flags if the user cancels closing or a save prompt interrupts shutdown.
- VBA UI entrypoints and workbook action handlers must not fail silently. Avoid bare `On Error Resume Next`, empty `SafeExit` branches, or guard-clause exits that make a clicked button do nothing with no feedback. If an action cannot run, show a modal `MsgBox` with the procedure name, error number, error description, and relevant context such as `Application.Caller`, shape name, sheet, table, or cell address.
- Public macros in `QuarterPlanActions` should delegate to `ThisWorkbook` methods. Do not duplicate sheet-name constants or task-moving logic in the standard module; duplicated Russian sheet-name strings can silently fail on Mac.
- VBA automation should find key worksheets by table names (`tblPlanActive`, `tblTaskEstimates`, `tblPlanningRuleSettings`, `tblHolidays`) instead of Russian worksheet names where possible. This avoids Mac VBA encoding/name lookup failures.
- Cell-action handlers must not select `D2` or otherwise move the active cell after row actions; this causes visible jumps on Mac. Refresh action shapes in place after a successful move instead.
- `++` must move backlog rows to the active plan in one batched operation and refresh action cells once; do not reintroduce a row-by-row compact/refresh loop.
- Heavy actions must use the stable public macros above. Do not reintroduce `Application.OnTime` wrappers or old `ThisWorkbook` scheduling paths with large `ByRef` argument lists.
- Mac-critical VBA paths must not use ActiveX-only helpers such as `CreateObject("Scripting.Dictionary")`; use plain VBA arrays/collections instead.

On `03_Оценка задач`:

- With the default config, sheet 03 has `100` data rows and `tblTaskEstimates` spans `C3:M103`; code must use resolved/generated limits rather than these default literals.
- `M3` is `Экспорт`; `M4:M103` contains shape-backed per-root export actions. `Примечание` is restored as a visible user field in column `G`.
- `N3` and root task rows in `N4:N103` contain the `+` comment-options action outside `tblTaskEstimates`; child, grandchild, and empty rows must not receive that action.
- The `+` action opens managed UserForm `QuarterPlanTaskCommentForm`, reading current non-empty values from `tblTaskCommentArtifacts` and `tblTaskCommentAdjacentTeams`. OK appends selected values to `Комментарий` with `; ` and case-insensitive de-duplication; Cancel and window close leave the comment unchanged.
- The top action layout is fixed: `A2:B2` = `Сбросить`, `C1` = `Экспорт`, `C2` = `Обновить`, `D1` = `Импорт`, `D2` = the dynamic `Импорт CSV (до <taskRows>)`, `E1` = title `Оценка задач`.
- `H1:K1` contains the merged title `Статистика`; `H2:K2` contains formula-driven totals for `AN`, `BE`, `FE`, and `QA` from `tblTaskEstimates`.
- `D1` is a shape-backed XLSX import button with caption `Импорт`. It imports files created by the sheet 03 `Экспорт` button and always appends after the last task/subtree.
- `D2` is a shape-backed CSV import button whose caption includes the configured `taskRows`.
- `A2:B2` is a shape-backed `Сбросить` button. It must show a confirmation modal before clearing task data and restore the table to the configured task range if rows were manually deleted.
- `C2` is a shape-backed `Обновить` button. It restores the configured task range before rebuilding row actions and parent formulas.
- `C2` is a shape-backed `Обновить` button. It is the normal user path to refresh row actions `>`/`x`, parent formulas, and sheet 03 shape buttons after manual edits.
- `C1` is a shape-backed `Экспорт` button. It exports visible user columns `C:M` to a standalone `.xlsx`, with parent estimates written as values and a default filename based on `02_Capacity!E3` plus export date/time.
- Per-root `M` export actions create an `.xlsx` from the visible `100_Шаблон экспресс оценки` sheet. The default filename is `<ЗНИ/Jira> Экспресс-оценка.xlsx`; values are written to `E2:K2` and `E3:K3`, and root/child/grandchild rows are written from row 17 with descriptions in `A:D`, `B:D`, or `C:D` and estimates in `G/I/J/K`.
- CSV import expects UTF-8, comma-delimited files with headers. It reads required `Ключ запроса` -> `ЗНИ/Jira` and `Тема` -> `Описание`, plus optional `Пользовательское поле (Тип Темы (ЭПИКА))` -> `Направление`; a missing optional direction header leaves `Направление` blank, and other extra columns are ignored.
- Use `assets/import1.csv` as the canonical sample file for quick manual/COM checks of the sheet 03 CSV import.
- Import asks whether to append after the last task/subtree or replace all tasks. If the CSV has more task rows than available capacity, it shows a modal error and does not change the workbook.
- CSV import must stay on the pure-VBA parser and `Application.GetOpenFilename`; do not switch it to `FileDialog`, QueryTables, Power Query, `OpenText`, ActiveX, or `CreateObject`.
- `tblTaskEstimates` starts at `C3`; its final row is derived from `taskRows`. Decomposition action cells `A:B` stay outside the ListObject, while root express-export actions live in `M`.
- Compact action captions are `>` to create a child task, `x` to remove a decomposition subtree, and `Экспорт` in column `M` to export the root express estimate. VBA recreates shape buttons over these action cells on sheet activation or through `RunTaskEstimateRepairActionButtons`, using the same public-macro `OnAction` pattern as sheet 04. Do not switch these back to Unicode symbols inside/near the ListObject without testing Excel activation on Windows and Mac.
- All Cyrillic sheet 03 text written by VBA must be generated through the Unicode-safe `TextFromCodes()` helpers. This includes action captions, restored table/statistics headers, export sheet names and filename parts, XLSX/CSV header matching, and import/export/reset dialog text. Raw Cyrillic VBA literals can render as mojibake in Excel 2016 or on another platform. Export and import must use the same header helper array so a file produced by `Экспорт` is accepted by `Импорт`. Existing top-level shapes must refresh a mismatched caption from the safe helper when the workbook opens or the UI is repaired.
- Decomposition supports only two levels below a root task: root -> child -> grandchild. The `>` action is not shown for level 2 rows.
- Visible task fields start at column `C`; technical fields `_TaskId`, `_ParentTaskId`, and `_Level` are maintained by VBA outside `tblTaskEstimates`.
- `RunTaskEstimateRepairActionButtons` is the manual repair macro for sheet 03 decomposition action cells and shape buttons.
- For normal sheet 03 add/delete actions, keep the optimized refresh path: mark the earliest dirty decomposition row, refresh only that range through the used estimate rows plus one spare row, and synchronize only missing/blank shape buttons. Do not touch properties on existing buttons during normal refresh. Full button removal/rebuild is reserved for manual repair/rebuild flows.
- Task level is shown by indentation in `Описание`; do not add a visible level column.
- Estimates are entered only on leaf rows. Parent rows on sheet 03 show formula-backed AN/BE/FE/QA sums from direct children with a grey calculated style; child rows with grandchildren behave the same way.
- If `>` is used on a row with AN/BE/FE/QA estimates, those estimates move to the new child row and are cleared from the parent.
- `x` is shown for every decomposed row: root with descendants, child, and grandchild. It asks whether to delete estimates, transfer them up, or transfer them down; transfer targets siblings of the same level/parent, falling back to the parent for child/grandchild when no sibling exists in the selected direction.
- Sheet 03 actions use shape overlays created on sheet activation or through `RunTaskEstimateRepairActionButtons`; direct cell double-click remains only a fallback.
- During load to sheet 04, only root tasks are loaded. For decomposed roots, AN/BE/FE/QA values are calculated from leaf descendants; parent formulas on sheet 03 display direct-child sums for users.

## Planning Rules

- Task estimates are in person-days.
- `Трудозатраты` formula remains `CEILING(estimate / core team focus factor, 1)` per stage and is not inflated by allocation or vacation pauses.
- Technical stage durations and ready dates are calculated by the sheet 04 employee scheduler from `tblTeamMembers`, so low allocation and vacation gaps can increase `T:W` and `Завершение (план)`.
- Empty AN/BE/FE/QA estimates count as zero.
- Numeric overlap settings are read from `99_Справочники!tblPlanningRuleSettings`; do not hardcode these values in VBA.
- Default settings are `AN_BE_MIN_WORKDAYS=5`, `AN_BE_COMPLETION_PERCENT=50%`, `BE_FE_MIN_WORKDAYS=5`, `BE_FE_COMPLETION_PERCENT=50%`, `QA_FIRST_PART_PERCENT=50%`, `RESOURCE_BALANCE_RISK_PERCENT=20%`.
- AN resources include analysts and the analysis lead.
- BE resources include backend developers and the team/tech lead.
- FE resources include frontend developers.
- QA resources include testers.
- Resource balance uses only `tblPlanActive`: capacity from `02_Capacity!E27:E30` minus active-plan AN/BE/FE/QA estimates increased by `RESOURCE_BALANCE_RISK_PERCENT`.
- AN/BE overlap: if AN is longer than the configured lag, BE may start after the later of configured AN working days and configured AN completion percent; if AN finishes earlier, BE starts after AN.
- FE overlap: if BE is longer than the configured lag, FE may start after the later of configured BE working days and configured BE completion percent; if BE finishes earlier, FE starts after BE.
- QA overlap: configured first QA part can start after AN is complete; the second QA part starts after BE and FE are complete.
- Active plan rows are calculated top to bottom. Grey zone and backlog dates are cleared.

## Build and Verification

Run the builder from the repo root:

```text
npm ci
npm run build:xlsx
```

Clean public release build and verification:

```text
npm run build:release
npm run verify:release
```

Release artifacts are written to ignored `dist/` as versioned XLSM/XLSX files plus `SHA256SUMS.txt`. The release interface uses `QUARTER_PLANNING_DATA_PATH`, `QUARTER_PLANNING_XLSX_OUTPUT`, `QUARTER_PLANNING_XLSX_INPUT`, `QUARTER_PLANNING_XLSM_OUTPUT`, and `RELEASE_VERSION`.

The builder exports:

- `outputs/quarter_planning_step1.xlsx`
- the visible express-estimate template sheet copied in Node from the source asset

PNG previews are optional and Windows Excel-only:

```text
npm run preview
```

After any Excel COM save that can rewrite font records or row heights, normalize strict workbook design:

```text
npm run normalize
```

VBA source sync, after editing `assets/vba/ThisWorkbook_holiday_macro.txt`:

```text
npm run sync:vba
```

Macro workbook build:

```text
npm run build:xlsm
```

Before finalizing workbook changes, verify:

- the `.xlsx` exports successfully;
- `.xlsx` has no `xl/vbaProject.bin`;
- `.xlsm` has `xl/vbaProject.bin`, macro-enabled workbook content type, and `workbook.xml` contains `codeName="ThisWorkbook"`;
- `.xlsm` worksheet XML files contain code names for all generated sheets, currently `Sheet1` through `Sheet8`;
- Excel COM can open the `.xlsm`, access `VBProject.VBComponents`, load sheet 03 into backlog, move one backlog task to active plan, and run recalculation;
- the target sheet preview is readable;
- formula error scan reports no matches.

Deterministic verification is enforced by project scripts:

- `contracts/*.json` contains machine-readable workbook, sheet, VBA, and build contracts.
- `scripts/verify_static_contracts.mjs` checks package structure, table ranges, required action cells, visible technical columns, hidden filter-button contracts, workbook design contracts, required public macros, and forbidden VBA patterns without Excel.
- `scripts/verify_excel_com.ps1` opens the `.xlsm` through desktop Excel COM, checks `VBProject`, runs smoke macros, validates table ranges and shape `OnAction` values, and scans formula errors.
- `npm run verify` is the portable build and static gate; `npm run verify:excel` is the separate Windows desktop Excel acceptance entrypoint.
- `npm run verify:scheduler` replaces the scheduler in a temporary XLSM copy with the current source, imports a test-only module, and compares all 15 output fields across exactly 52 pure-engine business scenarios; it is included in `verify:excel`. Desktop COM acceptance separately rejects an embedded scheduler that is stale relative to source.
- `.githooks/pre-commit` runs the portable `npm run verify:static`; keep `git config core.hooksPath .githooks` enabled for this repository.

## File Handling

- Do not delete `outputs/~$quarter_planning_step1.xlsx` or `outputs/~$quarter_planning_step2.xlsm` if present; it likely means Excel has the workbook open.
- Preserve user-provided assets in `assets/`.
- Historical private reference assets must remain outside the public repository and must not be reintroduced from local backups.
- If Excel COM leaves a hidden `EXCEL.EXE` process that locks a file after automation, close only automation-owned or windowless Excel processes. Do not close the user's visible workbook without explicit need.
