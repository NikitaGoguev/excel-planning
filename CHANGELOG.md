# Changelog

Все заметные изменения QuarterPlan Excel документируются в этом файле.

## [Unreleased]

### Changed

- structural workbook limits centralized in `config/workbook-limits.json` with generated VBA constants;
- portable builder split into orchestration, sheet builders, design/validation helpers and OOXML post-processing;
- `ThisWorkbook` reduced to events and compatibility wrappers; VBA logic split into a component manifest;
- scheduler extracted into an Excel-independent Variant-matrix domain module.

### Added

- config schema/alternate-layout tests and canonical XLSX package fingerprint;
- 32 test-only Excel COM scheduler business scenarios and `npm run verify:scheduler`.

## [1.0.0] - 2026-08-25

Первый публичный продуктовый релиз.

### Added

- чистые XLSM/XLSX release-профили и SHA-256 checksums;
- tag-driven GitHub Release workflow;
- публичная документация, security policy и issue templates;
- автоматический public-data scan и release-profile verification;
- обезличенные demo-данные и продуктовые скриншоты.

### Changed

- pre-commit использует переносимый static gate;
- workbook metadata нормализуется как `QuarterPlan Excel`;
- README ориентирован на пользователя, а технические детали вынесены в docs.

### Security

- удалены прежние корпоративные-looking данные, reference PNG и архивные baseline ZIP;
- публичный XLSM остаётся неподписанным и сопровождается checksum-инструкциями.
