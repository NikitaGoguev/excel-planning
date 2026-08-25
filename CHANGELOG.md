# Changelog

Все заметные изменения QuarterPlan Excel документируются в этом файле.

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
