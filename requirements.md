# Требования к окружению и переносимость сборки

## Текущий статус

Базовая XLSX-сборка и упаковка XLSM выполняются переносимыми Node.js-скриптами.
Они не требуют Codex runtime, PowerShell, desktop Excel или персональных путей.
Полная проверка VBA через Excel остаётся Windows-only; нативные Mac VBA sync и
Excel acceptance пока не реализованы.

## Переносимое окружение

| Компонент | Требование |
| --- | --- |
| Node.js | ветка `26.x` |
| npm | версия, поставляемая с поддерживаемым Node.js |
| XLSX engine | публичный `exceljs 4.4.0` |
| OOXML/ZIP | публичный `jszip 3.10.1` |
| Установка | `npm ci` по зафиксированному lock-файлу |

Обязательные portable-команды:

```text
npm ci
npm run env:check
npm run build:xlsx
npm run build:xlsm
npm run build
npm run verify:static
npm run verify
npm run build:release
npm run verify:release
```

`npm run build` использует обезличенный demo-профиль и поддерживает COM acceptance. `npm run build:release` использует отдельный clean-профиль без команды, руководителя и задач, пишет versioned XLSM/XLSX и `SHA256SUMS.txt` в `dist/`.

`npm run build:xlsx` создаёт `outputs/quarter_planning_step1.xlsx` полностью в
Node.js. Builder создаёт листы, таблицы, формулы, validations, стили, условное
форматирование и копирует `100_Шаблон экспресс оценки` из исходного asset.
Точечные package-контракты — скрытые filter buttons, `allowBlank`, шрифты и
высоты строк — нормализуются через JSZip.

`npm run build:xlsm` встраивает сохранённый целиком
`assets/vba/vbaProject.step2.bin`, выставляет macro-enabled content type и code
names. Скрипт не изменяет бинарные потоки VBA-проекта и не открывает Excel.

## Windows desktop Excel

Следующие команды намеренно отделены от portable build:

```text
npm run sync:vba
npm run verify:excel
npm run preview
```

Для `sync:vba` и `verify:excel` требуются Windows, установленный и активированный
desktop Excel, разрешение **Trust access to the VBA project object model** и
доступ к `VBProject.VBComponents`. `preview` также использует отдельный экземпляр
Windows Excel, но отсутствие preview backend не блокирует build или static gate.

Перед Excel automation канонические книги должны быть закрыты. Автоматизация не
должна закрывать видимые пользовательские экземпляры Excel и не должна удалять
lock-файлы открытых книг.

## Контракты переносимости

`scripts/verify_static_contracts.mjs` проверяет:

- отсутствие приватных runtime imports и персональных абсолютных путей в
  исполняемых скриптах и конфигурации;
- отсутствие `xl/vbaProject.bin` в XLSX и его наличие в XLSM;
- macro-enabled content type, workbook/worksheet code names;
- семь листов и их фиксированный порядок;
- диапазоны ListObject-таблиц, action cells и видимость технических колонок;
- Calibri 11, отсутствие custom row heights и скрытые filter buttons;
- наличие status conditional formatting;
- VBA source contracts и нормализованный structural snapshot.

CI выполняет `npm ci`, portable build и static gate на Windows, macOS и Linux под
Node.js 26.

## Оставшаяся работа для нативной автоматизации на Mac

XLSX/XLSM build уже переносим. Не реализованы только операции, которым нужен
Excel object model:

1. VBA source sync через AppleScript и доверенную helper-книгу.
2. Mac acceptance harness с запуском публичных макросов и сканированием ошибок.
3. Необязательный preview adapter для Excel for Mac.

Для полного Mac gate потребуется self-hosted Mac с установленным Microsoft 365
Excel, активной GUI-сессией, разрешёнными Apple Events, макросами и доступом к VBA
project object model. Обычный headless cloud runner выполняет только Node/static
часть.
