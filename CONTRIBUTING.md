# Contributing to QuarterPlan Excel

Спасибо за интерес к проекту. Для изменений workbook, VBA и build pipeline особенно важны воспроизводимость и совместимость с desktop Excel.

## Локальная проверка

```text
npm ci
npm run verify
npm run verify:release
```

Для изменений VBA или пользовательских действий дополнительно требуется Windows desktop Excel:

```text
npm run sync:vba
npm run verify:excel
```

## Pull request

- описывайте пользовательский сценарий и ожидаемое поведение;
- редактируйте source-файлы, затем пересобирайте outputs;
- не патчите `vbaProject.bin` и generated workbook вручную;
- обновляйте contracts и документацию вместе с изменением интерфейсов;
- не добавляйте реальные названия команд, сотрудников, внутренние URL или рабочие тикеты;
- прикладывайте скриншот только с обезличенными demo-данными.

Bug reports и feature requests можно создать через шаблоны GitHub Issues.
