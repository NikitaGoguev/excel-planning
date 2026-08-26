# Development and release

## Requirements

- Node.js 26.x and npm;
- locked dependencies installed through `npm ci`;
- Windows desktop Excel only for VBA sync, COM acceptance and PNG screenshots.

## Developer build

```text
npm run build
npm run verify
npm run verify:scheduler
npm run audit:deps
```

Canonical demo outputs remain in `outputs/` and are used by static/COM contracts.

## Release build

```text
npm run build:release
npm run verify:release
npm run verify:release:excel
```

`RELEASE_VERSION` defaults to `package.json` and may be provided without or with a leading `v`. Release artifacts are written to ignored `dist/`.

Supported environment interface:

- `QUARTER_PLANNING_DATA_PATH`;
- `QUARTER_PLANNING_XLSX_OUTPUT`;
- `QUARTER_PLANNING_XLSX_INPUT`;
- `QUARTER_PLANNING_XLSM_OUTPUT`;
- `RELEASE_VERSION`.

Developer-only alternate layout builds may additionally set `QUARTER_PLANNING_LIMITS_PATH`; the public source of truth remains `config/workbook-limits.json`.

## Structural limits

Edit only `config/workbook-limits.json`, then run:

```text
npm run generate:limits
npm run build:xlsx
npm run sync:vba
npm run build:xlsm
npm run verify:excel
```

The resolver derives team-member, holiday, estimate, active-plan, grey-zone and backlog ranges, including the bulk-action cell. `taskRows` controls both estimates and backlog capacity. Generated VBA constants are checked for drift by `verify:static`.

## VBA workflow

After changing VBA source:

```text
npm run build
npm run sync:vba
npm run build:xlsm
npm run normalize
npm run verify:excel
```

Sync reads the component manifest, updates its document/standard modules, removes only stale managed `QuarterPlan*` standard modules, saves through Excel, sanitizes Office metadata, copies the validated template and extracts the complete VBA project. It never patches binary VBA streams.

`npm run verify:scheduler` imports the test-only harness into a temporary XLSM copy and executes 32 business scenarios. `npm run verify:excel` includes this gate before the regular workbook acceptance.

## Publishing

1. Run portable, release and Excel acceptance gates.
2. Ensure the working tree contains the intended regenerated outputs.
3. Push the release commit.
4. Push an annotated `v*` tag matching `package.json`.
5. The release workflow builds clean XLSM/XLSX, verifies public data, writes checksums and creates GitHub Release.
