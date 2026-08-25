# Development and release

## Requirements

- Node.js 24.x and npm;
- locked dependencies installed through `npm ci`;
- Windows desktop Excel only for VBA sync, COM acceptance and PNG screenshots.

## Developer build

```text
npm run build
npm run verify
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

## VBA workflow

After changing VBA source:

```text
npm run build
npm run sync:vba
npm run build:xlsm
npm run normalize
npm run verify:excel
```

Sync saves through Excel, sanitizes Office metadata, copies the validated template and extracts the complete VBA project. It never patches binary VBA streams.

## Publishing

1. Run portable, release and Excel acceptance gates.
2. Ensure the working tree contains the intended regenerated outputs.
3. Push the release commit.
4. Push an annotated `v*` tag matching `package.json`.
5. The release workflow builds clean XLSM/XLSX, verifies public data, writes checksums and creates GitHub Release.
