NPM ?= npm

.DEFAULT_GOAL := help

.PHONY: help env-check build build-xlsx build-xlsm build-release sync-vba normalize preview preview-public verify-static verify verify-release verify-excel verify-skip-build

help:
	@echo Available targets:
	@echo   make env-check          Check Node 26, dependencies, and required assets
	@echo   make build              Build XLSX and package XLSM using portable Node scripts
	@echo   make build-xlsx         Build the base XLSX workbook
	@echo   make build-xlsm         Build the XLSM from the current XLSX and VBA project
	@echo   make build-release      Build clean public XLSM/XLSX files in dist
	@echo   make sync-vba           Sync VBA sources through desktop Excel
	@echo   make normalize          Normalize design in both canonical workbooks
	@echo   make preview            Optionally render PNG previews through Windows Excel
	@echo   make preview-public     Render sanitized README screenshots through Windows Excel
	@echo   make verify-static      Run static contracts without rebuilding
	@echo   make verify             Run the portable build and static gate
	@echo   make verify-release     Build and verify clean public release artifacts
	@echo   make verify-excel       Run Windows desktop Excel acceptance
	@echo   make verify-skip-build  Run all checks against the current outputs

env-check:
	$(NPM) run env:check

build:
	$(NPM) run build

build-xlsx:
	$(NPM) run build:xlsx

build-xlsm:
	$(NPM) run build:xlsm

build-release:
	$(NPM) run build:release

sync-vba:
	$(NPM) run sync:vba

normalize:
	$(NPM) run normalize

preview:
	$(NPM) run preview

preview-public:
	$(NPM) run preview:public

verify-static:
	$(NPM) run verify:static

verify:
	$(NPM) run verify

verify-release:
	$(NPM) run verify:release

verify-excel:
	$(NPM) run verify:excel

verify-skip-build:
	$(NPM) run verify:static
	$(NPM) run verify:excel
