param(
    [switch]$StaticOnly,
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"

function Write-Pass([string]$Name) {
    Write-Host "PASS $Name"
}

function Write-Fail([string]$Name, [string]$Message) {
    Write-Host "FAIL $Name`: $Message"
}

function Invoke-Step([string]$Name, [scriptblock]$Block) {
    try {
        & $Block
        Write-Pass $Name
    } catch {
        Write-Fail $Name $_.Exception.Message
        exit 1
    }
}

function Assert-NoWorkbookLocks {
    $locks = Get-ChildItem -LiteralPath "outputs" -Force -Filter '~$quarter_planning_step*.xls*' -ErrorAction SilentlyContinue
    if ($locks) {
        $excelProcesses = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
        if ($excelProcesses.Count -eq 0) {
            Write-Host "PASS workbook lock check (stale ignored)"
            return
        }
        $paths = ($locks | ForEach-Object { $_.FullName }) -join "; "
        throw "Workbook lock files are present. Close Excel before verification: $paths"
    }
}

function Assert-GitCleanAfterBuild {
    $gitTop = git rev-parse --show-toplevel 2>$null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "PASS git tracked-file drift skipped (no git repository)"
        return
    }

    $changes = git diff --name-only
    if ($LASTEXITCODE -ne 0) {
        throw "git diff failed"
    }
    $newChanges = @($changes | Where-Object { $script:InitialUnstagedTracked -notcontains $_ })
    if ($newChanges) {
        $generatedArtifacts = @(
            "assets/vba/quarter_planning_macro_template.xlsm",
            "assets/vba/vbaProject.step2.bin",
            "outputs/quarter_planning_step1.xlsx",
            "outputs/quarter_planning_step2.xlsm"
        )
        $unexpectedChanges = @($newChanges | Where-Object { $generatedArtifacts -notcontains $_ })
        if ($unexpectedChanges) {
            throw "Build/verification left new unstaged source or contract changes. Review and stage intentional changes before commit.`n$($unexpectedChanges -join "`n")"
        }

        git add -- $newChanges
        if ($LASTEXITCODE -ne 0) {
            throw "failed to stage generated artifacts"
        }
        Write-Host "PASS git tracked-file drift (generated artifacts staged)"
        return
    }
    Write-Pass "git tracked-file drift"
}

Set-Location (Resolve-Path (Join-Path $PSScriptRoot ".."))

$script:InitialUnstagedTracked = @()
git rev-parse --show-toplevel 2>$null | Out-Null
if ($LASTEXITCODE -eq 0) {
    $script:InitialUnstagedTracked = @(git diff --name-only)
}

Invoke-Step "workbook lock check" { Assert-NoWorkbookLocks }

$node = "node"

if (!$StaticOnly -and !$SkipBuild) {
    Invoke-Step "build step1 xlsx" {
        & $node ".\scripts\build_quarter_planning_step1.mjs"
        if ($LASTEXITCODE -ne 0) { throw "build_quarter_planning_step1.mjs failed" }
    }

    Invoke-Step "normalize workbook design after xlsx build" {
        & $node ".\scripts\normalize_workbook_design.mjs" ".\outputs\quarter_planning_step1.xlsx" ".\outputs\quarter_planning_step2.xlsm"
        if ($LASTEXITCODE -ne 0) { throw "normalize_workbook_design.mjs failed" }
    }

    Invoke-Step "sync VBA from source" {
        powershell -NoProfile -ExecutionPolicy Bypass -File ".\scripts\sync_vba_from_source.ps1"
        if ($LASTEXITCODE -ne 0) { throw "sync_vba_from_source.ps1 failed" }
    }

    Invoke-Step "build step2 xlsm" {
        & $node ".\scripts\create_quarter_planning_xlsm.mjs"
        if ($LASTEXITCODE -ne 0) { throw "create_quarter_planning_xlsm.mjs failed" }
    }

    Invoke-Step "normalize workbook design after xlsm build" {
        & $node ".\scripts\normalize_workbook_design.mjs" ".\outputs\quarter_planning_step2.xlsm"
        if ($LASTEXITCODE -ne 0) { throw "normalize_workbook_design.mjs failed" }
    }
}

Invoke-Step "static workbook contracts" {
    & $node ".\scripts\verify_static_contracts.mjs"
    if ($LASTEXITCODE -ne 0) { throw "verify_static_contracts.mjs failed" }
}

if (!$StaticOnly) {
    Invoke-Step "Excel COM acceptance" {
        powershell -NoProfile -ExecutionPolicy Bypass -File ".\scripts\verify_excel_com.ps1"
        if ($LASTEXITCODE -ne 0) { throw "verify_excel_com.ps1 failed" }
    }

    Invoke-Step "git tracked-file drift" { Assert-GitCleanAfterBuild }
}
