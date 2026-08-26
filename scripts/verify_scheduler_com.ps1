param(
    [string]$WorkbookPath = "outputs\quarter_planning_step2.xlsm",
    [string]$ContractPath = "contracts\vba.contract.json"
)

$ErrorActionPreference = "Stop"
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$contractFullPath = if ([System.IO.Path]::IsPathRooted($ContractPath)) { $ContractPath } else { Join-Path $repoRoot $ContractPath }
if (!(Test-Path -LiteralPath $contractFullPath)) { throw "VBA contract not found: $contractFullPath" }
$contract = Get-Content -LiteralPath $contractFullPath -Raw -Encoding UTF8 | ConvertFrom-Json
$schedulerContract = $contract.scheduler
$schedulerComponentContract = @($contract.components | Where-Object { [string]$_.name -eq [string]$schedulerContract.component })
if ($schedulerComponentContract.Count -ne 1) { throw "Scheduler component contract is missing or duplicated: $($schedulerContract.component)" }
$schedulerSourcePath = Join-Path $repoRoot ([string]$schedulerComponentContract[0].source)
$testSourcePath = Join-Path $repoRoot ([string]$schedulerContract.testSource)
if (!(Test-Path -LiteralPath $schedulerSourcePath)) { throw "Scheduler source not found: $schedulerSourcePath" }
if (!(Test-Path -LiteralPath $testSourcePath)) { throw "Scheduler test source not found: $testSourcePath" }
$expectedScenarioCount = [int]$schedulerContract.scenarioCount
if ($expectedScenarioCount -le 0) { throw "Scheduler scenarioCount must be positive" }
$sourceWorkbook = if ([System.IO.Path]::IsPathRooted($WorkbookPath)) { $WorkbookPath } else { Join-Path $repoRoot $WorkbookPath }
if (!(Test-Path -LiteralPath $sourceWorkbook)) { throw "Workbook not found: $sourceWorkbook" }
$temporaryDirectory = Join-Path ([System.IO.Path]::GetTempPath()) ("quarterplan-scheduler-" + [Guid]::NewGuid().ToString("N"))
$temporaryWorkbook = Join-Path $temporaryDirectory "scheduler-tests.xlsm"
$excel = $null
$workbook = $null
$schedulerComponent = $null
$schedulerCodeModule = $null
$testComponent = $null
$testCodeModule = $null
$vbProject = $null
$vbComponents = $null
$existingExcelProcessIds = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })

function Normalize-VbaSource([string]$Text) {
    return (($Text -replace "`r`n", "`n") -replace "`r", "`n").Trim()
}

try {
    New-Item -ItemType Directory -Path $temporaryDirectory | Out-Null
    Copy-Item -LiteralPath $sourceWorkbook -Destination $temporaryWorkbook
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Open($temporaryWorkbook, 0, $false)
    $vbProject = $workbook.VBProject
    $vbComponents = $vbProject.VBComponents

    $schedulerComponent = $vbComponents.Item([string]$schedulerContract.component)
    $schedulerCodeModule = $schedulerComponent.CodeModule
    if ($schedulerCodeModule.CountOfLines -gt 0) { $schedulerCodeModule.DeleteLines(1, $schedulerCodeModule.CountOfLines) }
    $schedulerSource = Get-Content -LiteralPath $schedulerSourcePath -Raw -Encoding UTF8
    $schedulerCodeModule.AddFromString($schedulerSource)
    $importedSchedulerSource = $schedulerCodeModule.Lines(1, $schedulerCodeModule.CountOfLines)
    if ((Normalize-VbaSource $importedSchedulerSource) -cne (Normalize-VbaSource $schedulerSource)) {
        throw "Imported scheduler source does not match $schedulerSourcePath"
    }

    $testComponent = $vbComponents.Add(1)
    $testComponent.Name = "QuarterPlanSchedulerTests"
    $testCodeModule = $testComponent.CodeModule
    $testCodeModule.AddFromString((Get-Content -LiteralPath $testSourcePath -Raw -Encoding UTF8))
    $testResult = [string]$excel.Run("'" + $workbook.Name + "'!RunQuarterPlanSchedulerUnitTests")
    $testParts = $testResult.Split('|', 3)
    if ($testParts.Count -lt 2) { throw "Scheduler tests returned an invalid result: $testResult" }
    if ($testParts[0] -eq "FAIL") {
        $failureMessage = if ($testParts.Count -ge 3) { $testParts[2] } else { "Scenario $($testParts[1]) failed without diagnostics" }
        throw $failureMessage
    }
    if ($testParts[0] -ne "PASS") { throw "Scheduler tests returned an invalid status: $testResult" }
    $scenarioCount = [int]$testParts[1]
    if ($scenarioCount -ne $expectedScenarioCount) { throw "Expected $expectedScenarioCount scheduler scenarios, received $scenarioCount" }
    Write-Host "PASS scheduler business scenarios ($scenarioCount)"
} finally {
    if ($testCodeModule -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($testCodeModule) | Out-Null; $testCodeModule = $null }
    if ($testComponent -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($testComponent) | Out-Null; $testComponent = $null }
    if ($schedulerCodeModule -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($schedulerCodeModule) | Out-Null; $schedulerCodeModule = $null }
    if ($schedulerComponent -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($schedulerComponent) | Out-Null; $schedulerComponent = $null }
    if ($vbComponents -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($vbComponents) | Out-Null; $vbComponents = $null }
    if ($vbProject -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($vbProject) | Out-Null; $vbProject = $null }
    if ($workbook -ne $null) {
        try { $workbook.Close($false) } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel -ne $null) {
        try { $excel.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    for ($attempt = 0; $attempt -lt 50; $attempt++) {
        $automationProcesses = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $existingExcelProcessIds -notcontains $_.Id })
        if ($automationProcesses.Count -eq 0) { break }
        Start-Sleep -Milliseconds 100
    }
    $automationProcesses = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $existingExcelProcessIds -notcontains $_.Id })
    foreach ($process in $automationProcesses) {
        if ($process.MainWindowHandle -eq 0) { Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue }
    }
    if (Test-Path -LiteralPath $temporaryDirectory) {
        $resolvedTempRoot = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath()).TrimEnd('\') + '\'
        $resolvedTarget = [System.IO.Path]::GetFullPath($temporaryDirectory)
        if (!$resolvedTarget.StartsWith($resolvedTempRoot, [System.StringComparison]::OrdinalIgnoreCase) -or !(Split-Path -Leaf $resolvedTarget).StartsWith("quarterplan-scheduler-")) {
            throw "Refusing to remove unexpected scheduler temporary directory: $resolvedTarget"
        }
        Remove-Item -LiteralPath $resolvedTarget -Recurse -Force
    }
}
