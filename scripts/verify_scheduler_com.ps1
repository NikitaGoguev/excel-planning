param([string]$WorkbookPath = "outputs\quarter_planning_step2.xlsm")

$ErrorActionPreference = "Stop"
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$sourcePath = Join-Path $repoRoot "tests\vba\QuarterPlanSchedulerTests_module.txt"
$sourceWorkbook = if ([System.IO.Path]::IsPathRooted($WorkbookPath)) { $WorkbookPath } else { Join-Path $repoRoot $WorkbookPath }
$temporaryDirectory = Join-Path ([System.IO.Path]::GetTempPath()) ("quarterplan-scheduler-" + [Guid]::NewGuid().ToString("N"))
$temporaryWorkbook = Join-Path $temporaryDirectory "scheduler-tests.xlsm"
$excel = $null
$workbook = $null
$component = $null
$vbProject = $null
$vbComponents = $null
$codeModule = $null
$existingExcelProcessIds = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })

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
    $component = $vbComponents.Add(1)
    $component.Name = "QuarterPlanSchedulerTests"
    $codeModule = $component.CodeModule
    $codeModule.AddFromString((Get-Content -LiteralPath $sourcePath -Raw -Encoding UTF8))
    $scenarioCount = [int]$excel.Run("'" + $workbook.Name + "'!RunQuarterPlanSchedulerUnitTests")
    if ($scenarioCount -lt 32) { throw "Only $scenarioCount scheduler scenarios ran" }
    Write-Host "PASS scheduler business scenarios ($scenarioCount)"
} finally {
    if ($codeModule -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($codeModule) | Out-Null; $codeModule = $null }
    if ($component -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($component) | Out-Null; $component = $null }
    if ($vbComponents -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($vbComponents) | Out-Null; $vbComponents = $null }
    if ($vbProject -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($vbProject) | Out-Null; $vbProject = $null }
    if ($workbook -ne $null) { $workbook.Close($false); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null }
    if ($excel -ne $null) { $excel.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    for ($attempt = 0; $attempt -lt 50; $attempt++) {
        $automationProcesses = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $existingExcelProcessIds -notcontains $_.Id })
        if ($automationProcesses.Count -eq 0) { break }
        Start-Sleep -Milliseconds 100
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
