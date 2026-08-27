param(
    [string]$WorkbookPath = "outputs\quarter_planning_step2.xlsm",
    [string]$OutputDirectory = "docs\images"
)

$ErrorActionPreference = "Stop"

function Resolve-ProjectPath([string]$Path) {
    $resolvedPath = $Path
    if (![System.IO.Path]::IsPathRooted($resolvedPath)) {
        $resolvedPath = Join-Path $repoRoot $resolvedPath
    }
    return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($resolvedPath)
}

function Export-RangePng($Sheet, [string]$RangeAddress, [string]$FileName) {
    $range = $null
    $chartObject = $null
    try {
        $Sheet.Activate() | Out-Null
        $range = $Sheet.Range($RangeAddress)
        $range.CopyPicture(1, 2) | Out-Null
        Start-Sleep -Milliseconds 200
        $width = [Math]::Max(640, [Math]::Min(2400, [double]$range.Width))
        $height = [Math]::Max(260, [Math]::Min(1800, [double]$range.Height))
        $chartObject = $Sheet.ChartObjects().Add(0, 0, $width, $height)
        $chartObject.Activate() | Out-Null
        $chartObject.Chart.Paste() | Out-Null
        Start-Sleep -Milliseconds 200
        $outputPath = Join-Path $fullOutputDirectory $FileName
        if (!$chartObject.Chart.Export($outputPath, "PNG")) {
            throw "Excel could not export $($Sheet.Name)!$RangeAddress"
        }
        Write-Host "PUBLIC SCREENSHOT $outputPath"
    } finally {
        if ($chartObject -ne $null) { $chartObject.Delete() | Out-Null }
        if ($range -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($range) | Out-Null }
    }
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$sourceWorkbookPath = Resolve-ProjectPath $WorkbookPath
$fullOutputDirectory = Resolve-ProjectPath $OutputDirectory
$tempDirectory = Resolve-ProjectPath "outputs\public_screenshot_temp"
$tempWorkbookPath = Join-Path $tempDirectory "quarterplan-public-screenshot.xlsm"

if (!(Test-Path -LiteralPath $sourceWorkbookPath)) { throw "Workbook not found: $sourceWorkbookPath" }
New-Item -ItemType Directory -Path $fullOutputDirectory -Force | Out-Null
New-Item -ItemType Directory -Path $tempDirectory -Force | Out-Null
Copy-Item -LiteralPath $sourceWorkbookPath -Destination $tempWorkbookPath -Force

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $true
    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Open($tempWorkbookPath, 0, $false)

    $reloadMacro = "'" + $workbook.Name + "'!ThisWorkbook.ReloadQuarterPlanBacklogFromEstimates"
    $excel.Run($reloadMacro, $false)
    $excel.Run("RunQuarterPlanAddAllBacklogToPlan")
    $excel.Run("RunQuarterPlanRecalculate")
    $excel.CalculateFull()

    Export-RangePng $workbook.Worksheets.Item("00_Настройки") "A1:L34" "settings.png"
    Export-RangePng $workbook.Worksheets.Item("03_Оценка задач") "A1:N14" "estimates.png"
    Export-RangePng $workbook.Worksheets.Item("04_Квартальный план") "A1:S16" "plan.png"
} finally {
    if ($workbook -ne $null) {
        $workbook.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel -ne $null) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if (Test-Path -LiteralPath $tempWorkbookPath) { Remove-Item -LiteralPath $tempWorkbookPath -Force }
}
