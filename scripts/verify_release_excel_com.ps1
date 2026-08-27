param(
    [string]$XlsxPath = "dist\QuarterPlan-Excel-v1.0.0-no-macros.xlsx",
    [string]$XlsmPath = "dist\QuarterPlan-Excel-v1.0.0.xlsm"
)

$ErrorActionPreference = "Stop"

function Resolve-ProjectPath([string]$Path) {
    $resolvedPath = $Path
    if (![System.IO.Path]::IsPathRooted($resolvedPath)) { $resolvedPath = Join-Path (Get-Location).Path $resolvedPath }
    return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($resolvedPath)
}

function Assert-Blank($Cell, [string]$Context) {
    if (![string]::IsNullOrWhiteSpace([string]$Cell.Text)) { throw "$Context is '$($Cell.Text)', expected blank" }
}

function Assert-Text($Cell, [string]$Expected, [string]$Context) {
    if ([string]$Cell.Text -cne $Expected) { throw "$Context is '$($Cell.Text)', expected '$Expected'" }
}

function Assert-ReleaseCommentOptions($Workbook, [string]$Context) {
    $lists = $Workbook.Worksheets.Item("101_Списки")
    $artifacts = $lists.ListObjects("tblTaskCommentArtifacts")
    $adjacent = $lists.ListObjects("tblTaskCommentAdjacentTeams")
    $expectedArtifacts = @("Бизнес-требования", "Макеты интерфейсов", "Архитектурное решение", "Проверка ИБ")
    $expectedAdjacent = @("Команда Альфа", "Команда Бета", "Команда Гамма", "Команда Дельта", "Команда Эпсилон", "Команда Дзета", "Платформа 1", "Интеграции", "Поддержка")
    for ($index = 1; $index -le $expectedArtifacts.Count; $index++) {
        Assert-Text $artifacts.DataBodyRange.Cells($index, 1) ([string]$expectedArtifacts[$index - 1]) "$Context artifact option $index"
    }
    for ($index = 1; $index -le $expectedAdjacent.Count; $index++) {
        Assert-Text $adjacent.DataBodyRange.Cells($index, 1) ([string]$expectedAdjacent[$index - 1]) "$Context adjacent option $index"
    }
}

function Assert-NoFormulaErrors($Workbook, [string]$Context) {
    foreach ($sheet in @($Workbook.Worksheets)) {
        $errorCells = $null
        try {
            $errorCells = $sheet.UsedRange.SpecialCells(-4123, 16)
        } catch {
            $errorCells = $null
        }
        if ($errorCells -ne $null) {
            $address = $errorCells.Address($false, $false)
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($errorCells) | Out-Null
            throw "$Context formula error on $($sheet.Name)!$address"
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null
    }
}

$xlsxFullPath = Resolve-ProjectPath $XlsxPath
$xlsmFullPath = Resolve-ProjectPath $XlsmPath
foreach ($path in @($xlsxFullPath, $xlsmFullPath)) {
    if (!(Test-Path -LiteralPath $path -PathType Leaf)) { throw "Release workbook not found: $path" }
}

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $true

    $excel.AutomationSecurity = 3
    $workbook = $excel.Workbooks.Open($xlsxFullPath, 0, $true)
    if ($workbook.HasVBProject) { throw "Release XLSX unexpectedly has a VBA project" }
    $settings = $workbook.Worksheets.Item("00_Настройки")
    $estimates = $workbook.Worksheets.Item("03_Оценка задач").ListObjects("tblTaskEstimates")
    Assert-Blank $settings.Range("B4") "XLSX team"
    Assert-Blank $settings.Range("B5") "XLSX project lead"
    Assert-Blank $estimates.DataBodyRange.Cells(1, 3) "XLSX first task description"
    Assert-ReleaseCommentOptions $workbook "XLSX"
    Assert-NoFormulaErrors $workbook "XLSX"
    $workbook.Close($false)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    $workbook = $null

    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Open($xlsmFullPath, 0, $true)
    if (!$workbook.HasVBProject) { throw "Release XLSM has no VBA project" }
    $components = $workbook.VBProject.VBComponents
    foreach ($componentName in @("ThisWorkbook", "QuarterPlanActions")) {
        $component = $components.Item($componentName)
        if ($component.CodeModule.CountOfLines -le 0) { throw "Release XLSM component $componentName has no code" }
    }
    if ([int]$components.Item("QuarterPlanTaskCommentForm").Type -ne 3) { throw "Release XLSM task comment UserForm is missing" }
    $settings = $workbook.Worksheets.Item("00_Настройки")
    $estimateSheet = $workbook.Worksheets.Item("03_Оценка задач")
    $planSheet = $workbook.Worksheets.Item("04_Квартальный план")
    Assert-Blank $settings.Range("B4") "XLSM team"
    Assert-Blank $settings.Range("B5") "XLSM project lead"
    Assert-Blank $estimateSheet.ListObjects("tblTaskEstimates").DataBodyRange.Cells(1, 3) "XLSM first task description"
    Assert-Blank $planSheet.ListObjects("tblPlanBacklog").DataBodyRange.Cells(1, 8) "XLSM first backlog description"
    Assert-ReleaseCommentOptions $workbook "XLSM"
    $excel.Run("RunTaskEstimateRepairActionButtons")
    $excel.Run("RunQuarterPlanRepairActionButtons")
    foreach ($shapeName in @("btnTaskClearAll", "btnTaskRefresh", "btnTaskExportXlsx", "btnTaskImportXlsx", "btnTaskImportCsv")) {
        if ([string]::IsNullOrWhiteSpace([string]$estimateSheet.Shapes.Item($shapeName).OnAction)) { throw "Release shape $shapeName has empty OnAction" }
    }
    foreach ($shapeName in @("btnQuarterPlanReloadBacklog", "btnQuarterPlanRecalculate", "btnQuarterPlanExportBacklog", "btnQuarterPlanAddAllBacklog")) {
        if ([string]::IsNullOrWhiteSpace([string]$planSheet.Shapes.Item($shapeName).OnAction)) { throw "Release shape $shapeName has empty OnAction" }
    }
    Assert-NoFormulaErrors $workbook "XLSM"
    Write-Host "PASS clean release XLSX/XLSM Excel COM smoke"
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
}
