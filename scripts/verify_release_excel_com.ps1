param(
    [string]$XlsxPath = "dist\QuarterPlan-Excel-v1.0.0-no-macros.xlsx",
    [string]$XlsmPath = "dist\QuarterPlan-Excel-v1.0.0.xlsm"
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms

$nodeExecutable = if ($env:QUARTER_PLANNING_NODE_EXECUTABLE) { $env:QUARTER_PLANNING_NODE_EXECUTABLE } else { "node" }
$layoutJson = & $nodeExecutable (Join-Path $PSScriptRoot "print_workbook_layout.mjs")
if ($LASTEXITCODE -ne 0) { throw "Could not resolve workbook layout" }
$workbookLayout = $layoutJson | ConvertFrom-Json

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

function Get-ClipboardPlainText {
    $lastError = $null
    for ($attempt = 1; $attempt -le 10; $attempt++) {
        try {
            return [System.Windows.Forms.Clipboard]::GetText()
        } catch {
            $lastError = $_.Exception
            Start-Sleep -Milliseconds 100
        }
    }
    throw "Could not read the Windows plain-text clipboard: $($lastError.Message)"
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
    foreach ($shapeName in @("btnTaskClearAll", "btnTaskRefresh", "btnTaskExportXlsx", "btnTaskImportXlsx", "btnTaskImportCsv", "btnJql03")) {
        if ([string]::IsNullOrWhiteSpace([string]$estimateSheet.Shapes.Item($shapeName).OnAction)) { throw "Release shape $shapeName has empty OnAction" }
    }
    foreach ($shapeName in @("btnQuarterPlanReloadBacklog", "btnQuarterPlanRecalculate", "btnQuarterPlanExportBacklog", "btnQuarterPlanAddAllBacklog", "btnJqlA", "btnJqlG", "btnJqlB")) {
        if ([string]::IsNullOrWhiteSpace([string]$planSheet.Shapes.Item($shapeName).OnAction)) { throw "Release shape $shapeName has empty OnAction" }
    }
    $activeTable = $planSheet.ListObjects("tblPlanActive")
    $activeTickets = $activeTable.ListColumns("ЗНИ/Jira").DataBodyRange
    $activeTickets.ClearContents()
    $activeTickets.Cells(1, 1).Value2 = "REL-1"
    $activeTickets.Cells(2, 1).Value2 = "REL-2"
    $referencesSheet = $workbook.Worksheets.Item("99_Правила планирования")
    $jqlClipboardCell = $referencesSheet.Range([string]$workbookLayout.references.jqlClipboardCell)
    $jqlClipboardCell.EntireColumn.Hidden = $true
    $workbook.Saved = $true
    $excel.Run("'" + $workbook.Name + "'!ThisWorkbook.CopyJqlForTableNameForTest", "tblPlanActive")
    Start-Sleep -Milliseconds 300
    $plainClipboardText = (Get-ClipboardPlainText) -replace "(`r`n|`n|`r)+$", ""
    if ($plainClipboardText -cne "issuekey in (REL-1, REL-2)") { throw "Release JQL clipboard is '$plainClipboardText'" }
    if ([bool]$jqlClipboardCell.EntireColumn.Hidden) { throw "Release JQL staging column remained hidden" }
    if (![bool]$workbook.Saved) { throw "Release JQL copy changed ThisWorkbook.Saved" }
    Assert-NoFormulaErrors $workbook "XLSM"
    Write-Host "PASS clean release XLSX/XLSM Excel COM smoke and JQL clipboard"
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
