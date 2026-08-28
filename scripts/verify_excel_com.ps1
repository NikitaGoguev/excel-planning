param(
    [string]$WorkbookPath = "outputs\quarter_planning_step2.xlsm"
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms

$nodeExecutable = if ($env:QUARTER_PLANNING_NODE_EXECUTABLE) { $env:QUARTER_PLANNING_NODE_EXECUTABLE } else { "node" }
$layoutJson = & $nodeExecutable (Join-Path $PSScriptRoot "print_workbook_layout.mjs")
if ($LASTEXITCODE -ne 0) { throw "Could not resolve workbook layout" }
$workbookLayout = $layoutJson | ConvertFrom-Json

function Write-Pass([string]$Name) {
    Write-Host "PASS $Name"
}

function Write-Fail([string]$Name, [string]$Message) {
    Write-Host "FAIL $Name`: $Message"
}

function Resolve-ProjectPath([string]$Path) {
    $resolvedPath = $Path
    if (![System.IO.Path]::IsPathRooted($resolvedPath)) {
        $resolvedPath = Join-Path (Get-Location).Path $resolvedPath
    }
    return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($resolvedPath)
}

function Normalize-VbaSource([string]$Text) {
    return (($Text -replace "`r`n", "`n") -replace "`r", "`n").Trim()
}

function Assert-PlainMacroName([string]$Value, [string]$ShapeName) {
    if ([string]::IsNullOrWhiteSpace($Value)) {
        throw "Shape $ShapeName has empty OnAction"
    }
    if ($Value -match "[!']|\(|\)") {
        throw "Shape $ShapeName has non-plain OnAction: $Value"
    }
}

function Assert-TextEquals([string]$Actual, [string]$Expected, [string]$Context) {
    if ($Actual -ne $Expected) {
        throw "$Context is '$Actual', expected '$Expected'"
    }
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

function Assert-NumberClose([double]$Actual, [double]$Expected, [string]$Context, [double]$Tolerance = 0.01) {
    if ([Math]::Abs($Actual - $Expected) -gt $Tolerance) {
        throw "$Context is $Actual, expected $Expected"
    }
}

function Test-NumberClose([double]$Actual, [double]$Expected, [double]$Tolerance = 0.01) {
    return [Math]::Abs($Actual - $Expected) -le $Tolerance
}

function Get-ComScalar($Value) {
    while ($Value -is [System.Array]) {
        if ($Value.Count -eq 0) { return $null }
        $Value = $Value[0]
    }
    return $Value
}

function New-Rect($Range) {
    return [PSCustomObject]@{
        Left = [double](Get-ComScalar $Range.Left)
        Top = [double](Get-ComScalar $Range.Top)
        Width = [double](Get-ComScalar $Range.Width)
        Height = [double](Get-ComScalar $Range.Height)
    }
}

function Get-ShapeAnchorRects($Shape) {
    $cell = $Shape.TopLeftCell
    if ([bool]$cell.MergeCells) {
        return @(New-Rect $cell.Worksheet.Range($cell.MergeArea.Address($false, $false)))
    }
    $bottomRightCell = $Shape.BottomRightCell
    $boundingRange = $cell.Worksheet.Range($cell, $bottomRightCell)
    return @((New-Rect $cell), (New-Rect $boundingRange))
}

function Assert-ActionShapeDesign($Shape, [string]$ShapeName) {
    $font = $Shape.TextFrame.Characters().Font
    Assert-TextEquals ([string](Get-ComScalar $font.Name)) "Calibri" "$ShapeName font"
    Assert-NumberClose ([double](Get-ComScalar $font.Size)) 11 "$ShapeName font size" 0.1
    if (![bool](Get-ComScalar $font.Bold)) {
        throw "$ShapeName font is not bold"
    }

    $shapeLeft = [double](Get-ComScalar $Shape.Left)
    $shapeTop = [double](Get-ComScalar $Shape.Top)
    $shapeWidth = [double](Get-ComScalar $Shape.Width)
    $shapeHeight = [double](Get-ComScalar $Shape.Height)
    foreach ($anchorRect in @(Get-ShapeAnchorRects $Shape)) {
        if ((Test-NumberClose $shapeLeft $anchorRect.Left 0.75) -and
            (Test-NumberClose $shapeTop $anchorRect.Top 0.75) -and
            (Test-NumberClose $shapeWidth $anchorRect.Width 0.75) -and
            (Test-NumberClose $shapeHeight $anchorRect.Height 0.75)) {
            return
        }
    }
    throw "$ShapeName geometry does not match its anchor cell range"
}

function Assert-CellBlank($Cell, [string]$Context) {
    if (![string]::IsNullOrWhiteSpace([string]$Cell.Text)) {
        throw "$Context is '$($Cell.Text)', expected blank"
    }
}

function Assert-CellNotBlank($Cell, [string]$Context) {
    if ([string]::IsNullOrWhiteSpace([string]$Cell.Text)) {
        throw "$Context is blank"
    }
}

function Assert-Strikethrough($Cell, [bool]$Expected, [string]$Context) {
    $actual = [bool]$Cell.Font.Strikethrough
    if ($actual -ne $Expected) {
        throw "$Context strikethrough is $actual, expected $Expected"
    }
}

function Assert-SameExcelDate($ActualCell, $ExpectedCell, [string]$Context) {
    if ([string]::IsNullOrWhiteSpace([string]$ExpectedCell.Text)) {
        throw "$Context expected comparison cell is blank"
    }
    Assert-NumberClose ([double]$ActualCell.Value2) ([double]$ExpectedCell.Value2) $Context 0.0001
}

function Assert-ExcelDateEquals($ActualCell, [datetime]$ExpectedDate, [string]$Context) {
    if ([string]::IsNullOrWhiteSpace([string]$ActualCell.Text)) {
        throw "$Context is blank, expected $($ExpectedDate.ToString("yyyy-MM-dd"))"
    }
    Assert-NumberClose ([double]$ActualCell.Value2) ([double]$ExpectedDate.ToOADate()) $Context 0.0001
}

function Assert-LocalDateFormat($Cell, [string]$Context) {
    $actualFormat = ([string]$Cell.NumberFormatLocal).ToUpperInvariant()
    if ($actualFormat -ne "ДД-ММ-ГГГГ") {
        throw "$Context format is '$actualFormat', expected 'ДД-ММ-ГГГГ'"
    }
}

$workbookFullPath = Resolve-ProjectPath $WorkbookPath
if (!(Test-Path -LiteralPath $workbookFullPath)) {
    Write-Fail "Excel COM prerequisites" "Workbook not found: $workbookFullPath"
    exit 1
}
$vbaContractPath = Resolve-ProjectPath "contracts\vba.contract.json"
$vbaContract = Get-Content -LiteralPath $vbaContractPath -Raw -Encoding UTF8 | ConvertFrom-Json
$schedulerContract = $vbaContract.scheduler
$schedulerComponentContract = @($vbaContract.components | Where-Object { [string]$_.name -eq [string]$schedulerContract.component })
if ($schedulerComponentContract.Count -ne 1) { throw "Scheduler component contract is missing or duplicated: $($schedulerContract.component)" }
$schedulerSourcePath = Resolve-ProjectPath ([string]$schedulerComponentContract[0].source)

$securityKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
$accessVbom = $null
if (Test-Path $securityKey) {
    $accessVbom = (Get-ItemProperty -Path $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
}
if ($accessVbom -ne 1) {
    Write-Fail "Excel COM prerequisites" "AccessVBOM must be enabled at $securityKey"
    exit 1
}
Write-Pass "Excel COM prerequisites"

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $true
    $excel.AutomationSecurity = 1

    $workbook = $excel.Workbooks.Open($workbookFullPath, 0, $false)
    Write-Pass "Excel COM opens xlsm"

    if (!$workbook.HasVBProject) {
        throw "Workbook.HasVBProject is false"
    }
    $components = $workbook.VBProject.VBComponents
    $null = $components.Count
    Write-Pass "VBProject is readable"

    $commentFormComponent = $components.Item("QuarterPlanTaskCommentForm")
    if ([int]$commentFormComponent.Type -ne 3) {
        throw "QuarterPlanTaskCommentForm component type is $([int]$commentFormComponent.Type), expected 3"
    }
    Write-Pass "task comment UserForm component"

    $schedulerComponent = $components.Item([string]$schedulerContract.component)
    $schedulerModule = $schedulerComponent.CodeModule
    $embeddedSchedulerSource = $schedulerModule.Lines(1, $schedulerModule.CountOfLines)
    $expectedSchedulerSource = Get-Content -LiteralPath $schedulerSourcePath -Raw -Encoding UTF8
    if ((Normalize-VbaSource $embeddedSchedulerSource) -cne (Normalize-VbaSource $expectedSchedulerSource)) {
        throw "Embedded $($schedulerContract.component) is stale; run npm run sync:vba and rebuild the XLSM"
    }
    Write-Pass "embedded scheduler source parity"

    $requiredMacros = @(
        "RunQuarterPlanReloadBacklog",
        "RunQuarterPlanRecalculate",
        "RunQuarterPlanCellAction",
        "RunQuarterPlanRepairActionButtons",
        "RunCopyJql",
        "RunTaskEstimateRepairActionButtons"
    )
    $allCode = ""
    for ($i = 1; $i -le $components.Count; $i++) {
        $module = $components.Item($i).CodeModule
        if ($module.CountOfLines -gt 0) {
            $allCode += "`n" + $module.Lines(1, $module.CountOfLines)
        }
    }
    foreach ($macro in $requiredMacros) {
        if ($allCode -notmatch "(?i)\bPublic\s+Sub\s+$macro\b") {
            throw "Required public macro not found in VBProject: $macro"
        }
    }
    Write-Pass "VBProject public macros"

    $planSheet = $null
    $estimateSheet = $null
    $settingsSheet = $null
    $quarterSheet = $null
    $capacitySheet = $workbook.Worksheets.Item("02_Capacity")
    $listsSheet = $workbook.Worksheets.Item("101_Списки")
    foreach ($sheet in @($workbook.Worksheets)) {
        foreach ($table in @($sheet.ListObjects)) {
            if ($table.Name -eq "tblPlanBacklog") { $planSheet = $sheet }
            if ($table.Name -eq "tblTaskEstimates") { $estimateSheet = $sheet }
            if ($table.Name -eq "tblTeamComposition") { $settingsSheet = $sheet }
            if ($table.Name -eq "tblHolidays") { $quarterSheet = $sheet }
        }
    }
    if ($null -eq $planSheet) { throw "Worksheet with tblPlanBacklog not found" }
    if ($null -eq $estimateSheet) { throw "Worksheet with tblTaskEstimates not found" }
    if ($null -eq $settingsSheet) { throw "Worksheet with tblTeamComposition not found" }
    if ($null -eq $quarterSheet) { throw "Worksheet with tblHolidays not found" }
    Write-Pass "COM table lookup"

    $teamCompositionTable = $settingsSheet.ListObjects("tblTeamComposition")
    if ($teamCompositionTable.Range.Address($false, $false) -ne "A13:B20") {
        throw "tblTeamComposition range is $($teamCompositionTable.Range.Address($false, $false))"
    }
    $teamMembersTable = $settingsSheet.ListObjects("tblTeamMembers")
    if ($teamMembersTable.Range.Address($false, $false) -ne $workbookLayout.teamMembers.tableRange) {
        throw "tblTeamMembers range is $($teamMembersTable.Range.Address($false, $false))"
    }
    $holidayTable = $quarterSheet.ListObjects("tblHolidays")
    if ($holidayTable.Range.Row -ne [int]$workbookLayout.holidays.headerRow -or $holidayTable.Range.Rows.Count -gt ([int]$workbookLayout.limits.holidayRows + 1)) {
        throw "tblHolidays active range is outside configured holiday capacity: $($holidayTable.Range.Address($false, $false))"
    }
    Assert-LocalDateFormat $quarterSheet.Range("B4") "Quarter start date"
    Assert-LocalDateFormat $quarterSheet.Range("B5") "Quarter end date"
    Assert-LocalDateFormat $holidayTable.DataBodyRange.Cells(1, 1) "Holiday date"
    Write-Pass "COM local date formats"
    $estimateTable = $estimateSheet.ListObjects("tblTaskEstimates")
    if ($estimateTable.Range.Address($false, $false) -ne $workbookLayout.taskEstimates.tableRange) {
        throw "tblTaskEstimates range is $($estimateTable.Range.Address($false, $false))"
    }
    $backlogTable = $planSheet.ListObjects("tblPlanBacklog")
    if ($backlogTable.Range.Address($false, $false) -ne $workbookLayout.backlog.tableRange) {
        throw "tblPlanBacklog range is $($backlogTable.Range.Address($false, $false))"
    }
    $activeTableForLayout = $planSheet.ListObjects("tblPlanActive")
    $greyTableForLayout = $planSheet.ListObjects("tblPlanGrey")
    if ($activeTableForLayout.Range.Address($false, $false) -ne $workbookLayout.activePlan.tableRange) {
        throw "tblPlanActive range is $($activeTableForLayout.Range.Address($false, $false))"
    }
    if ($greyTableForLayout.Range.Address($false, $false) -ne $workbookLayout.greyZone.tableRange) {
        throw "tblPlanGrey range is $($greyTableForLayout.Range.Address($false, $false))"
    }
    Assert-NumberClose ([double]$planSheet.Rows.Item([int]$workbookLayout.backlog.dataStartRow + 2).RowHeight) ([double]$planSheet.StandardHeight) "Sheet 04 backlog row height" 0.1
    Write-Pass "COM table ranges"

    $artifactOptionsTable = $listsSheet.ListObjects("tblTaskCommentArtifacts")
    $adjacentOptionsTable = $listsSheet.ListObjects("tblTaskCommentAdjacentTeams")
    $statusesTable = $listsSheet.ListObjects("tblStatuses")
    $yesNoTable = $listsSheet.ListObjects("tblYesNo")
    $expectedArtifactOptions = @("БТ", "Макеты", "ОТАР", "ПСИ ИБ")
    $expectedAdjacentOptions = @("Echo", "Sierra", "Bravo", ("Fox" + "trot"), "Uniform", "India", "Ф1", "ЕСК", "Тесса")
    for ($index = 1; $index -le $expectedArtifactOptions.Count; $index++) {
        Assert-TextEquals ([string]$artifactOptionsTable.DataBodyRange.Cells($index, 1).Text) ([string]$expectedArtifactOptions[$index - 1]) "Artifact option $index"
    }
    for ($index = 1; $index -le $expectedAdjacentOptions.Count; $index++) {
        Assert-TextEquals ([string]$adjacentOptionsTable.DataBodyRange.Cells($index, 1).Text) ([string]$expectedAdjacentOptions[$index - 1]) "Adjacent-team option $index"
    }
    Assert-TextEquals ([string]$statusesTable.DataBodyRange.Cells(1, 1).Text) "Готова аналитика" "First planning status on sheet 101"
    Assert-TextEquals ([string]$statusesTable.DataBodyRange.Cells(7, 1).Text) "Отложено" "Last planning status on sheet 101"
    Assert-TextEquals ([string]$yesNoTable.DataBodyRange.Cells(1, 1).Text) "Да" "Yes value on sheet 101"
    Assert-TextEquals ([string]$yesNoTable.DataBodyRange.Cells(2, 1).Text) "Нет" "No value on sheet 101"
    Write-Pass "COM sheet 101 comment options"

    $referencesSheet = $workbook.Worksheets.Item("99_Правила планирования")
    Assert-TextEquals ([string]$referencesSheet.ListObjects("tblExpertise").DataBodyRange.Cells(2, 1).Text) "DE" "DE expertise code"
    Assert-NumberClose ([double]$referencesSheet.ListObjects("tblExpertise").DataBodyRange.Cells(2, 3).Value2) 1 "DE waterfall order"
    Assert-TextEquals ([string]$referencesSheet.Range("Z1").Address($false, $false)) "Z1" "JQL staging cell on planning rules sheet"
    Write-Pass "COM planning rules references"

    Assert-TextEquals ([string]$capacitySheet.Range("F6").Text) "Авто 00" "Capacity auto header"
    Assert-TextEquals ([string]$capacitySheet.Range("G6").Text) "Переопределение" "Capacity override header"
    if ([int]$capacitySheet.Range("F8").Interior.Color -ne 15987699) {
        throw "Capacity F8 technical fill is $([int]$capacitySheet.Range("F8").Interior.Color), expected 15987699"
    }
    if ([int]$capacitySheet.Range("G8").Font.Color -ne 8421504) {
        throw "Capacity G8 technical font color is $([int]$capacitySheet.Range("G8").Font.Color), expected 8421504"
    }
    Assert-NumberClose ([double]$capacitySheet.Range("F8").Value2) 2 "Initial analyst auto count"
    Assert-NumberClose ([double]$capacitySheet.Range("E8").Value2) 2 "Initial analyst effective count"
    Assert-NumberClose ([double]$capacitySheet.Range("F14").Value2) 0 "Initial analyst auto vacation"
    Assert-NumberClose ([double]$capacitySheet.Range("E14").Value2) 0 "Initial analyst effective vacation"
    Assert-TextEquals ([string]$capacitySheet.Range("C23").Text) "Количество дизайнеров" "Designer count label"
    Assert-TextEquals ([string]$capacitySheet.Range("C24").Text) "Отпуска дизайнеров" "Designer vacation label"
    Assert-TextEquals ([string]$capacitySheet.Range("C25").Text) "Фокус-фактор дизайнеров" "Designer focus factor label"
    if ([int]$capacitySheet.Range("F25").HorizontalAlignment -ne -4131) {
        throw "Designer focus-factor note alignment is $($capacitySheet.Range("F25").HorizontalAlignment), expected left"
    }
    Assert-NumberClose ([double]$capacitySheet.Range("F23").Value2) 1 "Initial designer auto count"
    Assert-NumberClose ([double]$capacitySheet.Range("E23").Value2) 1 "Initial designer effective count"
    Assert-NumberClose ([double]$capacitySheet.Range("F24").Value2) 0 "Initial designer auto vacation"
    Assert-NumberClose ([double]$capacitySheet.Range("E24").Value2) 0 "Initial designer effective vacation"
    Assert-NumberClose ([double]$capacitySheet.Range("E25").Value2) 0.7 "Initial designer focus factor"
    Assert-NumberClose ([double]$capacitySheet.Range("E34").Value2) 43 "Initial designer capacity"
    Assert-NumberClose ([double]$capacitySheet.Range("E35").Value2) 493 "Initial total capacity with designers"
    $capacitySheet.Range("G23").Value2 = 2
    $excel.CalculateFull()
    Assert-NumberClose ([double]$capacitySheet.Range("E23").Value2) 2 "Designer count override"
    Assert-NumberClose ([double]$capacitySheet.Range("E34").Value2) 86 "Designer capacity with count override"
    $capacitySheet.Range("G23").ClearContents()
    $capacitySheet.Range("G24").Value2 = 2
    $excel.CalculateFull()
    Assert-NumberClose ([double]$capacitySheet.Range("E34").Value2) 42 "Designer capacity with vacation override"
    $capacitySheet.Range("G24").ClearContents()
    $excel.CalculateFull()
    $capacitySheet.Range("G8").Value2 = 9
    $excel.CalculateFull()
    Assert-NumberClose ([double]$capacitySheet.Range("F8").Value2) 2 "Analyst auto count with override"
    Assert-NumberClose ([double]$capacitySheet.Range("E8").Value2) 9 "Analyst override effective count"
    $capacitySheet.Range("G8").ClearContents()
    $excel.CalculateFull()
    Assert-NumberClose ([double]$capacitySheet.Range("E8").Value2) 2 "Analyst effective count after override clear"
    Write-Pass "COM capacity team override"

    Assert-TextEquals ([string]$settingsSheet.Range("C3").Text) "Комментарий" "Settings comments header"
    if ($settingsSheet.Range("C4").MergeArea.Address($false, $false) -ne "C4:G4") {
        throw "Settings comment row 4 merge range is $($settingsSheet.Range("C4").MergeArea.Address($false, $false)), expected C4:G4"
    }
    Assert-TextEquals ([string]$settingsSheet.Range("C13").Text) "Комментарий" "Team comments header"
    if ($settingsSheet.Range("C14").MergeArea.Address($false, $false) -ne "C14:G14") {
        throw "Team role comment row 14 merge range is $($settingsSheet.Range("C14").MergeArea.Address($false, $false)), expected C14:G14"
    }
    Assert-TextEquals ([string]$teamCompositionTable.DataBodyRange.Cells(7, 1).Text) "Дизайнер" "Designer composition role"
    if ($settingsSheet.Range("C20").MergeArea.Address($false, $false) -ne "C20:G20") {
        throw "Designer comment merge range is $($settingsSheet.Range("C20").MergeArea.Address($false, $false)), expected C20:G20"
    }
    Assert-TextEquals ([string]$teamMembersTable.DataBodyRange.Cells(12, 1).Text) "Дизайнер" "Initial designer member role"
    Write-Pass "COM sheet 00 comments layout"

    Assert-TextEquals ([string]$teamMembersTable.DataBodyRange.Cells(4, 1).Text) "Аналитик" "Initial second analyst role"
    Assert-TextEquals ([string]$teamMembersTable.DataBodyRange.Cells(4, 2).Text) "Аналитик 2" "Initial second analyst name"
    Assert-NumberClose ([double]$teamMembersTable.DataBodyRange.Cells(4, 3).Value2) 1 "Initial second analyst allocation"
    Assert-CellBlank $teamMembersTable.DataBodyRange.Cells(4, 4) "Initial first vacation start"
    Assert-CellBlank $teamMembersTable.DataBodyRange.Cells(4, 6) "Initial first vacation days"

    $teamCompositionTable.DataBodyRange.Cells(3, 2).Value2 = 3
    $excel.CalculateFull()
    Assert-TextEquals ([string]$teamMembersTable.DataBodyRange.Cells(5, 1).Text) "Аналитик" "Added analyst role"
    Assert-TextEquals ([string]$teamMembersTable.DataBodyRange.Cells(5, 2).Text) "Аналитик 3" "Added analyst default name"
    Assert-NumberClose ([double]$teamMembersTable.DataBodyRange.Cells(5, 3).Value2) 1 "Added analyst allocation"
    foreach ($col in 4..12) {
        Assert-CellBlank $teamMembersTable.DataBodyRange.Cells(5, $col) "Added analyst vacation column $col"
    }

    $teamMembersTable.DataBodyRange.Cells(3, 2).Value2 = "Manual Analyst"
    $teamMembersTable.DataBodyRange.Cells(3, 3).Value2 = 0.75
    $holidayTable.DataBodyRange.Cells(1, 1).Value2 = [datetime]"2026-05-13"
    $holidayTable.DataBodyRange.Cells(1, 2).Value2 = "COM test holiday"
    $holidayTable.DataBodyRange.Cells(1, 4).Value2 = "Да"
    $teamMembersTable.DataBodyRange.Cells(3, 4).Value2 = [datetime]"2026-05-11"
    $teamMembersTable.DataBodyRange.Cells(3, 5).Value2 = [datetime]"2026-05-15"
    $teamCompositionTable.DataBodyRange.Cells(4, 2).Value2 = 3
    $excel.CalculateFull()
    Assert-TextEquals ([string]$teamMembersTable.DataBodyRange.Cells(3, 2).Text) "Manual Analyst" "Manual analyst name preserved"
    Assert-NumberClose ([double]$teamMembersTable.DataBodyRange.Cells(3, 3).Value2) 0.75 "Manual analyst allocation preserved"
    Assert-CellNotBlank $teamMembersTable.DataBodyRange.Cells(3, 4) "Manual analyst vacation start preserved"
    Assert-NumberClose ([double]$teamMembersTable.DataBodyRange.Cells(3, 6).Value2) 4 "Manual analyst vacation day count excludes holiday"
    Assert-NumberClose ([double]$capacitySheet.Range("F8").Value2) 2.75 "Analyst auto count after allocation change"
    Assert-NumberClose ([double]$capacitySheet.Range("E8").Value2) 2.75 "Analyst effective count after allocation change"
    Assert-NumberClose ([double]$capacitySheet.Range("F14").Value2) 3 "Analyst weighted auto vacation"
    Assert-NumberClose ([double]$capacitySheet.Range("E14").Value2) 3 "Analyst effective vacation after allocation change"
    $capacitySheet.Range("G14").Value2 = 11
    $excel.CalculateFull()
    Assert-NumberClose ([double]$capacitySheet.Range("F14").Value2) 3 "Analyst auto vacation with override"
    Assert-NumberClose ([double]$capacitySheet.Range("E14").Value2) 11 "Analyst vacation override effective value"
    Assert-NumberClose ([double]$capacitySheet.Range("E27").Value2) 142 "Analyst capacity uses vacation override"
    $capacitySheet.Range("G14").ClearContents()
    $excel.CalculateFull()
    Assert-NumberClose ([double]$capacitySheet.Range("E14").Value2) 3 "Analyst effective vacation after override clear"
    Assert-NumberClose ([double]$capacitySheet.Range("E27").Value2) 148 "Analyst capacity restored after vacation override clear"

    $teamCompositionTable.DataBodyRange.Cells(3, 2).Value2 = 1
    $teamCompositionTable.DataBodyRange.Cells(4, 2).Value2 = 2
    $excel.CalculateFull()
    Assert-TextEquals ([string]$teamMembersTable.DataBodyRange.Cells(3, 1).Text) "Аналитик" "Reduced analyst remaining role"
    Assert-TextEquals ([string]$teamMembersTable.DataBodyRange.Cells(3, 2).Text) "Manual Analyst" "Reduced analyst remaining manual name"
    Assert-TextEquals ([string]$teamMembersTable.DataBodyRange.Cells(4, 1).Text) "Разработчик бэкенд" "Reduced composition next role"
    foreach ($col in 1..12) {
        Assert-CellBlank $teamMembersTable.DataBodyRange.Cells(12, $col) "Reduced composition cleared trailing row column $col"
    }

    $testerCompositionRow = 6
    $testerVacationFormulaBefore = @{}
    foreach ($col in @(6, 9, 12)) {
        $testerVacationFormulaBefore[$col] = [string]$teamMembersTable.DataBodyRange.Cells(10, $col).FormulaR1C1
        if ([string]::IsNullOrWhiteSpace($testerVacationFormulaBefore[$col])) {
            throw "Tester trailing row vacation formula is blank before composition change in column $col"
        }
    }
    $designerCompositionRow = 7
    $teamCompositionTable.DataBodyRange.Cells($designerCompositionRow, 2).Value2 = 2
    $excel.CalculateFull()
    Assert-TextEquals ([string]$teamMembersTable.DataBodyRange.Cells(12, 1).Text) "Дизайнер" "Added second designer role"
    Assert-TextEquals ([string]$teamMembersTable.DataBodyRange.Cells(12, 2).Text) "Дизайнер 2" "Added second designer default name"
    $designerVacationFormulaBefore = @{}
    foreach ($col in @(6, 9, 12)) {
        $designerVacationFormulaBefore[$col] = [string]$teamMembersTable.DataBodyRange.Cells(12, $col).FormulaR1C1
    }
    $teamCompositionTable.DataBodyRange.Cells($designerCompositionRow, 2).Value2 = 1
    $excel.CalculateFull()
    foreach ($col in @(1, 2, 3, 4, 5, 7, 8, 10, 11)) {
        Assert-CellBlank $teamMembersTable.DataBodyRange.Cells(12, $col) "Reduced designer count cleared editable trailing column $col"
    }
    foreach ($col in @(6, 9, 12)) {
        if (![bool]$teamMembersTable.DataBodyRange.Cells(12, $col).HasFormula) { throw "Reduced designer count removed vacation formula from column $col" }
        Assert-TextEquals ([string]$teamMembersTable.DataBodyRange.Cells(12, $col).FormulaR1C1) ([string]$designerVacationFormulaBefore[$col]) "Reduced designer count preserves vacation formula $col"
    }
    $teamCompositionTable.DataBodyRange.Cells($testerCompositionRow, 2).Value2 = 3
    $excel.CalculateFull()
    Assert-TextEquals ([string]$teamMembersTable.DataBodyRange.Cells(10, 1).Text) "Тестировщик" "Added third tester role"
    Assert-TextEquals ([string]$teamMembersTable.DataBodyRange.Cells(10, 2).Text) "Тестировщик 3" "Added third tester default name"
    $teamCompositionTable.DataBodyRange.Cells($testerCompositionRow, 2).Value2 = 2
    $excel.CalculateFull()
    Assert-TextEquals ([string]$teamMembersTable.DataBodyRange.Cells(10, 1).Text) "Дизайнер" "Reduced tester count shifts designer into the first freed row"
    Assert-TextEquals ([string]$teamMembersTable.DataBodyRange.Cells(10, 2).Text) "Дизайнер 1" "Reduced tester count preserves designer default name"
    foreach ($col in @(6, 9, 12)) {
        if (![bool]$teamMembersTable.DataBodyRange.Cells(10, $col).HasFormula) {
            throw "Reduced tester count removed vacation formula from trailing column $col"
        }
        Assert-TextEquals ([string]$teamMembersTable.DataBodyRange.Cells(10, $col).FormulaR1C1) ([string]$testerVacationFormulaBefore[$col]) "Reduced tester count preserves vacation formula text in column $col"
    }
    Write-Pass "COM sheet 00 team member sync"

    $csvPath = Resolve-ProjectPath "assets\import1.csv"
    $tempDir = Resolve-ProjectPath "outputs\com_acceptance"
    if (Test-Path -LiteralPath $tempDir) {
        Remove-Item -LiteralPath $tempDir -Recurse -Force
    }
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
    $legacyCsvPath = Join-Path $tempDir "task-estimates-legacy.csv"
    $taskExportPath = Join-Path $tempDir "task-estimates-export.xlsx"
    $legacyXlsxPath = Join-Path $tempDir "task-estimates-legacy.xlsx"
    $expressExportPath = Join-Path $tempDir "express-export.xlsx"
    $backlogExportPath = Join-Path $tempDir "quarter-plan-export.xlsx"
    $utf8WithoutBom = New-Object System.Text.UTF8Encoding -ArgumentList $false
    [System.IO.File]::WriteAllText($legacyCsvPath, "Ключ запроса,Тема`r`nPLAN-LEGACY,Legacy CSV import`r`n", $utf8WithoutBom)

    $excel.Run("RunQuarterPlanRepairActionButtons")
    $excel.Run("RunTaskEstimateRepairActionButtons")

    $taskExpectedDescription = "Настройка центра уведомлений"
    $taskExpectedKey = "PLAN-2001"
    $taskExpectedDirection = "Бизнес-эффект (производство)"

    $excel.Run("RunTaskEstimateImportCsvPath", $csvPath, $true)
    Assert-TextEquals ([string]$estimateTable.DataBodyRange.Cells(1, 2).Text) $taskExpectedDirection "Imported task direction"
    Assert-TextEquals ([string]$estimateTable.DataBodyRange.Cells(1, 3).Text) $taskExpectedDescription "Imported task description"
    Assert-TextEquals ([string]$estimateTable.DataBodyRange.Cells(1, 4).Text) $taskExpectedKey "Imported task key"
    Assert-CellBlank $estimateTable.DataBodyRange.Cells(1, 6) "CSV import leaves DE blank"
    Assert-TextEquals ([string]$estimateTable.DataBodyRange.Cells(2, 3).Text) "" "Second task row after replace import"

    $excel.Run("RunTaskEstimateImportCsvPath", $legacyCsvPath, $true)
    Assert-TextEquals ([string]$estimateTable.DataBodyRange.Cells(1, 2).Text) "" "Legacy CSV imported task direction"
    Assert-TextEquals ([string]$estimateTable.DataBodyRange.Cells(1, 3).Text) "Legacy CSV import" "Legacy CSV imported task description"
    Assert-TextEquals ([string]$estimateTable.DataBodyRange.Cells(1, 4).Text) "PLAN-LEGACY" "Legacy CSV imported task key"

    $excel.Run("RunTaskEstimateImportCsvPath", $csvPath, $true)
    Assert-TextEquals ([string]$estimateTable.DataBodyRange.Cells(1, 2).Text) $taskExpectedDirection "Restored imported task direction"
    $estimateTable.DataBodyRange.Cells(1, 1).Value2 = 10
    $estimateTable.DataBodyRange.Cells(1, 2).Value2 = "BE"
    $estimateTable.DataBodyRange.Cells(1, 5).Value2 = "root note"
    $estimateTable.DataBodyRange.Cells(1, 6).Value2 = 1.5
    $estimateTable.DataBodyRange.Cells(1, 7).Value2 = 2.5
    $estimateTable.DataBodyRange.Cells(1, 8).Value2 = 3.5
    $estimateTable.DataBodyRange.Cells(1, 9).Value2 = 4.5
    $estimateTable.DataBodyRange.Cells(1, 10).Value2 = 5.5
    $estimateTable.DataBodyRange.Cells(1, 11).Value2 = "root comment"
    Write-Pass "COM sheet 03 CSV replace import"

    $estimateSheet.Range("C3").Value2 = "broken header"
    $estimateSheet.Range("H1").Value2 = "broken statistics"
    $excel.Run("RunTaskEstimateRefresh")
    $excel.Run("'" + $workbook.Name + "'!ThisWorkbook.AssertTaskEstimateUnicodeTextForTest")
    $excel.Run("'" + $workbook.Name + "'!ThisWorkbook.AssertJqlTextForTest")
    $expectedTaskHeaders = @("Приоритет", "Направление", "Описание", "ЗНИ/Jira", "Примечание", "DE", "AN", "BE", "FE", "QA", "Комментарий")
    for ($col = 1; $col -le $expectedTaskHeaders.Count; $col++) {
        Assert-TextEquals ([string]$estimateTable.HeaderRowRange.Cells(1, $col).Text) ([string]$expectedTaskHeaders[$col - 1]) "Restored sheet 03 header $col"
    }
    Assert-TextEquals ([string]$estimateSheet.Range("H1").Text) "Статистика" "Restored sheet 03 statistics title"
    $excel.Run("RunTaskEstimateRepairActionButtons")

    $expectedTaskActions = @{
        "btnTaskClearAll" = "RunTaskEstimateClearAll"
        "btnTaskRefresh" = "RunTaskEstimateRefresh"
        "btnTaskExportXlsx" = "RunTaskEstimateExportXlsx"
        "btnTaskImportXlsx" = "RunTaskEstimateImportXlsx"
        "btnTaskImportCsv" = "RunTaskEstimateImportCsv"
        "btnJql03" = "RunCopyJql"
        "tea_01_1" = "RunTaskEstimateCellAction"
        "tea_01_15" = "RunTaskEstimateExpressExport"
        "tea_01_14" = "RunTaskEstimateCellAction"
    }
    foreach ($shapeName in $expectedTaskActions.Keys) {
        $shape = $estimateSheet.Shapes.Item($shapeName)
        $onAction = [string]$shape.OnAction
        Assert-PlainMacroName $onAction $shapeName
        if ($onAction -ne $expectedTaskActions[$shapeName]) {
            throw "Shape $shapeName OnAction is $onAction, expected $($expectedTaskActions[$shapeName])"
        }
        Assert-ActionShapeDesign $shape $shapeName
    }
    $expectedTaskCaptions = @{
        "btnTaskClearAll" = "Сбросить"
        "btnTaskRefresh" = "Обновить"
        "btnTaskExportXlsx" = "Экспорт"
        "btnTaskImportXlsx" = "Импорт"
        "btnTaskImportCsv" = "Импорт CSV (до $($workbookLayout.limits.taskRows))"
        "btnJql03" = "JQL"
        "tea_01_1" = ">"
        "tea_01_15" = "Экспорт"
        "tea_01_14" = "+"
    }
    foreach ($shapeName in $expectedTaskCaptions.Keys) {
        $shape = $estimateSheet.Shapes.Item($shapeName)
        $expectedCaption = [string]$expectedTaskCaptions[$shapeName]
        Assert-TextEquals ([string]$shape.TextFrame.Characters().Text) $expectedCaption "$shapeName caption"
        Assert-TextEquals ([string]$shape.TopLeftCell.Text) $expectedCaption "$shapeName anchor cell caption"
    }
    Assert-TextEquals ([string]$estimateSheet.Shapes.Item("btnJql03").TopLeftCell.Address($false, $false)) ([string]$workbookLayout.taskEstimates.jqlActionCell) "btnJql03 button anchor"
    Write-Pass "COM sheet 03 Unicode-safe rebuild and action buttons"

    $excel.Run("'" + $workbook.Name + "'!ThisWorkbook.AssertTaskEstimateCommentOptionsForTest")
    Assert-TextEquals ([string]$estimateTable.DataBodyRange.Cells(1, 11).Text) "root comment" "Comment options test restores root comment"
    Write-Pass "COM sheet 03 comment-options form and append behavior"

    $taskShapeMacro = "'" + $workbook.Name + "'!ThisWorkbook.HandleTaskEstimateShapeAction"
    $excel.Run($taskShapeMacro, "tea_01_1")
    Assert-TextEquals ([string]$estimateSheet.Range("B4").Text) "x" "Parent delete action after decomposition"
    Assert-TextEquals ([string]$estimateTable.DataBodyRange.Cells(2, 3).Text) $taskExpectedDescription "Child task description after decomposition"
    Assert-NumberClose ([double]$estimateTable.DataBodyRange.Cells(2, 6).Value2) 1.5 "Child DE after decomposition"
    Assert-TextEquals ([string]$estimateSheet.Range("B5").Text) "x" "Child delete action after decomposition"
    Assert-TextEquals ([string]$estimateSheet.Range("N4").Text) "+" "Root comment-options action after decomposition"
    Assert-TextEquals ([string]$estimateSheet.Range("N5").Text) "" "Child has no comment-options action"
    try {
        $null = $estimateSheet.Shapes.Item("tea_02_14")
        throw "Child comment-options shape tea_02_14 exists unexpectedly"
    } catch {
        if ($_.Exception.Message -eq "Child comment-options shape tea_02_14 exists unexpectedly") { throw }
    }
    Write-Pass "COM sheet 03 decompose action"

    $excel.Run("RunTaskEstimateExportXlsxPath", $taskExportPath)
    if (!(Test-Path -LiteralPath $taskExportPath)) {
        throw "Task estimate export file was not created: $taskExportPath"
    }
    $exportWorkbook = $excel.Workbooks.Open($taskExportPath, 0, $true)
    try {
        $exportSheet = $exportWorkbook.Worksheets.Item(1)
        Assert-TextEquals ([string]$exportSheet.Name) "Оценка задач" "Task export sheet name"
        for ($col = 1; $col -le $expectedTaskHeaders.Count; $col++) {
            Assert-TextEquals ([string]$exportSheet.Cells(1, $col).Text) ([string]$expectedTaskHeaders[$col - 1]) "Task export header $col"
        }
        Assert-TextEquals ([string]$exportSheet.Cells(2, 3).Text) $taskExpectedDescription "Task export description"
        Assert-TextEquals ([string]$exportSheet.Cells(2, 4).Text) $taskExpectedKey "Task export key"
        Assert-TextEquals ([string]$exportSheet.Cells(2, 5).Text) "root note" "Task export note"
        Assert-TextEquals ([string]$exportSheet.Cells(2, 11).Text) "root comment" "Task export comment"
        Assert-TextEquals ([string]$exportSheet.Cells(3, 3).Text) $taskExpectedDescription "Task export child description"
        if ([int]$exportSheet.Cells(3, 3).IndentLevel -ne 1) {
            throw "Task export child indent is $($exportSheet.Cells(3, 3).IndentLevel), expected 1"
        }
    } finally {
        $exportWorkbook.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($exportWorkbook) | Out-Null
    }
    Write-Pass "COM sheet 03 task export"

    $excel.Run("RunTaskEstimateImportCsvPath", $csvPath, $true)
    $excel.Run("RunTaskEstimateImportXlsxPath", $taskExportPath)
    Assert-TextEquals ([string]$estimateTable.DataBodyRange.Cells(2, 3).Text) $taskExpectedDescription "XLSX import root description"
    Assert-TextEquals ([string]$estimateTable.DataBodyRange.Cells(2, 4).Text) $taskExpectedKey "XLSX import root key"
    Assert-TextEquals ([string]$estimateTable.DataBodyRange.Cells(2, 5).Text) "root note" "XLSX import note"
    Assert-TextEquals ([string]$estimateTable.DataBodyRange.Cells(2, 11).Text) "root comment" "XLSX import comment"
    Assert-TextEquals ([string]$estimateTable.DataBodyRange.Cells(3, 3).Text) $taskExpectedDescription "XLSX import child description"
    Assert-TextEquals ([string]$estimateSheet.Range("B5").Text) "x" "XLSX import root delete action"
    Assert-TextEquals ([string]$estimateSheet.Range("B6").Text) "x" "XLSX import child delete action"
    Assert-NumberClose ([double]$estimateTable.DataBodyRange.Cells(3, 6).Value2) 1.5 "XLSX import child DE"
    Assert-NumberClose ([double]$estimateTable.DataBodyRange.Cells(3, 7).Value2) 2.5 "XLSX import child AN"
    Assert-NumberClose ([double]$estimateTable.DataBodyRange.Cells(3, 8).Value2) 3.5 "XLSX import child BE"
    Assert-NumberClose ([double]$estimateTable.DataBodyRange.Cells(3, 9).Value2) 4.5 "XLSX import child FE"
    Assert-NumberClose ([double]$estimateTable.DataBodyRange.Cells(3, 10).Value2) 5.5 "XLSX import child QA"
    Write-Pass "COM sheet 03 XLSX append import"

    $legacyWorkbook = $excel.Workbooks.Add()
    try {
        $legacySheet = $legacyWorkbook.Worksheets.Item(1)
        $legacyHeaders = @("Приоритет", "Направление", "Описание", "ЗНИ/Jira", "Примечание", "AN", "BE", "FE", "QA", "Комментарий", "Экспорт")
        for ($col = 1; $col -le $legacyHeaders.Count; $col++) { $legacySheet.Cells(1, $col).Value2 = $legacyHeaders[$col - 1] }
        $legacySheet.Cells(2, 1).Value2 = 90
        $legacySheet.Cells(2, 3).Value2 = "Legacy XLSX task"
        $legacySheet.Cells(2, 4).Value2 = "LEGACY-XLSX"
        $legacySheet.Cells(2, 6).Value2 = 7
        $legacySheet.Cells(2, 7).Value2 = 8
        $legacySheet.Cells(2, 8).Value2 = 9
        $legacySheet.Cells(2, 9).Value2 = 10
        $legacySheet.Cells(2, 10).Value2 = "legacy comment"
        $legacyWorkbook.SaveAs($legacyXlsxPath, 51)
    } finally {
        $legacyWorkbook.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($legacyWorkbook) | Out-Null
    }
    $excel.Run("RunTaskEstimateImportXlsxPath", $legacyXlsxPath)
    Assert-TextEquals ([string]$estimateTable.DataBodyRange.Cells(4, 3).Text) "Legacy XLSX task" "Legacy XLSX description"
    Assert-CellBlank $estimateTable.DataBodyRange.Cells(4, 6) "Legacy XLSX leaves DE blank"
    Assert-NumberClose ([double]$estimateTable.DataBodyRange.Cells(4, 7).Value2) 7 "Legacy XLSX maps AN"
    Assert-TextEquals ([string]$estimateTable.DataBodyRange.Cells(4, 11).Text) "legacy comment" "Legacy XLSX maps comment"
    Write-Pass "COM sheet 03 legacy XLSX import"

    $excel.Run("RunTaskEstimateExpressExportPath", 2, $expressExportPath)
    if (!(Test-Path -LiteralPath $expressExportPath)) {
        throw "Express estimate export file was not created: $expressExportPath"
    }
    $expressWorkbook = $excel.Workbooks.Open($expressExportPath, 0, $true)
    try {
        $expressSheet = $expressWorkbook.Worksheets.Item(1)
        Assert-TextEquals ([string]$expressSheet.Range("E2").Text) $taskExpectedKey "Express export key"
        Assert-TextEquals ([string]$expressSheet.Range("E3").Text) $taskExpectedDescription "Express export description"
        Assert-TextEquals ([string]$expressSheet.Range("A17").Text) $taskExpectedDescription "Express export first task"
        Assert-NumberClose ([double]$expressSheet.Range("H17").Value2) 1.5 "Express export maps DE to Design"
    } finally {
        $expressWorkbook.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($expressWorkbook) | Out-Null
    }
    Write-Pass "COM sheet 03 express export"

    $reloadMacro = "'" + $workbook.Name + "'!ThisWorkbook.ReloadQuarterPlanBacklogFromEstimates"
    $excel.Run($reloadMacro, $false)
    if ([string]::IsNullOrWhiteSpace([string]$backlogTable.DataBodyRange.Cells(1, 8).Text)) {
        throw "Backlog reload did not populate the first backlog task description"
    }

    $excel.Run("RunQuarterPlanAddAllBacklogToPlan")
    $activeTable = $planSheet.ListObjects("tblPlanActive")
    if ([string]::IsNullOrWhiteSpace([string]$activeTable.DataBodyRange.Cells(1, 8).Text)) {
        throw "Bulk backlog move did not populate the first active-plan task description"
    }

    $excel.Run("RunQuarterPlanRecalculate")
    Write-Pass "COM backlog load move and recalc"

    $statusValidation = [string]$activeTable.DataBodyRange.Cells(1, 19).Validation.Formula1
    foreach ($status in @("Готова аналитика", "Готова разработка", "Готова разработка (бэк)", "Готова разработка (фронт)", "Готово к релизу", "ПРОМ", "Отложено")) {
        if (!$statusValidation.Contains($status)) {
            throw "Status validation list does not contain '$status': $statusValidation"
        }
    }
    foreach ($oldStatus in @("Не начата", "В работе", "Отложена")) {
        if ($statusValidation.Contains($oldStatus)) {
            throw "Status validation list still contains old status '$oldStatus': $statusValidation"
        }
    }
    Write-Pass "COM sheet 04 status validation"

    function Reset-ActivePlanStatusScenario {
        for ($row = 1; $row -le $activeTable.DataBodyRange.Rows.Count; $row++) {
            for ($col = 6; $col -le 20; $col++) {
                $activeTable.DataBodyRange.Cells($row, $col).ClearContents() | Out-Null
            }
        }
        $activeTable.DataBodyRange.Cells(1, 8).Value2 = "Status scenario row 1"
        $activeTable.DataBodyRange.Cells(1, 11).Value2 = 2
        $activeTable.DataBodyRange.Cells(1, 12).Value2 = 2
        $activeTable.DataBodyRange.Cells(1, 13).Value2 = 3
        $activeTable.DataBodyRange.Cells(1, 14).Value2 = 5
        $activeTable.DataBodyRange.Cells(1, 15).Value2 = 7
    }

    function Assert-StatusScenario([string]$Status, [double]$ExpectedDe, [double]$ExpectedAn, [double]$ExpectedBe, [double]$ExpectedFe, [double]$ExpectedQa, [string]$ReadyKind) {
        Reset-ActivePlanStatusScenario
        if ($Status -ne "<blank>") {
            $activeTable.DataBodyRange.Cells(1, 19).Value2 = $Status
        }
        $excel.Run("RunQuarterPlanRecalculate")
        Assert-NumberClose ([double]$planSheet.Range("K3").Value2) $ExpectedDe "DE planned with risk for status $Status"
        Assert-NumberClose ([double]$planSheet.Range("L3").Value2) $ExpectedAn "AN planned with risk for status $Status"
        Assert-NumberClose ([double]$planSheet.Range("M3").Value2) $ExpectedBe "BE planned with risk for status $Status"
        Assert-NumberClose ([double]$planSheet.Range("N3").Value2) $ExpectedFe "FE planned with risk for status $Status"
        Assert-NumberClose ([double]$planSheet.Range("O3").Value2) $ExpectedQa "QA planned with risk for status $Status"

        Assert-NumberClose ([double]$activeTable.DataBodyRange.Cells(1, 16).Value2) 29 "Full effort with DE for status $Status" 0.0001
        if ($ReadyKind -eq "NONE") {
            Assert-CellBlank $activeTable.DataBodyRange.Cells(1, 17) "Ready date for status $Status"
            foreach ($cell in @("U6", "V6", "W6", "X6", "Y6", "Z6", "AB6", "AD6", "AF6", "AH6")) { Assert-CellBlank $planSheet.Range($cell) "Schedule cell $cell for status $Status" }
            return
        }

        Assert-CellNotBlank $planSheet.Range("U6") "DE duration for status $Status"
        Assert-CellNotBlank $planSheet.Range("Z6") "DE start for status $Status"
        Assert-CellNotBlank $planSheet.Range("AA6") "DE finish for status $Status"
        Assert-CellNotBlank $planSheet.Range("V6") "AN duration for status $Status"
        Assert-CellNotBlank $planSheet.Range("AB6") "AN start for status $Status"
        Assert-CellNotBlank $planSheet.Range("AC6") "AN finish for status $Status"
        Assert-SameExcelDate $planSheet.Range("Z6") $planSheet.Range("AB6") "DE and AN parallel start for status $Status"

        if ($ReadyKind -eq "AN") {
            Assert-CellBlank $planSheet.Range("W6") "BE duration for status $Status"
            Assert-CellBlank $planSheet.Range("X6") "FE duration for status $Status"
            Assert-CellBlank $planSheet.Range("Y6") "QA duration for status $Status"
            Assert-SameExcelDate $activeTable.DataBodyRange.Cells(1, 17) $planSheet.Range("AC6") "Ready date for status $Status"
        } elseif ($ReadyKind -eq "DEV") {
            Assert-CellNotBlank $planSheet.Range("W6") "BE duration for status $Status"
            Assert-CellNotBlank $planSheet.Range("X6") "FE duration for status $Status"
            Assert-CellBlank $planSheet.Range("Y6") "QA duration for status $Status"
            Assert-CellBlank $planSheet.Range("AH6") "QA start for status $Status"
            Assert-CellBlank $planSheet.Range("AI6") "QA finish for status $Status"
            if ([double]$planSheet.Range("AE6").Value2 -ge [double]$planSheet.Range("AG6").Value2) {
                Assert-SameExcelDate $activeTable.DataBodyRange.Cells(1, 17) $planSheet.Range("AE6") "Ready date for status $Status"
            } else {
                Assert-SameExcelDate $activeTable.DataBodyRange.Cells(1, 17) $planSheet.Range("AG6") "Ready date for status $Status"
            }
        } elseif ($ReadyKind -eq "BACK") {
            Assert-CellNotBlank $planSheet.Range("W6") "BE duration for status $Status"
            Assert-CellBlank $planSheet.Range("X6") "FE duration for status $Status"
            Assert-CellBlank $planSheet.Range("Y6") "QA duration for status $Status"
            Assert-CellNotBlank $planSheet.Range("AD6") "BE start for status $Status"
            Assert-CellNotBlank $planSheet.Range("AE6") "BE finish for status $Status"
            Assert-CellBlank $planSheet.Range("AF6") "FE start for status $Status"
            Assert-CellBlank $planSheet.Range("AG6") "FE finish for status $Status"
            Assert-CellBlank $planSheet.Range("AH6") "QA start for status $Status"
            Assert-CellBlank $planSheet.Range("AI6") "QA finish for status $Status"
            Assert-SameExcelDate $activeTable.DataBodyRange.Cells(1, 17) $planSheet.Range("AE6") "Ready date for status $Status"
        } elseif ($ReadyKind -eq "FRONT") {
            Assert-CellBlank $planSheet.Range("W6") "BE duration for status $Status"
            Assert-CellNotBlank $planSheet.Range("X6") "FE duration for status $Status"
            Assert-CellBlank $planSheet.Range("Y6") "QA duration for status $Status"
            Assert-CellBlank $planSheet.Range("AD6") "BE start for status $Status"
            Assert-CellBlank $planSheet.Range("AE6") "BE finish for status $Status"
            Assert-CellNotBlank $planSheet.Range("AF6") "FE start for status $Status"
            Assert-CellNotBlank $planSheet.Range("AG6") "FE finish for status $Status"
            Assert-CellBlank $planSheet.Range("AH6") "QA start for status $Status"
            Assert-CellBlank $planSheet.Range("AI6") "QA finish for status $Status"
            Assert-SameExcelDate $activeTable.DataBodyRange.Cells(1, 17) $planSheet.Range("AG6") "Ready date for status $Status"
        } else {
            Assert-CellNotBlank $planSheet.Range("W6") "BE duration for status $Status"
            Assert-CellNotBlank $planSheet.Range("X6") "FE duration for status $Status"
            Assert-CellNotBlank $planSheet.Range("Y6") "QA duration for status $Status"
            Assert-CellNotBlank $planSheet.Range("AH6") "QA start for status $Status"
            Assert-CellNotBlank $planSheet.Range("AI6") "QA finish for status $Status"
            Assert-SameExcelDate $activeTable.DataBodyRange.Cells(1, 17) $planSheet.Range("AI6") "Ready date for status $Status"
        }
    }

    Assert-StatusScenario "Готова аналитика" 2.4 2.4 0 0 0 "AN"
    Assert-StatusScenario "Готова разработка" 2.4 2.4 3.6 6 0 "DEV"
    Assert-StatusScenario "Готова разработка (бэк)" 2.4 2.4 3.6 0 0 "BACK"
    Assert-StatusScenario "Готова разработка (фронт)" 2.4 2.4 0 6 0 "FRONT"
    Assert-StatusScenario "<blank>" 2.4 2.4 3.6 6 8.4 "FULL"
    Assert-StatusScenario "Готово к релизу" 2.4 2.4 3.6 6 8.4 "FULL"
    Assert-StatusScenario "ПРОМ" 2.4 2.4 3.6 6 8.4 "FULL"
    Assert-StatusScenario "Отложено" 0 0 0 0 0 "NONE"
    Write-Pass "COM sheet 04 status balance and schedule"

    Reset-ActivePlanStatusScenario
    $activeTable.DataBodyRange.Cells(1, 19).Value2 = "ПРОМ"
    $capacitySheet.Range("E25").Value2 = 0.7
    $excel.CalculateFull()
    $excel.Run("RunQuarterPlanRecalculate")
    $deDurationAt70 = [double]$planSheet.Range("U6").Value2
    $anDurationAt70 = [double]$planSheet.Range("V6").Value2
    $readyAt70 = [double]$activeTable.DataBodyRange.Cells(1, 17).Value2
    Assert-NumberClose ([double]$activeTable.DataBodyRange.Cells(1, 16).Value2) 29 "Effort with 70 percent design focus" 0.0001
    $capacitySheet.Range("E25").Value2 = 0.5
    $excel.CalculateFull()
    $excel.Run("RunQuarterPlanRecalculate")
    Assert-NumberClose ([double]$activeTable.DataBodyRange.Cells(1, 16).Value2) 30 "Effort with 50 percent design focus" 0.0001
    Assert-NumberClose ([double]$planSheet.Range("U6").Value2) 4 "DE duration uses separate 50 percent design focus" 0.0001
    Assert-NumberClose ([double]$planSheet.Range("V6").Value2) $anDurationAt70 "AN duration is unchanged by design focus" 0.0001
    Assert-NumberClose ([double]$activeTable.DataBodyRange.Cells(1, 17).Value2) $readyAt70 "Ready date is unchanged by DE duration" 0.0001
    if ([double]$planSheet.Range("U6").Value2 -le $deDurationAt70) { throw "DE duration did not increase after lowering design focus factor" }
    $capacitySheet.Range("E25").Value2 = 0.7
    $excel.CalculateFull()
    Write-Pass "COM sheet 04 separate designer focus factor"

    Reset-ActivePlanStatusScenario
    $activeTable.DataBodyRange.Cells(2, 8).Value2 = "Status scenario row 2"
    $activeTable.DataBodyRange.Cells(2, 13).Value2 = 5
    $activeTable.DataBodyRange.Cells(1, 19).Value2 = "Готова аналитика"
    $excel.Run("RunQuarterPlanRecalculate")
    Assert-CellBlank $planSheet.Range("AD6") "Row 1 BE start for analysis-only status"
    Assert-SameExcelDate $planSheet.Range("AD7") $planSheet.Range("AB6") "Row 2 BE start is not delayed by row 1 analysis-only status"

    Reset-ActivePlanStatusScenario
    $activeTable.DataBodyRange.Cells(2, 8).Value2 = "Status scenario row 2"
    $activeTable.DataBodyRange.Cells(2, 12).Value2 = 2
    $activeTable.DataBodyRange.Cells(1, 19).Value2 = "Отложено"
    $excel.Run("RunQuarterPlanRecalculate")
    Assert-CellBlank $planSheet.Range("AB6") "Row 1 AN start for deferred status"
    Assert-CellBlank $activeTable.DataBodyRange.Cells(1, 17) "Row 1 ready date for deferred status"
    Assert-CellNotBlank $planSheet.Range("AB7") "Row 2 AN start after row 1 deferred status"
    Write-Pass "COM sheet 04 status resource allocation"

    function Clear-TeamMembersForScheduleScenario {
        for ($row = 1; $row -le $teamMembersTable.DataBodyRange.Rows.Count; $row++) {
            foreach ($col in @(1, 2, 3, 4, 5, 7, 8, 10, 11)) {
                $teamMembersTable.DataBodyRange.Cells($row, $col).ClearContents() | Out-Null
            }
        }
    }

    function Set-TeamMemberForScheduleScenario([int]$Row, [string]$Role, [double]$Allocation, [object]$VacationStart = $null, [object]$VacationEnd = $null, [int]$VacationSlot = 1) {
        $teamMembersTable.DataBodyRange.Cells($Row, 1).Value2 = $Role
        $teamMembersTable.DataBodyRange.Cells($Row, 2).Value2 = "$Role $Row"
        $teamMembersTable.DataBodyRange.Cells($Row, 3).Value2 = $Allocation
        if ($null -ne $VacationStart -and $null -ne $VacationEnd) {
            $vacationStartColumn = switch ($VacationSlot) { 1 { 4 } 2 { 7 } 3 { 10 } default { throw "Vacation slot must be 1, 2, or 3" } }
            $teamMembersTable.DataBodyRange.Cells($Row, $vacationStartColumn).Value2 = [datetime]$VacationStart
            $teamMembersTable.DataBodyRange.Cells($Row, $vacationStartColumn + 1).Value2 = [datetime]$VacationEnd
        }
    }

    function Reset-ActivePlanResourceScenario([double]$AnEstimate) {
        for ($row = 1; $row -le $activeTable.DataBodyRange.Rows.Count; $row++) {
            for ($col = 6; $col -le 20; $col++) {
                $activeTable.DataBodyRange.Cells($row, $col).ClearContents() | Out-Null
            }
        }
        $activeTable.DataBodyRange.Cells(1, 8).Value2 = "Resource calendar scenario"
        $activeTable.DataBodyRange.Cells(1, 12).Value2 = $AnEstimate
        $activeTable.DataBodyRange.Cells(1, 19).Value2 = "Готова аналитика"
    }

    for ($row = 1; $row -le $holidayTable.DataBodyRange.Rows.Count; $row++) {
        $holidayTable.DataBodyRange.Cells($row, 4).Value2 = "Нет"
    }

    Clear-TeamMembersForScheduleScenario
    Set-TeamMemberForScheduleScenario 1 "Аналитик" 0.5
    Reset-ActivePlanResourceScenario 1.4
    $excel.CalculateFull()
    $excel.Run("RunQuarterPlanRecalculate")
    Assert-NumberClose ([double]$activeTable.DataBodyRange.Cells(1, 16).Value2) 2 "Effort stays normalized with 50 percent allocation" 0.0001
    Assert-NumberClose ([double]$planSheet.Range("V6").Value2) 4 "AN duration stretches with 50 percent allocation" 0.0001
    Assert-ExcelDateEquals $activeTable.DataBodyRange.Cells(1, 17) ([datetime]"2026-04-06") "Ready date stretches with 50 percent allocation"

    Clear-TeamMembersForScheduleScenario
    Set-TeamMemberForScheduleScenario 1 "Аналитик" 1 ([datetime]"2026-04-02") ([datetime]"2026-04-03")
    Reset-ActivePlanResourceScenario 2.1
    $excel.CalculateFull()
    $excel.Run("RunQuarterPlanRecalculate")
    Assert-NumberClose ([double]$activeTable.DataBodyRange.Cells(1, 16).Value2) 3 "Effort stays normalized through vacation pause" 0.0001
    Assert-NumberClose ([double]$planSheet.Range("V6").Value2) 5 "AN duration includes vacation pause without replacement" 0.0001
    Assert-ExcelDateEquals $activeTable.DataBodyRange.Cells(1, 17) ([datetime]"2026-04-07") "Ready date waits for analyst vacation without replacement"

    Clear-TeamMembersForScheduleScenario
    Set-TeamMemberForScheduleScenario -Row 1 -Role "Аналитик" -Allocation 1 -VacationStart ([datetime]"2026-04-02") -VacationEnd ([datetime]"2026-04-03") -VacationSlot 2
    Reset-ActivePlanResourceScenario 2.1
    $excel.CalculateFull()
    $excel.Run("RunQuarterPlanRecalculate")
    Assert-NumberClose ([double]$planSheet.Range("V6").Value2) 5 "AN duration includes second vacation pause" 0.0001
    Assert-ExcelDateEquals $activeTable.DataBodyRange.Cells(1, 17) ([datetime]"2026-04-07") "Ready date waits for second analyst vacation"

    Clear-TeamMembersForScheduleScenario
    Set-TeamMemberForScheduleScenario -Row 1 -Role "Аналитик" -Allocation 1 -VacationStart ([datetime]"2026-04-02") -VacationEnd ([datetime]"2026-04-03") -VacationSlot 3
    Reset-ActivePlanResourceScenario 2.1
    $excel.CalculateFull()
    $excel.Run("RunQuarterPlanRecalculate")
    Assert-NumberClose ([double]$planSheet.Range("V6").Value2) 5 "AN duration includes third vacation pause" 0.0001
    Assert-ExcelDateEquals $activeTable.DataBodyRange.Cells(1, 17) ([datetime]"2026-04-07") "Ready date waits for third analyst vacation"

    Clear-TeamMembersForScheduleScenario
    Set-TeamMemberForScheduleScenario 1 "Аналитик" 1 ([datetime]"2026-04-02") ([datetime]"2026-04-03")
    Set-TeamMemberForScheduleScenario 2 "Аналитик" 1
    Reset-ActivePlanResourceScenario 2.1
    $excel.CalculateFull()
    $excel.Run("RunQuarterPlanRecalculate")
    Assert-NumberClose ([double]$planSheet.Range("V6").Value2) 3 "AN duration uses replacement analyst during vacation" 0.0001
    Assert-ExcelDateEquals $activeTable.DataBodyRange.Cells(1, 17) ([datetime]"2026-04-03") "Ready date uses replacement analyst during vacation"
    Write-Pass "COM sheet 04 allocation and all vacation slots scheduling"

    $analyticsFormatFound = $false
    $devFormatFound = $false
    $backendFormatFound = $false
    $frontendBeFormatFound = $false
    $frontendQaFormatFound = $false
    $deferredFormatFound = $false
    foreach ($condition in @($planSheet.Range("M6:O25").FormatConditions)) {
        if ([string]$condition.Formula1 -like "*Готова аналитика*") { $analyticsFormatFound = $true }
    }
    foreach ($condition in @($planSheet.Range("O6:O25").FormatConditions)) {
        if ([string]$condition.Formula1 -like "*Готова разработка`"*") { $devFormatFound = $true }
        if ([string]$condition.Formula1 -like "*Готова разработка (фронт)*") { $frontendQaFormatFound = $true }
    }
    foreach ($condition in @($planSheet.Range("N6:O25").FormatConditions)) {
        if ([string]$condition.Formula1 -like "*Готова разработка (бэк)*") { $backendFormatFound = $true }
    }
    foreach ($condition in @($planSheet.Range("M6:M25").FormatConditions)) {
        if ([string]$condition.Formula1 -like "*Готова разработка (фронт)*") { $frontendBeFormatFound = $true }
    }
    foreach ($condition in @($planSheet.Range("K6:O25").FormatConditions)) {
        if ([string]$condition.Formula1 -like "*Отложено*") { $deferredFormatFound = $true }
    }
    if (!$analyticsFormatFound) { throw "Analysis-ready conditional format not found on M6:O25" }
    if (!$devFormatFound) { throw "Dev-ready conditional format not found on O6:O25" }
    if (!$backendFormatFound) { throw "Backend-ready conditional format not found on N6:O25" }
    if (!$frontendBeFormatFound) { throw "Frontend-ready BE conditional format not found on M6:M25" }
    if (!$frontendQaFormatFound) { throw "Frontend-ready QA conditional format not found on O6:O25" }
    if (!$deferredFormatFound) { throw "Deferred conditional format not found on K6:O25" }
    Write-Pass "COM sheet 04 excluded estimate formatting"

    $excel.Run("RunTaskEstimateRepairActionButtons")
    $excel.Run("RunQuarterPlanRepairActionButtons")

    $expectedActions = @{
        "btnQuarterPlanReloadBacklog" = "RunQuarterPlanReloadBacklog"
        "btnQuarterPlanRecalculate" = "RunQuarterPlanRecalculate"
        "btnQuarterPlanExportBacklog" = "RunQuarterPlanExportBacklog"
        "btnQuarterPlanAddAllBacklog" = "RunQuarterPlanAddAllBacklogToPlan"
        "btnJqlA" = "RunCopyJql"
        "btnJqlG" = "RunCopyJql"
        "btnJqlB" = "RunCopyJql"
    }
    $expectedAnchors = @{
        "btnQuarterPlanReloadBacklog" = "G2"
        "btnQuarterPlanRecalculate" = "Q2"
        "btnQuarterPlanExportBacklog" = "R2"
        "btnQuarterPlanAddAllBacklog" = [string]$workbookLayout.backlog.actionCell
        "btnJqlA" = [string]$workbookLayout.activePlan.jqlActionCell
        "btnJqlG" = [string]$workbookLayout.greyZone.jqlActionCell
        "btnJqlB" = [string]$workbookLayout.backlog.jqlActionCell
    }
    foreach ($shapeName in $expectedActions.Keys) {
        $shape = $planSheet.Shapes.Item($shapeName)
        $onAction = [string]$shape.OnAction
        Assert-PlainMacroName $onAction $shapeName
        if ($onAction -ne $expectedActions[$shapeName]) {
            throw "Shape $shapeName OnAction is $onAction, expected $($expectedActions[$shapeName])"
        }
        Assert-TextEquals ([string]$shape.TopLeftCell.Address($false, $false)) $expectedAnchors[$shapeName] "$shapeName button anchor"
        if ($shapeName -eq "btnQuarterPlanExportBacklog") {
            Assert-TextEquals ([string]$shape.TextFrame.Characters().Text) "Экспорт плана" "Export plan button caption"
        }
        if ($shapeName -in @("btnJqlA", "btnJqlG", "btnJqlB")) {
            Assert-TextEquals ([string]$shape.TextFrame.Characters().Text) "JQL" "$shapeName button caption"
        }
        Assert-ActionShapeDesign $shape $shapeName
    }
    Write-Pass "COM top action buttons"

    $excel.Run("'" + $workbook.Name + "'!ThisWorkbook.AssertQuarterPlanBacklogDefaultExportFileNameForTest")
    Write-Pass "COM sheet 04 backlog export default filename"

    $excel.Run($reloadMacro, $false)
    $excel.Run("RunQuarterPlanAddAllBacklogToPlan")
    $statusFixtures = @(
        @{ Row = 1; Status = "Готова аналитика"; Description = $taskExpectedDescription },
        @{ Row = 2; Status = "Готова разработка (бэк)"; Description = "Export style backend" },
        @{ Row = 3; Status = "Готова разработка (фронт)"; Description = "Export style frontend" },
        @{ Row = 4; Status = "Отложено"; Description = "Export style deferred" }
    )
    foreach ($fixture in $statusFixtures) {
        $row = [int]$fixture.Row
        $activeTable.DataBodyRange.Cells($row, 8).Value2 = [string]$fixture.Description
        $activeTable.DataBodyRange.Cells($row, 11).Value2 = 2
        $activeTable.DataBodyRange.Cells($row, 12).Value2 = 3
        $activeTable.DataBodyRange.Cells($row, 13).Value2 = 5
        $activeTable.DataBodyRange.Cells($row, 14).Value2 = 7
        $activeTable.DataBodyRange.Cells($row, 15).Value2 = 11
        $activeTable.DataBodyRange.Cells($row, 19).Value2 = [string]$fixture.Status
    }
    $excel.Run("RunQuarterPlanExportBacklogPath", $backlogExportPath)
    if (!(Test-Path -LiteralPath $backlogExportPath)) {
        throw "Quarter plan export file was not created: $backlogExportPath"
    }
    $backlogExportWorkbook = $excel.Workbooks.Open($backlogExportPath, 0, $true)
    try {
        $backlogExportSheet = $backlogExportWorkbook.Worksheets.Item(1)
        Assert-TextEquals ([string]$backlogExportSheet.Cells(1, 5).Text) "Емкость" "Quarter plan export capacity label"
        Assert-TextEquals ([string]$backlogExportSheet.Cells(2, 5).Text) "План+риск" "Quarter plan export planned with risk label"
        Assert-TextEquals ([string]$backlogExportSheet.Cells(3, 5).Text) "Баланс" "Quarter plan export balance label"
        foreach ($statCol in 6..10) {
            Assert-CellNotBlank $backlogExportSheet.Cells(1, $statCol) "Quarter plan export capacity value column $statCol"
            Assert-CellNotBlank $backlogExportSheet.Cells(2, $statCol) "Quarter plan export planned with risk value column $statCol"
            Assert-CellNotBlank $backlogExportSheet.Cells(3, $statCol) "Quarter plan export balance value column $statCol"
        }
        Assert-TextEquals ([string]$backlogExportSheet.Cells(5, 1).Text) "Приоритет" "Quarter plan export first header"
        Assert-TextEquals ([string]$backlogExportSheet.Cells(5, 3).Text) "Описание" "Quarter plan export description header"
        Assert-TextEquals ([string]$backlogExportSheet.Cells(5, 11).Text) "Трудозатраты" "Quarter plan export effort header"
        Assert-TextEquals ([string]$backlogExportSheet.Cells(5, 12).Text) "Завершение (план)" "Quarter plan export ready date header"
        Assert-TextEquals ([string]$backlogExportSheet.Cells(5, 15).Text) "Комментарий" "Quarter plan export last header"
        Assert-TextEquals ([string]$backlogExportSheet.Cells(6, 3).Text) $taskExpectedDescription "Quarter plan export first task description"
        if (![string]::IsNullOrWhiteSpace([string]$backlogExportSheet.Cells(5, 16).Text)) {
            throw "Quarter plan export contains an unexpected column after F:T user data"
        }
        if ([string]$backlogExportSheet.Cells(5, 1).Text -eq "+") {
            throw "Quarter plan export contains action columns"
        }
        Assert-Strikethrough $backlogExportSheet.Cells(6, 6) $false "Analysis-ready exported DE"
        Assert-Strikethrough $backlogExportSheet.Cells(6, 7) $false "Analysis-ready exported AN"
        Assert-Strikethrough $backlogExportSheet.Cells(6, 8) $true "Analysis-ready exported BE"
        Assert-Strikethrough $backlogExportSheet.Cells(6, 9) $true "Analysis-ready exported FE"
        Assert-Strikethrough $backlogExportSheet.Cells(6, 10) $true "Analysis-ready exported QA"
        Assert-Strikethrough $backlogExportSheet.Cells(7, 6) $false "Backend-ready exported DE"
        Assert-Strikethrough $backlogExportSheet.Cells(7, 7) $false "Backend-ready exported AN"
        Assert-Strikethrough $backlogExportSheet.Cells(7, 8) $false "Backend-ready exported BE"
        Assert-Strikethrough $backlogExportSheet.Cells(7, 9) $true "Backend-ready exported FE"
        Assert-Strikethrough $backlogExportSheet.Cells(7, 10) $true "Backend-ready exported QA"
        Assert-Strikethrough $backlogExportSheet.Cells(8, 6) $false "Frontend-ready exported DE"
        Assert-Strikethrough $backlogExportSheet.Cells(8, 7) $false "Frontend-ready exported AN"
        Assert-Strikethrough $backlogExportSheet.Cells(8, 8) $true "Frontend-ready exported BE"
        Assert-Strikethrough $backlogExportSheet.Cells(8, 9) $false "Frontend-ready exported FE"
        Assert-Strikethrough $backlogExportSheet.Cells(8, 10) $true "Frontend-ready exported QA"
        foreach ($estimateCol in 6..10) {
            Assert-Strikethrough $backlogExportSheet.Cells(9, $estimateCol) $true "Deferred exported estimate column $estimateCol"
        }
    } finally {
        $backlogExportWorkbook.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($backlogExportWorkbook) | Out-Null
    }
    Write-Pass "COM sheet 04 active plan export"

    $estimateTable.ListColumns("ЗНИ/Jira").DataBodyRange.ClearContents()
    $estimateLevelRange = $estimateSheet.Range(
        $estimateSheet.Cells([int]$workbookLayout.taskEstimates.dataStartRow, 30),
        $estimateSheet.Cells([int]$workbookLayout.taskEstimates.dataEndRow, 30)
    )
    $estimateLevelRange.ClearContents()
    $estimateTable.DataBodyRange.Cells(1, 4).Value2 = " ROOT-1 "
    $estimateTable.DataBodyRange.Cells(2, 4).Value2 = "CHILD-1"
    $estimateTable.DataBodyRange.Cells(3, 4).Value2 = "root-1"
    $estimateTable.DataBodyRange.Cells(4, 4).Value2 = "ROOT-2"
    $estimateSheet.Cells([int]$workbookLayout.taskEstimates.dataStartRow, 30).Value2 = 0
    $estimateSheet.Cells([int]$workbookLayout.taskEstimates.dataStartRow + 1, 30).Value2 = 1
    $estimateSheet.Cells([int]$workbookLayout.taskEstimates.dataStartRow + 2, 30).Value2 = 0
    $estimateSheet.Cells([int]$workbookLayout.taskEstimates.dataStartRow + 3, 30).Value2 = 0
    $excel.Run("'" + $workbook.Name + "'!ThisWorkbook.AssertJqlForTableNameForTest", "tblTaskEstimates", "issuekey in (ROOT-1, ROOT-2)")

    $activeTable.ListColumns("ЗНИ/Jira").DataBodyRange.ClearContents()
    $greyTableForLayout.ListColumns("ЗНИ/Jira").DataBodyRange.ClearContents()
    $backlogTable.ListColumns("ЗНИ/Jira").DataBodyRange.ClearContents()
    $activeTable.DataBodyRange.Cells(1, 9).Value2 = " PLAN-1 "
    $activeTable.DataBodyRange.Cells(2, 9).Value2 = "plan-1"
    $activeTable.DataBodyRange.Cells(3, 9).Value2 = "PLAN-2"
    $greyTableForLayout.DataBodyRange.Cells(1, 9).Value2 = "GREY-1"
    $backlogTable.DataBodyRange.Cells(1, 9).Value2 = "BACK-1"
    $excel.Run("'" + $workbook.Name + "'!ThisWorkbook.AssertJqlForTableNameForTest", "tblPlanActive", "issuekey in (PLAN-1, PLAN-2)")
    $excel.Run("'" + $workbook.Name + "'!ThisWorkbook.AssertJqlForTableNameForTest", "tblPlanGrey", "issuekey in (GREY-1)")
    $excel.Run("'" + $workbook.Name + "'!ThisWorkbook.AssertJqlForTableNameForTest", "tblPlanBacklog", "issuekey in (BACK-1)")
    Write-Pass "COM JQL generation"

    $referencesSheet = $workbook.Worksheets.Item("99_Правила планирования")
    $jqlClipboardCell = $referencesSheet.Range([string]$workbookLayout.references.jqlClipboardCell)
    $jqlClipboardCell.EntireColumn.Hidden = $true
    $workbook.Saved = $true
    $savedBeforeJqlCopy = [bool]$workbook.Saved
    $excel.Run("'" + $workbook.Name + "'!ThisWorkbook.CopyJqlForTableNameForTest", "tblPlanActive")
    Start-Sleep -Milliseconds 300
    Assert-TextEquals ([string]$jqlClipboardCell.Value2) "issuekey in (PLAN-1, PLAN-2)" "JQL staging cell"
    if ([bool]$jqlClipboardCell.EntireColumn.Hidden) {
        throw "JQL staging column remained hidden, so external applications receive an empty plain-text clipboard value"
    }
    $plainClipboardText = (Get-ClipboardPlainText) -replace "(`r`n|`n|`r)+$", ""
    Assert-TextEquals $plainClipboardText "issuekey in (PLAN-1, PLAN-2)" "Windows plain-text clipboard"
    if ([bool]$workbook.Saved -ne $savedBeforeJqlCopy) {
        throw "JQL copy changed ThisWorkbook.Saved"
    }
    Write-Pass "COM JQL external plain-text clipboard"

    foreach ($sheet in @($workbook.Worksheets)) {
        foreach ($shape in @($sheet.Shapes)) {
            $onAction = [string]$shape.OnAction
            if (![string]::IsNullOrWhiteSpace($onAction)) {
                Assert-PlainMacroName $onAction $shape.Name
                Assert-ActionShapeDesign $shape $shape.Name
            }
        }
    }
    Write-Pass "COM all shape OnAction values"

    $errorTokens = @("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A")
    foreach ($sheet in @($workbook.Worksheets)) {
        $usedRange = $sheet.UsedRange
        $values = $usedRange.Text
        foreach ($token in $errorTokens) {
            $found = $usedRange.Find($token)
            if ($null -ne $found) {
                throw "Formula error token $token found at $($sheet.Name)!$($found.Address($false, $false))"
            }
        }
    }
    Write-Pass "COM formula error scan"

    $workbook.Close($false)
    $workbook = $null
} catch {
    Write-Fail "Excel COM acceptance" $_.Exception.Message
    try { $workbook.Close($false) | Out-Null } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null } catch {}
    try { $excel.Quit() | Out-Null } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    $workbook = $null
    $excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    exit 1
} finally {
    try { $workbook.Close($false) | Out-Null } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null } catch {}
    try { $excel.Quit() | Out-Null } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
