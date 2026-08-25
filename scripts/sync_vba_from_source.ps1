param(
    [string]$WorkbookPath = "outputs\quarter_planning_step2.xlsm",
    [string]$SourcePath = "assets\vba\ThisWorkbook_holiday_macro.txt",
    [string]$ActionModuleSourcePath = "assets\vba\QuarterPlanActions_module.txt",
    [string]$VbaProjectPath = "assets\vba\vbaProject.step2.bin",
    [string]$TemplatePath = "assets\vba\quarter_planning_macro_template.xlsm"
)

$ErrorActionPreference = "Stop"

function Resolve-ProjectPath([string]$Path) {
    $resolvedPath = $Path
    if (![System.IO.Path]::IsPathRooted($resolvedPath)) {
        $resolvedPath = Join-Path (Get-Location).Path $resolvedPath
    }
    return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($resolvedPath)
}

$workbookFullPath = Resolve-ProjectPath $WorkbookPath
$sourceFullPath = Resolve-ProjectPath $SourcePath
$actionModuleSourceFullPath = Resolve-ProjectPath $ActionModuleSourcePath
$vbaProjectFullPath = Resolve-ProjectPath $VbaProjectPath
$templateFullPath = Resolve-ProjectPath $TemplatePath

if (!(Test-Path -LiteralPath $workbookFullPath)) {
    throw "Workbook not found: $workbookFullPath"
}
if (!(Test-Path -LiteralPath $sourceFullPath)) {
    throw "VBA source not found: $sourceFullPath"
}
if (!(Test-Path -LiteralPath $actionModuleSourceFullPath)) {
    throw "Action module source not found: $actionModuleSourceFullPath"
}

$securityKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
if (!(Test-Path $securityKey)) {
    New-Item -Path $securityKey -Force | Out-Null
}
New-ItemProperty -Path $securityKey -Name AccessVBOM -Value 1 -PropertyType DWord -Force | Out-Null

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1

    $workbook = $excel.Workbooks.Open($workbookFullPath, 0, $false)
    if (!$workbook.HasVBProject) {
        throw "Workbook has no VBA project: $workbookFullPath"
    }

    $component = $workbook.VBProject.VBComponents.Item("ThisWorkbook")
    $codeModule = $component.CodeModule
    if ($codeModule.CountOfLines -gt 0) {
        $codeModule.DeleteLines(1, $codeModule.CountOfLines)
    }

    $source = Get-Content -LiteralPath $sourceFullPath -Raw -Encoding UTF8
    $codeModule.AddFromString($source)

    $actionComponent = $null
    try {
        $actionComponent = $workbook.VBProject.VBComponents.Item("QuarterPlanActions")
    } catch {
        $actionComponent = $workbook.VBProject.VBComponents.Add(1)
        $actionComponent.Name = "QuarterPlanActions"
    }

    $actionCodeModule = $actionComponent.CodeModule
    if ($actionCodeModule.CountOfLines -gt 0) {
        $actionCodeModule.DeleteLines(1, $actionCodeModule.CountOfLines)
    }

    $actionSource = Get-Content -LiteralPath $actionModuleSourceFullPath -Raw -Encoding UTF8
    $actionCodeModule.AddFromString($actionSource)

    $workbook.Save()
    $workbook.Close($false)
    $workbook = $null
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

Copy-Item -LiteralPath $workbookFullPath -Destination $templateFullPath -Force

$nodeExecutable = $env:QUARTER_PLANNING_NODE_EXECUTABLE
if ([string]::IsNullOrWhiteSpace($nodeExecutable)) {
    $nodeExecutable = "node"
}
& $nodeExecutable ".\scripts\sanitize_workbook_metadata.mjs" $workbookFullPath $templateFullPath
if ($LASTEXITCODE -ne 0) {
    throw "sanitize_workbook_metadata.mjs failed"
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead($workbookFullPath)
try {
    $entry = $zip.GetEntry("xl/vbaProject.bin")
    if ($entry -eq $null) {
        throw "xl/vbaProject.bin not found after Excel save"
    }
    $outDir = Split-Path -Parent $vbaProjectFullPath
    if (!(Test-Path -LiteralPath $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }
    $inStream = $entry.Open()
    try {
        $outStream = [System.IO.File]::Create($vbaProjectFullPath)
        try {
            $inStream.CopyTo($outStream)
        } finally {
            $outStream.Dispose()
        }
    } finally {
        $inStream.Dispose()
    }
} finally {
    $zip.Dispose()
}

Write-Host "SYNCED $sourceFullPath -> $workbookFullPath"
Write-Host "UPDATED $vbaProjectFullPath"
Write-Host "UPDATED $templateFullPath"
