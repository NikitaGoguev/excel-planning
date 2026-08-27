param(
    [string]$WorkbookPath = "outputs\quarter_planning_step2.xlsm",
    [string]$ContractPath = "contracts\vba.contract.json",
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
$contractFullPath = Resolve-ProjectPath $ContractPath
$vbaProjectFullPath = Resolve-ProjectPath $VbaProjectPath
$templateFullPath = Resolve-ProjectPath $TemplatePath

if (!(Test-Path -LiteralPath $workbookFullPath)) {
    throw "Workbook not found: $workbookFullPath"
}
if (!(Test-Path -LiteralPath $contractFullPath)) { throw "VBA contract not found: $contractFullPath" }

$nodeExecutable = $env:QUARTER_PLANNING_NODE_EXECUTABLE
if ([string]::IsNullOrWhiteSpace($nodeExecutable)) { $nodeExecutable = "node" }
& $nodeExecutable ".\scripts\generate_vba_limits.mjs" "--check"
if ($LASTEXITCODE -ne 0) { throw "Generated VBA limits are stale; run npm run generate:limits" }

$vbaContract = Get-Content -LiteralPath $contractFullPath -Raw -Encoding UTF8 | ConvertFrom-Json
$components = @($vbaContract.components)
if ($components.Count -eq 0) { throw "VBA component manifest is empty" }
foreach ($item in $components) {
    $item | Add-Member -NotePropertyName FullSourcePath -NotePropertyValue (Resolve-ProjectPath ([string]$item.source))
    if (!(Test-Path -LiteralPath $item.FullSourcePath)) { throw "VBA source not found: $($item.FullSourcePath)" }
    if ($item.type -eq "form") {
        $formBinaryPath = [System.IO.Path]::ChangeExtension($item.FullSourcePath, ".frx")
        if (!(Test-Path -LiteralPath $formBinaryPath)) { throw "VBA form binary companion not found: $formBinaryPath" }
    }
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

    $expectedStandardModules = @($components | Where-Object { $_.type -eq "standard" } | ForEach-Object { [string]$_.name })
    $staleComponents = @()
    foreach ($existing in @($workbook.VBProject.VBComponents)) {
        if ([int]$existing.Type -eq 1 -and [string]$existing.Name -like "QuarterPlan*" -and $expectedStandardModules -notcontains [string]$existing.Name) {
            $staleComponents += $existing
        }
    }
    foreach ($stale in $staleComponents) { $workbook.VBProject.VBComponents.Remove($stale) }

    foreach ($item in $components) {
        $component = $null
        try { $component = $workbook.VBProject.VBComponents.Item([string]$item.name) } catch {}
        if ($item.type -eq "form") {
            if ($component -ne $null) { $workbook.VBProject.VBComponents.Remove($component) }
            $component = $workbook.VBProject.VBComponents.Import($item.FullSourcePath)
            if ([string]$component.Name -ne [string]$item.name) { $component.Name = [string]$item.name }
            continue
        }
        if ($null -eq $component) {
            if ($item.type -ne "standard") { throw "Document VBA component not found: $($item.name)" }
            $component = $workbook.VBProject.VBComponents.Add(1)
            $component.Name = [string]$item.name
        }
        $codeModule = $component.CodeModule
        if ($codeModule.CountOfLines -gt 0) { $codeModule.DeleteLines(1, $codeModule.CountOfLines) }
        $source = Get-Content -LiteralPath $item.FullSourcePath -Raw -Encoding UTF8
        $codeModule.AddFromString($source)
    }

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

Write-Host "SYNCED $($components.Count) VBA components -> $workbookFullPath"
Write-Host "UPDATED $vbaProjectFullPath"
Write-Host "UPDATED $templateFullPath"
