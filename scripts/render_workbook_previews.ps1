param(
    [string]$WorkbookPath = "outputs\quarter_planning_step1.xlsx",
    [string]$OutputDirectory = "outputs\previews_step1"
)

$ErrorActionPreference = "Stop"

if (-not $IsWindows -and $PSVersionTable.PSEdition -eq "Core") {
    throw "Preview generation currently requires Windows desktop Excel. Workbook build and static verification remain portable."
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$fullWorkbookPath = [System.IO.Path]::GetFullPath((Join-Path $repoRoot $WorkbookPath))
$fullOutputDirectory = [System.IO.Path]::GetFullPath((Join-Path $repoRoot $OutputDirectory))
if (!(Test-Path -LiteralPath $fullWorkbookPath)) {
    throw "Workbook not found: $fullWorkbookPath"
}
New-Item -ItemType Directory -Path $fullOutputDirectory -Force | Out-Null

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    $workbook = $excel.Workbooks.Open($fullWorkbookPath, 0, $true)

    foreach ($sheet in @($workbook.Worksheets)) {
        $usedRange = $sheet.UsedRange
        $chartObject = $null
        try {
            $sheet.Activate() | Out-Null
            $usedRange.CopyPicture(1, 2) | Out-Null
            Start-Sleep -Milliseconds 150
            $width = [Math]::Max(320, [Math]::Min(6000, [double]$usedRange.Width))
            $height = [Math]::Max(200, [Math]::Min(12000, [double]$usedRange.Height))
            $chartObject = $sheet.ChartObjects().Add(0, 0, $width, $height)
            $chartObject.Activate() | Out-Null
            $chartObject.Chart.Paste() | Out-Null
            Start-Sleep -Milliseconds 150
            $safeName = [regex]::Replace($sheet.Name, '[\\/:*?"<>|]', '_')
            $outputPath = Join-Path $fullOutputDirectory "$safeName.png"
            if (!$chartObject.Chart.Export($outputPath, "PNG")) {
                throw "Excel could not export preview for sheet '$($sheet.Name)'"
            }
            Write-Host "PREVIEW $outputPath"
        } finally {
            if ($chartObject -ne $null) { $chartObject.Delete() | Out-Null }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($usedRange) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null
        }
    }
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
