[CmdletBinding()]
param(
    [switch]$CoreOnly,
    [switch]$SkipGui,
    [switch]$SkipEnvironmentCompatibility,
    [switch]$IncludeWorkbookHealth
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$isWindows = [Environment]::OSVersion.Platform -eq [PlatformID]::Win32NT

Write-Host "Parsing PowerShell files..."
$files = @(Get-ChildItem -Path $projectRoot -Filter *.ps1 -File -Recurse |
    Where-Object { $_.FullName -notmatch "\\.git\\|\\output\\|\\dist\\" })
foreach ($file in $files) {
    $tokens = $null
    $errors = $null
    [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors) | Out-Null
    if ($errors.Count -gt 0) {
        throw "Parse errors in $($file.FullName): $($errors | ForEach-Object { $_.Message } | Out-String)"
    }
}

if (-not $CoreOnly -and -not $SkipGui -and $isWindows) {
    Write-Host "Building WinForms launcher..."
    $windowsPowerShell = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
    if (-not (Test-Path -LiteralPath $windowsPowerShell)) {
        $windowsPowerShell = "powershell.exe"
    }
    & $windowsPowerShell -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File (Join-Path $projectRoot "app\cli\Launch-AnalyticsJobCrawlerGui.ps1") -BuildSelfTest
    if ($LASTEXITCODE -ne 0) {
        throw "WinForms launcher build self-test failed."
    }
    & $windowsPowerShell -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File (Join-Path $projectRoot "app\cli\Launch-AnalyticsJobCrawlerGui.ps1") -SmokeTest
    if ($LASTEXITCODE -ne 0) {
        throw "WinForms launcher smoke test failed."
    }
    & $windowsPowerShell -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File (Join-Path $projectRoot "app\cli\Launch-AnalyticsJobCrawlerGui.ps1") -RunSmokeTest
    if ($LASTEXITCODE -ne 0) {
        throw "WinForms launcher run smoke test failed."
    }
}
elseif (-not $isWindows) {
    Write-Host "Skipping WinForms launcher tests on this platform."
}
else {
    Write-Host "Skipping WinForms launcher tests."
}

$testScripts = @(
    "Test-ScoringRules.ps1",
    "Test-ParserFixtures.ps1",
    "Test-ProfileBuilder.ps1",
    "Test-SourceAdapters.ps1",
    "Test-PipelineGuards.ps1",
    "Test-OpenXmlWorkbook.ps1",
    "Test-OpenXmlWorkbookHealth.ps1",
    "Test-FeedbackSuggestions.ps1",
    "Test-Integration.ps1",
    "Test-JobCrawlerConfig.ps1",
    "Test-PortableProjectRoot.ps1"
)
if (-not $CoreOnly -and -not $SkipEnvironmentCompatibility) {
    $testScripts += "Test-EnvironmentCompatibility.ps1"
}

foreach ($testScript in $testScripts) {
    $path = Join-Path $PSScriptRoot $testScript
    Write-Host ("Running {0}..." -f $testScript)
    & $path
}

Write-Host "Running release safety check..."
& (Join-Path $projectRoot "tools\release\Test-ReleaseSafety.ps1")

if ($IncludeWorkbookHealth) {
    Write-Host "Running workbook health check..."
    & (Join-Path $PSScriptRoot "Test-JobTrackerHealth.ps1")
}

Write-Host "All tests passed."
