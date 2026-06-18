[CmdletBinding()]
param(
    [switch]$RequireFullWindowsStack
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
. (Join-Path $projectRoot "app\core\JobTracker.Common.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Config.ps1")

$issues = New-Object System.Collections.Generic.List[string]
$rows = New-Object System.Collections.Generic.List[object]

function Add-CompatibilityRow {
    param(
        [string]$Area,
        [string]$Status,
        [string]$Detail,
        [switch]$IsBlocking
    )

    $rows.Add([PSCustomObject]@{
        Area   = $Area
        Status = $Status
        Detail = $Detail
    }) | Out-Null

    if ($IsBlocking) {
        $issues.Add(("{0}: {1}" -f $Area, $Detail)) | Out-Null
    }
}

function Test-WindowsPlatform {
    return [Environment]::OSVersion.Platform -eq [PlatformID]::Win32NT
}

function Test-ComObjectAvailable {
    param([string]$ProgId)

    $instance = $null
    try {
        $instance = New-Object -ComObject $ProgId
        return $true
    }
    catch {
        return $false
    }
    finally {
        if ($null -ne $instance) {
            try {
                if ($ProgId -eq "Excel.Application") {
                    $instance.Quit()
                }
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($instance)
            }
            catch {
            }
        }
    }
}

$isWindows = Test-WindowsPlatform
$psVersion = $PSVersionTable.PSVersion
$edition = if ($PSVersionTable.ContainsKey("PSEdition")) { [string]$PSVersionTable.PSEdition } else { "Desktop" }

Add-CompatibilityRow -Area "OS" -Status $(if ($isWindows) { "full" } else { "limited" }) -Detail ([Environment]::OSVersion.VersionString)
Add-CompatibilityRow -Area "PowerShell" -Status $(if ($psVersion.Major -ge 5) { "ok" } else { "blocking" }) -Detail ("{0} ({1})" -f $psVersion, $edition) -IsBlocking:($psVersion.Major -lt 5)

foreach ($relativePath in @(
    "app\cli\Find-AnalyticsJobs.ps1",
    "app\cli\Launch-AnalyticsJobCrawlerGui.ps1",
    "app\core\JobTracker.Config.ps1",
    "app\core\JobTracker.Context.ps1",
    "app\core\JobTracker.Profile.ps1",
    "app\core\JobTracker.Pipeline.ps1",
    "app\core\JobTracker.SourceAdapter.ps1",
    "app\core\JobTracker.OutputMaintenance.ps1",
    "app\core\JobTracker.OpenXml.ps1",
    "config\sources.json",
    "config\runtime.json",
    "config\crawl_modes.json",
    "config\matching_rules.json",
    "config\workbook.json",
    "config\profiles\digital_analytics.json"
)) {
    $path = Join-Path $projectRoot $relativePath
    Add-CompatibilityRow -Area ("file:{0}" -f $relativePath) -Status $(if (Test-Path -LiteralPath $path) { "ok" } else { "missing" }) -Detail $relativePath -IsBlocking:(-not (Test-Path -LiteralPath $path))
}

try {
    $config = Get-JobCrawlerConfig -ConfigDirectory (Join-Path $projectRoot "config")
    $validation = Test-JobCrawlerConfig -Config $config
    Add-CompatibilityRow -Area "Config" -Status $(if ($validation.IsValid) { "ok" } else { "blocking" }) -Detail $(if ($validation.IsValid) { "Active profile: {0}" -f $config.Profile.Id } else { ($validation.Issues -join "; ") }) -IsBlocking:(-not $validation.IsValid)
}
catch {
    Add-CompatibilityRow -Area "Config" -Status "blocking" -Detail $_.Exception.Message -IsBlocking
}

if ($isWindows) {
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
        Add-CompatibilityRow -Area "WinForms GUI" -Status "ok" -Detail "System.Windows.Forms and System.Drawing are available."
    }
    catch {
        Add-CompatibilityRow -Area "WinForms GUI" -Status "missing" -Detail $_.Exception.Message -IsBlocking:$RequireFullWindowsStack
    }

    $openXmlWriterAvailable = Test-Path -LiteralPath (Join-Path $projectRoot "app\core\JobTracker.OpenXml.ps1")
    Add-CompatibilityRow -Area "No-Excel XLSX writer" -Status $(if ($openXmlWriterAvailable) { "ok" } else { "missing" }) -Detail $(if ($openXmlWriterAvailable) { "Built-in OpenXML writer can create jobs_tracker.xlsx without desktop Excel." } else { "OpenXML writer module is missing." }) -IsBlocking:(-not $openXmlWriterAvailable)
    $excelAvailable = Test-ComObjectAvailable -ProgId "Excel.Application"
    Add-CompatibilityRow -Area "Excel workbook" -Status $(if ($excelAvailable) { "optional" } else { "optional missing" }) -Detail $(if ($excelAvailable) { "Excel COM automation is available for the richest workbook formatting." } else { "Desktop Excel is not installed; auto mode will use the built-in OpenXML writer." })
    Add-CompatibilityRow -Area "Windows launchers" -Status "ok" -Detail ".cmd and .vbs launchers are supported on Windows."
}
else {
    $openXmlWriterAvailable = Test-Path -LiteralPath (Join-Path $projectRoot "app/core/JobTracker.OpenXml.ps1")
    Add-CompatibilityRow -Area "WinForms GUI" -Status "unsupported" -Detail "The current GUI uses Windows Forms and is Windows-only."
    Add-CompatibilityRow -Area "No-Excel XLSX writer" -Status $(if ($openXmlWriterAvailable) { "ok" } else { "missing" }) -Detail $(if ($openXmlWriterAvailable) { "Built-in OpenXML writer can create jobs_tracker.xlsx without desktop Excel." } else { "OpenXML writer module is missing." }) -IsBlocking:(-not $openXmlWriterAvailable)
    Add-CompatibilityRow -Area "Excel workbook" -Status "unsupported optional" -Detail "Excel COM automation is Windows-only; use OpenXML output on this platform."
    Add-CompatibilityRow -Area "Windows launchers" -Status "unsupported" -Detail ".cmd and .vbs launchers do not run on macOS/Linux."
    Add-CompatibilityRow -Area "CLI core" -Status "partial" -Detail "Config parsing, crawling, and OpenXML XLSX output can run in PowerShell 7; the current GUI remains Windows-only."
}

$rows | Format-Table -AutoSize

if ($issues.Count -gt 0) {
    Write-Host "Environment compatibility check failed:"
    foreach ($issue in @($issues.ToArray())) {
        Write-Host "- $issue"
    }
    exit 1
}

Write-Host "Environment compatibility check completed."
