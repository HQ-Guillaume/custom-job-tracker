[CmdletBinding()]
param(
    [string]$ConfigDirectory = "config",
    [string]$Profile = "",
    [switch]$Cache,
    [switch]$Logs,
    [switch]$Diagnostics,
    [switch]$Backups,
    [switch]$All,
    [int]$OlderThanDays = -1,
    [switch]$WhatIf
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$CoreRoot = Join-Path $ProjectRoot "app\core"

. (Join-Path $CoreRoot "JobTracker.Common.ps1")
. (Join-Path $CoreRoot "JobTracker.Config.ps1")
. (Join-Path $CoreRoot "JobTracker.Runtime.ps1")
. (Join-Path $CoreRoot "JobTracker.OutputMaintenance.ps1")

$configPath = Resolve-JobCrawlerPath -BasePath $ProjectRoot -Path $ConfigDirectory
$script:JobCrawlerConfig = Get-JobCrawlerConfig -ConfigDirectory $configPath -ProfileId $Profile
$script:JobCrawlerRuntimeConfig = $script:JobCrawlerConfig.Runtime
$script:CacheDirectory = Resolve-JobCrawlerPath -BasePath $ProjectRoot -Path ([string](Get-ConfigPathValue -Object $script:JobCrawlerRuntimeConfig -Path "defaults.cache_directory" -DefaultValue "output\cache"))
$script:ProjectRoot = $ProjectRoot

if (-not ($Cache -or $Logs -or $Diagnostics -or $Backups -or $All)) {
    $outputDirectory = Split-Path -Parent (Resolve-JobCrawlerPath -BasePath $ProjectRoot -Path ([string](Get-ConfigPathValue -Object $script:JobCrawlerRuntimeConfig -Path "defaults.tracker_path" -DefaultValue "output\jobs_tracker.xlsx")))
    $stats = @(
        (Get-JobCrawlerDirectoryStats -Label "cache" -Path $script:CacheDirectory),
        (Get-JobCrawlerDirectoryStats -Label "launcher logs" -Path (Join-Path $outputDirectory "launcher_logs")),
        (Get-JobCrawlerDirectoryStats -Label "diagnostics" -Path (Resolve-JobCrawlerPath -BasePath $ProjectRoot -Path ([string](Get-ConfigPathValue -Object $script:JobCrawlerRuntimeConfig -Path "output.diagnostics_directory" -DefaultValue "output\diagnostics")))),
        (Get-JobCrawlerDirectoryStats -Label "backups" -Path (Join-Path $outputDirectory "backups"))
    )
    $stats | Format-Table -AutoSize
    Write-Host ""
    Write-Host "No cleanup switch was provided. Add -Cache, -Logs, -Diagnostics, -Backups, or -All to remove old managed files."
    return
}

$results = Invoke-JobCrawlerOutputCleanup -ProjectRoot $ProjectRoot -CacheDirectory $script:CacheDirectory -Cache:$Cache -Logs:$Logs -Diagnostics:$Diagnostics -Backups:$Backups -All:$All -OlderThanDays $OlderThanDays -WhatIf:$WhatIf
$results | Format-Table -AutoSize
