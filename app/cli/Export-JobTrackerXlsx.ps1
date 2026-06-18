[CmdletBinding()]
param(
    [string]$TrackerPath = "",
    [string]$ConfigDirectory = "config",
    [string]$Profile = ""
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$CoreRoot = Join-Path $ProjectRoot "app\core"

. (Join-Path $CoreRoot "JobTracker.Common.ps1")
. (Join-Path $CoreRoot "JobTracker.Config.ps1")
. (Join-Path $CoreRoot "JobTracker.Runtime.ps1")
. (Join-Path $CoreRoot "JobTracker.Scoring.ps1")
. (Join-Path $CoreRoot "JobTracker.Deduplication.ps1")
. (Join-Path $CoreRoot "JobTracker.Excel.ps1")
. (Join-Path $CoreRoot "JobTracker.Pipeline.ps1")

$configPath = Resolve-JobCrawlerPath -BasePath $ProjectRoot -Path $ConfigDirectory
$JobCrawlerConfig = Get-JobCrawlerConfig -ConfigDirectory $configPath -ProfileId $Profile
$JobCrawlerRuntimeConfig = $JobCrawlerConfig.Runtime
$JobCrawlerSourcesConfig = $JobCrawlerConfig.Sources
$JobCrawlerMatchingRules = $JobCrawlerConfig.MatchingRules
$JobCrawlerWorkbookConfig = $JobCrawlerConfig.Workbook

$validation = Test-JobCrawlerConfig -Config $JobCrawlerConfig
if (-not $validation.IsValid) {
    throw ("Invalid crawler config:`n- {0}" -f (($validation.Issues) -join "`n- "))
}

if ([string]::IsNullOrWhiteSpace($TrackerPath)) {
    $TrackerPath = Resolve-JobCrawlerPath -BasePath $ProjectRoot -Path ([string](Get-ConfigPathValue -Object $JobCrawlerRuntimeConfig -Path "defaults.tracker_path" -DefaultValue "output\jobs_tracker.xlsx"))
}
if (-not (Test-Path -LiteralPath $TrackerPath)) {
    throw "Tracker workbook not found: $TrackerPath"
}

$RunStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$RunDate = Get-Date -Format "yyyy-MM-dd"
$DaysBack = [int](Get-ConfigPathValue -Object $JobCrawlerRuntimeConfig -Path "defaults.days_back" -DefaultValue 7)
$Cutoff = [DateTimeOffset]::Now.AddDays(-[Math]::Abs($DaysBack))
$CutoffDate = $Cutoff.ToString("yyyy-MM-dd")
$CrawlMode = [string](Get-ConfigPathValue -Object $JobCrawlerRuntimeConfig -Path "defaults.crawl_mode" -DefaultValue "Default")
$Location = [string](Get-ConfigPathValue -Object $JobCrawlerRuntimeConfig -Path "defaults.location" -DefaultValue "France")
$CacheDirectory = Resolve-JobCrawlerPath -BasePath $ProjectRoot -Path ([string](Get-ConfigPathValue -Object $JobCrawlerRuntimeConfig -Path "defaults.cache_directory" -DefaultValue "output\cache"))
$CacheTtlHours = [int](Get-ConfigPathValue -Object $JobCrawlerRuntimeConfig -Path "defaults.cache_ttl_hours" -DefaultValue 24)
$MinimumMatchScore = [int](Get-ConfigPathValue -Object $JobCrawlerMatchingRules -Path "thresholds.minimum_match_score" -DefaultValue 35)
$LinkedInQueries = @(Get-JobCrawlerSourceQueryList -SourcesConfig $JobCrawlerSourcesConfig -SourceKey "linkedin" -FallbackKeys @("api"))
$HelloWorkQueries = @(Get-JobCrawlerSourceQueryList -SourcesConfig $JobCrawlerSourcesConfig -SourceKey "hellowork" -FallbackKeys @("api"))
$ApecQueries = @(Get-JobCrawlerSourceQueryList -SourcesConfig $JobCrawlerSourcesConfig -SourceKey "apec" -FallbackKeys @("api"))
$FranceTravailQueries = @(Get-JobCrawlerSourceQueryList -SourcesConfig $JobCrawlerSourcesConfig -SourceKey "france_travail" -FallbackKeys @("api"))
$AdzunaQueries = @(Get-JobCrawlerSourceQueryList -SourcesConfig $JobCrawlerSourcesConfig -SourceKey "adzuna" -FallbackKeys @("api"))
$ApiSearchQueries = @(Get-ConfigStringArray (Get-ConfigPathValue -Object $JobCrawlerSourcesConfig -Path "queries.api" -DefaultValue @()))
$SourceRunStats = New-Object System.Collections.Generic.List[object]
$JobCrawlerPreferences = Get-JobCrawlerPreferences
$MasterColumns = Get-JobTrackerMasterColumns
$ColumnLabels = Get-JobTrackerColumnLabels

$rows = @(Import-TrackerRows -Path $TrackerPath)
$pipelineValidation = Test-JobPipelineInvariants -Rows $rows -Stage "reformat_existing"
if (-not $pipelineValidation.IsValid) {
    Write-Warning ("Existing tracker contains {0} non-application row(s) that fail current pipeline rules. Reformatting is kept non-destructive." -f @($pipelineValidation.Issues).Count)
}
$summary = @{
    Profile = ("{0} ({1})" -f $JobCrawlerConfig.Profile.Label, $JobCrawlerConfig.Profile.Id)
    TotalMatched = @($rows).Count
    ExcludedContractCount = 0
    CurrentCount = @($rows | Where-Object { (Get-RowValue -Row $_ -Name "seen_in_current_crawl") -eq "yes" }).Count
    TrackerCount = @($rows).Count
    DuplicateCount = 0
    RemovedCount = 0
    PreservedAppliedCount = @($rows | Where-Object { Test-IsAppliedStatus (Get-RowValue -Row $_ -Name "status") }).Count
    SourceDiagnostics = "Reformatted existing workbook; no crawl run."
    BackupPath = ""
    CrawlCaps = "Reformat only"
    DryRun = "no"
    DiagnosticMode = "no"
    DiagnosticPath = ""
}

Export-TrackerWorkbook -Rows $rows -Path $TrackerPath -Summary $summary
Write-Host ("Formatted workbook: {0}" -f (Resolve-Path $TrackerPath).Path)
