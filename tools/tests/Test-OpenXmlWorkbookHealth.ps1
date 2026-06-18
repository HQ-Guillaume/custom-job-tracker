[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
. (Join-Path $projectRoot "app\core\JobTracker.Common.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Config.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Runtime.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Scoring.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Deduplication.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Excel.ps1")

$script:JobCrawlerConfig = Get-JobCrawlerConfig -ConfigDirectory (Join-Path $projectRoot "config")
$script:JobCrawlerRuntimeConfig = $script:JobCrawlerConfig.Runtime
$script:JobCrawlerSourcesConfig = $script:JobCrawlerConfig.Sources
$script:JobCrawlerMatchingRules = $script:JobCrawlerConfig.MatchingRules
$script:JobCrawlerWorkbookConfig = $script:JobCrawlerConfig.Workbook
$script:JobCrawlerWorkbookConfig.output_backend = "openxml"
$script:JobCrawlerPreferences = Get-JobCrawlerPreferences
$script:MasterColumns = Get-JobTrackerMasterColumns
$script:ColumnLabels = Get-JobTrackerColumnLabels
$script:SourceRunStats = New-Object System.Collections.Generic.List[object]
$script:FeedbackLearningProfile = $null
$script:SeenResultKeys = @{}
$script:RunStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$script:RunDate = Get-Date -Format "yyyy-MM-dd"
$script:DaysBack = 7
$script:Cutoff = [DateTimeOffset]::Now.AddDays(-7)
$script:CutoffDate = $script:Cutoff.ToString("yyyy-MM-dd")
$script:CrawlMode = "Test"
$script:Location = "France"
$script:CacheDirectory = Join-Path $projectRoot "output\cache"
$script:CacheTtlHours = 24
$script:MinimumMatchScore = 35

$tempPath = Join-Path ([IO.Path]::GetTempPath()) ("custom-job-tracker-health-test-{0}.xlsx" -f ([Guid]::NewGuid().ToString("N")))
try {
    $match = Get-JobMatch -Title "Web Analyst" -Text "GA4 Google Tag Manager"
    $row = New-JobResult -Title "Web Analyst" -CompanyName "Example" -JobLocation "Paris" -ContractType "CDI" -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url "https://example.test/job" -Platform "Test" -PublishedAt ([DateTimeOffset]::Now) -SourceText "GA4 Google Tag Manager"
    Export-TrackerWorkbook -Rows @($row) -Path $tempPath -Summary @{ Profile = "Test"; TotalMatched = 1; CurrentCount = 1; TrackerCount = 1 }
    & (Join-Path $projectRoot "tools\diagnostics\Test-WorkbookHealthOpenXml.ps1") -TrackerPath $tempPath
}
finally {
    Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
}

Write-Host "OpenXML workbook health tests passed."
