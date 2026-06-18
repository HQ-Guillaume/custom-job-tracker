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
$script:Cutoff = [DateTimeOffset]::Now.AddDays(-7)
$script:CutoffDate = $script:Cutoff.ToString("yyyy-MM-dd")
$script:CrawlMode = "Test"
$script:Location = "France"
$script:CacheDirectory = Join-Path $projectRoot "output\cache"
$script:CacheTtlHours = 24
$script:MinimumMatchScore = 35

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ("custom-job-tracker-feedback-{0}" -f ([Guid]::NewGuid().ToString("N")))
New-Item -ItemType Directory -Path $tempRoot | Out-Null
try {
    $trackerPath = Join-Path $tempRoot "tracker.xlsx"
    $outputPath = Join-Path $tempRoot "suggestions.csv"
    $match = Get-JobMatch -Title "Web Analyst" -Text "GA4 Google Tag Manager"
    $rows = @(
        (New-OrderedJobRecord @{ status = "applied"; job_title = "Web Analyst"; company_name = "A"; location = "Paris"; contract_type = "CDI"; match_score = $match.Score; match_level = $match.Level; matched_keywords = $match.Keywords; job_url_raw = "https://example.test/a"; platform = "Test"; published_date = (Get-Date -Format "yyyy-MM-dd") }),
        (New-OrderedJobRecord @{ status = "ignored"; notes = "ignore_reason=agency_consulting_esn; detail="; job_title = "Consultant Digital Analytics"; company_name = "B"; location = "Paris"; contract_type = "CDI"; match_score = $match.Score; match_level = $match.Level; matched_keywords = $match.Keywords; job_url_raw = "https://example.test/b"; platform = "Test"; published_date = (Get-Date -Format "yyyy-MM-dd") }),
        (New-OrderedJobRecord @{ status = "ignored"; notes = "ignore_reason=agency_consulting_esn; detail="; job_title = "Consultant Tracking"; company_name = "C"; location = "Paris"; contract_type = "CDI"; match_score = $match.Score; match_level = $match.Level; matched_keywords = $match.Keywords; job_url_raw = "https://example.test/c"; platform = "Test"; published_date = (Get-Date -Format "yyyy-MM-dd") })
    )
    Export-TrackerWorkbook -Rows $rows -Path $trackerPath -Summary @{ Profile = "Test"; TotalMatched = 3; CurrentCount = 3; TrackerCount = 3 }
    & (Join-Path $projectRoot "tools\diagnostics\Get-FeedbackTuningSuggestions.ps1") -TrackerPath $trackerPath -OutputPath $outputPath
    if (-not (Test-Path -LiteralPath $outputPath)) {
        throw "Expected feedback suggestions CSV."
    }
    $csv = Import-Csv -LiteralPath $outputPath
    if (@($csv | Where-Object { $_.Area -eq "Employer preference" }).Count -eq 0) {
        throw "Expected employer preference suggestion."
    }
}
finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "Feedback suggestion tests passed."
