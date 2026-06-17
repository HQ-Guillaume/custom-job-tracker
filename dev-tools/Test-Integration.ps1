[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $projectRoot "JobTracker.Common.ps1")
. (Join-Path $projectRoot "app\JobTracker.Config.ps1")
. (Join-Path $projectRoot "app\JobTracker.Runtime.ps1")
. (Join-Path $projectRoot "app\JobTracker.Scoring.ps1")
. (Join-Path $projectRoot "app\JobTracker.Deduplication.ps1")
. (Join-Path $projectRoot "app\JobTracker.Excel.ps1")
. (Join-Path $projectRoot "app\sources\Source.Wttj.ps1")

function Assert-Integration {
    param(
        [bool]$Condition,
        [string]$Message
    )

    if (-not $Condition) {
        throw "Integration test failed: $Message"
    }
}

$script:JobCrawlerConfig = Get-JobCrawlerConfig -ConfigDirectory (Join-Path $projectRoot "config")
$script:JobCrawlerRuntimeConfig = $script:JobCrawlerConfig.Runtime
$script:JobCrawlerSourcesConfig = $script:JobCrawlerConfig.Sources
$script:JobCrawlerMatchingRules = $script:JobCrawlerConfig.MatchingRules
$script:JobCrawlerWorkbookConfig = $script:JobCrawlerConfig.Workbook
$script:JobCrawlerPreferences = Get-JobCrawlerPreferences
$script:MasterColumns = Get-JobTrackerMasterColumns
$script:ColumnLabels = Get-JobTrackerColumnLabels
$script:SeenResultKeys = @{}
$script:FeedbackLearningProfile = $null
$script:SourceRunStats = New-Object System.Collections.Generic.List[object]
$script:DisableCache = $false
$script:CacheTtlHours = 24
$script:Location = "France"
$script:CrawlMode = "Default"
$script:RunDate = Get-Date -Format "yyyy-MM-dd"
$script:RunStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$script:Cutoff = [DateTimeOffset]::Now.AddDays(-7)
$script:CutoffDate = $script:Cutoff.ToString("yyyy-MM-dd")
$script:MinimumMatchScore = [int](Get-ConfigPathValue -Object $script:JobCrawlerMatchingRules -Path "thresholds.minimum_match_score" -DefaultValue 35)

$sources = @(Get-JobCrawlerSourceDefinitions -SourcesConfig $script:JobCrawlerSourcesConfig)
Assert-Integration -Condition (@($sources | Where-Object { $_.Key -eq "linkedin" -and $_.CrawlFunction -eq "Get-LinkedInJobs" }).Count -eq 1) -Message "Expected LinkedIn source registry metadata."
Assert-Integration -Condition (@($sources | Where-Object { $_.Key -eq "wttj_public" -and $_.FallbackFor -eq "welcome_kit" }).Count -eq 1) -Message "Expected WTTJ fallback relationship in source registry."
Assert-Integration -Condition (@($sources | Where-Object { $_.Key -eq "wttj_public" -and $_.SkipSwitch -eq "DisableWttjPublicFallback" }).Count -eq 1) -Message "Expected WTTJ public fallback to have its own skip switch."
Assert-Integration -Condition (@($sources | Where-Object { $_.Key -eq "welcome_kit" -and $_.SkipSwitch -eq "DisableWelcomeKit" }).Count -eq 1) -Message "Expected WelcomeKit to have its own skip switch."

$match = Get-JobMatch -Title "Web Analyst" -Text "Google Tag Manager GA4 ContentSquare dataLayer"
$rows = @(
    (New-JobResult -Title "Web Analyst" -CompanyName "Radio France" -JobLocation "Paris" -ContractType "CDI" -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url "https://www.linkedin.com/jobs/view/222" -Platform "LinkedIn" -PublishedAt ([DateTimeOffset]::Now) -SourceText "GA4 Google Tag Manager"),
    (New-JobResult -Title "Web Analyst H/F" -CompanyName "Radio France" -JobLocation "75 - Paris" -ContractType "CDI" -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url "https://www.hellowork.com/fr-fr/emplois/222.html" -Platform "HelloWork" -PublishedAt ([DateTimeOffset]::Now) -SourceText "GA4 Google Tag Manager")
)
$merge = Merge-JobsWithTracker -CurrentRows $rows -ExistingRows @() -Path "integration.xlsx" -SkipBackup
Assert-Integration -Condition (@($merge.TrackerRows).Count -eq 1) -Message "Expected similar cross-platform rows to merge."
Assert-Integration -Condition ((Get-RowValue -Row $merge.TrackerRows[0] -Name "source_count") -eq "2") -Message "Expected merged source count to be 2."

$foreignExistingWttjRow = New-OrderedJobRecord @{
    status         = "ignored"
    job_title      = "Head Of Uk Sports Gtm"
    company_name   = "Example Company"
    location       = "London"
    contract_type  = "CDI"
    match_score    = "80"
    match_level    = "Good"
    job_url_raw    = "https://www.welcometothejungle.com/en/companies/acme/jobs/head-of-uk-sports-gtm_london"
    platform       = "Welcome to the Jungle"
    published_date = ([DateTimeOffset]::Now.ToString("yyyy-MM-dd"))
}
$cleanupMerge = Merge-JobsWithTracker -CurrentRows @() -ExistingRows @($foreignExistingWttjRow) -Path "integration.xlsx" -SkipBackup
Assert-Integration -Condition ([int]$cleanupMerge.RemovedCount -eq 1) -Message "Expected invalid WTTJ existing row cleanup."

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ("job-crawler-integration-{0}" -f ([Guid]::NewGuid().ToString("N")))
New-Item -ItemType Directory -Path $tempRoot | Out-Null
try {
    $oldFile = Join-Path $tempRoot "old.txt"
    $newFile = Join-Path $tempRoot "new.txt"
    Set-Content -LiteralPath $oldFile -Value "old" -Encoding UTF8
    Set-Content -LiteralPath $newFile -Value "new" -Encoding UTF8
    (Get-Item -LiteralPath $oldFile).LastWriteTime = (Get-Date).AddDays(-60)
    $prune = Invoke-JobCrawlerCachePrune -Path $tempRoot -Enabled:$true
    Assert-Integration -Condition ([int]$prune.RemovedFiles -ge 1 -and -not (Test-Path -LiteralPath $oldFile) -and (Test-Path -LiteralPath $newFile)) -Message "Expected cache prune to remove old cache files only."

    $historyPath = Join-Path $tempRoot "run_history.jsonl"
    for ($i = 0; $i -lt 3; $i++) {
        Write-RunHistoryEntry -Path $historyPath -MaxEntries 2 -Summary @{
            DryRun = "yes"
            DiagnosticMode = "no"
            TotalMatched = $i
            CurrentCount = $i
            TrackerCount = $i
            DuplicateCount = 0
            RemovedCount = 0
            PreservedAppliedCount = 0
            ExcludedContractCount = 0
        }
    }
    Assert-Integration -Condition (@(Get-Content -LiteralPath $historyPath).Count -eq 2) -Message "Expected run history pruning to keep the configured max entries."
}
finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "Integration tests passed."
