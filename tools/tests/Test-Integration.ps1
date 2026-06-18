[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
. (Join-Path $projectRoot "app\core\JobTracker.Common.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Config.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Context.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Runtime.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.OutputMaintenance.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.SourceAdapter.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Scoring.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Deduplication.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Excel.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Pipeline.ps1")
$script:LoadedSourceAdapters = @(Get-JobCrawlerSourceAdapterFiles -SourcesRoot (Join-Path $projectRoot "app\sources"))
foreach ($sourceAdapter in $script:LoadedSourceAdapters) {
    . $sourceAdapter.Path
}

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

Assert-Integration -Condition ($script:JobCrawlerConfig.Profile.Id -eq "digital_analytics") -Message "Expected Digital Analytics to remain the default profile."
Assert-Integration -Condition (@(Get-ConfigStringArray (Get-ConfigPathValue -Object $script:JobCrawlerSourcesConfig -Path "queries.linkedin" -DefaultValue @())).Count -gt 0) -Message "Expected profile-level LinkedIn queries to merge into sources config."
Assert-Integration -Condition (@(Get-ConfigPathValue -Object $script:JobCrawlerMatchingRules -Path "positive_signals" -DefaultValue @()).Count -gt 0) -Message "Expected profile-level positive matching signals."

$sources = @(Get-JobCrawlerSourceDefinitions -SourcesConfig $script:JobCrawlerSourcesConfig)
$sourceContract = Test-JobCrawlerSourceContract -SourceDefinitions $sources -LoadedFiles $script:LoadedSourceAdapters
Assert-Integration -Condition $sourceContract.IsValid -Message "Expected configured source functions to be loaded dynamically."
Assert-Integration -Condition (@($sources | Where-Object { $_.Key -eq "linkedin" -and $_.CrawlFunction -eq "Get-LinkedInJobs" }).Count -eq 1) -Message "Expected LinkedIn source registry metadata."
Assert-Integration -Condition (@($sources | Where-Object { $_.Key -eq "wttj_public" -and $_.FallbackFor -eq "welcome_kit" }).Count -eq 1) -Message "Expected WTTJ fallback relationship in source registry."
Assert-Integration -Condition (@($sources | Where-Object { $_.Key -eq "wttj_public" -and $_.SkipSwitch -eq "DisableWttjPublicFallback" }).Count -eq 1) -Message "Expected WTTJ public fallback to have its own skip switch."
Assert-Integration -Condition (@($sources | Where-Object { $_.Key -eq "welcome_kit" -and $_.SkipSwitch -eq "DisableWelcomeKit" }).Count -eq 1) -Message "Expected WelcomeKit to have its own skip switch."

$customSourcesConfig = [PSCustomObject]@{
    source_order = @("custom_board")
    sources      = [PSCustomObject]@{
        custom_board = [PSCustomObject]@{
            label                = "Custom board"
            short_label          = "Custom"
            enabled_by_default   = $true
            requires_credentials = $false
            crawl_function       = "Get-CustomJobs"
        }
    }
}
$customSources = @(Get-JobCrawlerSourceDefinitions -SourcesConfig $customSourcesConfig)
Assert-Integration -Condition ($customSources.Count -eq 1 -and $customSources[0].Key -eq "custom_board" -and $customSources[0].CrawlFunction -eq "Get-CustomJobs") -Message "Expected custom config-defined source metadata without code changes."

$match = Get-JobMatch -Title "Web Analyst" -Text "Google Tag Manager GA4 ContentSquare dataLayer"
$rows = @(
    (New-JobResult -Title "Web Analyst" -CompanyName "Radio France" -JobLocation "Paris" -ContractType "CDI" -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url "https://www.linkedin.com/jobs/view/222" -Platform "LinkedIn" -PublishedAt ([DateTimeOffset]::Now) -SourceText "GA4 Google Tag Manager"),
    (New-JobResult -Title "Web Analyst H/F" -CompanyName "Radio France" -JobLocation "75 - Paris" -ContractType "CDI" -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url "https://www.hellowork.com/fr-fr/emplois/222.html" -Platform "HelloWork" -PublishedAt ([DateTimeOffset]::Now) -SourceText "GA4 Google Tag Manager")
)
$merge = Merge-JobsWithTracker -CurrentRows $rows -ExistingRows @() -Path "integration.xlsx" -SkipBackup
Assert-Integration -Condition (@($merge.TrackerRows).Count -eq 1) -Message "Expected similar cross-platform rows to merge."
Assert-Integration -Condition ((Get-RowValue -Row $merge.TrackerRows[0] -Name "source_count") -eq "2") -Message "Expected merged source count to be 2."

$oldCurrentRow = New-JobResult -Title "Web Analyst" -CompanyName "Old Company" -JobLocation "Paris" -ContractType "CDI" -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url "https://example.test/old-job" -Platform "Test" -PublishedAt ([DateTimeOffset]::Now.AddDays(-8)) -SourceText "GA4 Google Tag Manager"
$oldCurrentMerge = Merge-JobsWithTracker -CurrentRows @($oldCurrentRow) -ExistingRows @() -Path "integration.xlsx" -SkipBackup
Assert-Integration -Condition (@($oldCurrentMerge.TrackerRows).Count -eq 0 -and [int]$oldCurrentMerge.RemovedCount -eq 0) -Message "Expected current rows outside the published-date retention window to be rejected at merge time without counting as removed tracker rows."

$existingCddRow = New-OrderedJobRecord @{
    status         = "ignored"
    job_title      = "Web Analyst"
    company_name   = "CDD Company"
    location       = "Paris"
    contract_type  = "CDD"
    match_score    = "80"
    match_level    = "High"
    job_url_raw    = "https://example.test/cdd-existing"
    platform       = "LinkedIn"
    published_date = ([DateTimeOffset]::Now.ToString("yyyy-MM-dd"))
}
$currentBlankContractRow = New-JobResult -Title "Web Analyst" -CompanyName "CDD Company" -JobLocation "Paris" -ContractType "" -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url "https://example.test/cdd-current" -Platform "Test" -PublishedAt ([DateTimeOffset]::Now) -SourceText "GA4 Google Tag Manager"
$excludedContractMerge = Merge-JobsWithTracker -CurrentRows @($currentBlankContractRow) -ExistingRows @($existingCddRow) -Path "integration.xlsx" -SkipBackup
Assert-Integration -Condition (@($excludedContractMerge.TrackerRows).Count -eq 0 -and [int]$excludedContractMerge.RemovedCount -eq 1) -Message "Expected excluded existing contract values not to leak back into current non-application rows."

$nextonDuplicateRows = @(
    (New-OrderedJobRecord @{
        status         = "ignored"
        job_title      = "Data Analyst Marketing H/F - NEXTON - CDI à Lyon"
        company_name   = "Nexton Consulting"
        location       = "Paris, FR"
        contract_type  = "CDI"
        match_score    = "52"
        match_level    = "Medium"
        job_url_raw    = "https://www.welcometothejungle.com/fr/companies/nexton-consulting/jobs/data-analyst-senior-digital-analytics-h-f_lyon"
        platform       = "Welcome to the Jungle; LinkedIn"
        published_date = ([DateTimeOffset]::Now.ToString("yyyy-MM-dd"))
    }),
    (New-OrderedJobRecord @{
        status         = "ignored"
        job_title      = "Data Analyst Senior Digital Analytics H F"
        company_name   = "NEXTON"
        location       = "Lyon"
        contract_type  = "CDI"
        match_score    = "55"
        match_level    = "Medium"
        job_url_raw    = "https://www.welcometothejungle.com/fr/companies/nexton-consulting/jobs/data-analyst-senior-digital-analytics-h-f_lyon?utm_source=test"
        platform       = "Welcome to the Jungle"
        published_date = ([DateTimeOffset]::Now.ToString("yyyy-MM-dd"))
    })
)
$nextonMerge = Merge-JobsWithTracker -CurrentRows @() -ExistingRows $nextonDuplicateRows -Path "integration.xlsx" -SkipBackup
Assert-Integration -Condition (@($nextonMerge.TrackerRows).Count -eq 1 -and [int]$nextonMerge.DuplicateCount -eq 1) -Message "Expected exact canonical URL duplicates to merge even when titles, company labels, and locations differ."

$olivierRows = @(
    (New-JobResult -Title "Web Analyst CRO" -CompanyName "L'Olivier Assurance" -JobLocation "Paris" -ContractType "CDI" -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url "https://fr.linkedin.com/jobs/view/web-analyst-cro-at-lolivier-assurance-111" -Platform "LinkedIn" -PublishedAt ([DateTimeOffset]::Now) -SourceText "GA4 Google Tag Manager CRO"),
    (New-JobResult -Title "Web Analyst CRO H/F" -CompanyName "Olivier" -JobLocation "Paris, Ile-de-France" -ContractType "CDI" -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url "https://www.welcometothejungle.com/fr/companies/olivier/jobs/web-analyst-cro_paris" -Platform "Welcome to the Jungle" -PublishedAt ([DateTimeOffset]::Now) -SourceText "GA4 Google Tag Manager CRO")
)
$olivierMerge = Merge-JobsWithTracker -CurrentRows $olivierRows -ExistingRows @() -Path "integration.xlsx" -SkipBackup
Assert-Integration -Condition (@($olivierMerge.TrackerRows).Count -eq 1 -and (Get-RowValue -Row $olivierMerge.TrackerRows[0] -Name "platform") -match "LinkedIn" -and (Get-RowValue -Row $olivierMerge.TrackerRows[0] -Name "platform") -match "Welcome to the Jungle") -Message "Expected company alias hierarchy to merge L'Olivier Assurance and Olivier when title and location also match."

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

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ("custom-job-tracker-integration-{0}" -f ([Guid]::NewGuid().ToString("N")))
New-Item -ItemType Directory -Path $tempRoot | Out-Null
try {
    $oldFile = Join-Path $tempRoot "old.txt"
    $newFile = Join-Path $tempRoot "new.txt"
    Set-Content -LiteralPath $oldFile -Value "old" -Encoding UTF8
    Set-Content -LiteralPath $newFile -Value "new" -Encoding UTF8
    (Get-Item -LiteralPath $oldFile).LastWriteTime = (Get-Date).AddDays(-60)
    $prune = Invoke-JobCrawlerCachePrune -Path $tempRoot -Enabled:$true
    Assert-Integration -Condition ([int]$prune.RemovedFiles -ge 1 -and -not (Test-Path -LiteralPath $oldFile) -and (Test-Path -LiteralPath $newFile)) -Message "Expected cache prune to remove old cache files only."

    $managedRoot = Join-Path $tempRoot "project"
    $managedCache = Join-Path $managedRoot "output\cache"
    New-Item -ItemType Directory -Path $managedCache -Force | Out-Null
    $managedOld = Join-Path $managedCache "old-cache.txt"
    $managedNew = Join-Path $managedCache "new-cache.txt"
    Set-Content -LiteralPath $managedOld -Value "old" -Encoding UTF8
    Set-Content -LiteralPath $managedNew -Value "new" -Encoding UTF8
    (Get-Item -LiteralPath $managedOld).LastWriteTime = (Get-Date).AddDays(-20)
    $cleanupRows = @(Invoke-JobCrawlerOutputCleanup -ProjectRoot $managedRoot -CacheDirectory $managedCache -Cache -OlderThanDays 14)
    Assert-Integration -Condition (($cleanupRows | Measure-Object RemovedFiles -Sum).Sum -eq 1 -and -not (Test-Path -LiteralPath $managedOld) -and (Test-Path -LiteralPath $managedNew)) -Message "Expected managed output cleanup to remove old cache files inside the project root only."

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
