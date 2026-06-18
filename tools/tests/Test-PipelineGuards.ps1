[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
. (Join-Path $projectRoot "app\core\JobTracker.Common.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Config.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Context.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Runtime.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.SourceAdapter.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Scoring.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Deduplication.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Excel.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Pipeline.ps1")
$script:LoadedSourceAdapters = @(Get-JobCrawlerSourceAdapterFiles -SourcesRoot (Join-Path $projectRoot "app\sources"))
foreach ($sourceAdapter in $script:LoadedSourceAdapters) {
    . $sourceAdapter.Path
}

function Assert-Pipeline {
    param(
        [bool]$Condition,
        [string]$Message
    )

    if (-not $Condition) {
        throw "Pipeline guard test failed: $Message"
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
$script:Location = "France"
$script:CrawlMode = "Default"
$script:RunDate = Get-Date -Format "yyyy-MM-dd"
$script:RunStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$script:Cutoff = [DateTimeOffset]::Now.AddDays(-7)
$script:CutoffDate = $script:Cutoff.ToString("yyyy-MM-dd")
$script:MinimumMatchScore = [int](Get-ConfigPathValue -Object $script:JobCrawlerMatchingRules -Path "thresholds.minimum_match_score" -DefaultValue 35)

$match = Get-JobMatch -Title "Web Analyst" -Text "Google Analytics GA4 Google Tag Manager"
$recentRow = New-JobResult -Title "Web Analyst" -CompanyName "Recent Company" -JobLocation "Paris" -ContractType "CDI" -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url "https://example.test/recent" -Platform "LinkedIn" -PublishedAt ([DateTimeOffset]::Now) -SourceText "Google Analytics GA4 Google Tag Manager"
$oldRow = New-JobResult -Title "Web Analyst" -CompanyName "Old Company" -JobLocation "Paris" -ContractType "CDI" -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url "https://example.test/old" -Platform "Welcome to the Jungle" -PublishedAt ([DateTimeOffset]::Now.AddDays(-8)) -SourceText "Google Analytics GA4 Google Tag Manager"
$cddRow = New-JobResult -Title "Web Analyst CDD" -CompanyName "CDD Company" -JobLocation "Paris" -ContractType "CDD" -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url "https://example.test/cdd" -Platform "France Travail" -PublishedAt ([DateTimeOffset]::Now) -SourceText "Google Analytics GA4 Google Tag Manager CDD"
$foreignWttjRow = New-JobResult -Title "Web Analyst" -CompanyName "Foreign Company" -JobLocation "London" -ContractType "CDI" -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url "https://www.welcometothejungle.com/en/companies/acme/jobs/web-analyst_london" -Platform "Welcome to the Jungle" -PublishedAt ([DateTimeOffset]::Now) -SourceText "Google Analytics GA4 Google Tag Manager"

$gate = Invoke-JobPipelineEligibilityGate -Rows @($recentRow, $oldRow, $cddRow, $foreignWttjRow) -Stage "post_enrichment"
Assert-Pipeline -Condition ($gate.KeptCount -eq 1) -Message "Expected only the recent eligible row to pass the post-enrichment gate."
Assert-Pipeline -Condition ((Get-JobPipelineReasonCount -GateResult $gate -Reason "outside_published_window") -eq 1) -Message "Expected old rows to be excluded by published-date rule."
Assert-Pipeline -Condition ((Get-JobPipelineReasonCount -GateResult $gate -Reason "excluded_contract") -eq 1) -Message "Expected CDD rows to be excluded by contract rule."
Assert-Pipeline -Condition ((Get-JobPipelineReasonCount -GateResult $gate -Reason "invalid_location") -eq 1) -Message "Expected foreign WTTJ rows to be excluded by location rule."

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
$currentBlankContractRow = New-JobResult -Title "Web Analyst" -CompanyName "CDD Company" -JobLocation "Paris" -ContractType "" -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url "https://example.test/cdd-current" -Platform "Test" -PublishedAt ([DateTimeOffset]::Now) -SourceText "Google Analytics GA4 Google Tag Manager"
$trackerRecord = ConvertTo-TrackerRecord -CurrentRow $currentBlankContractRow -ExistingRow $existingCddRow -SeenInCurrentCrawl:$true
$decision = Get-JobPipelineEligibility -Row $trackerRecord -CurrentRow $currentBlankContractRow -ExistingRow $existingCddRow -Stage "merge_current_final"
Assert-Pipeline -Condition (-not $decision.IsEligible -and $decision.Reason -eq "excluded_contract") -Message "Expected merge gate to reject rows whose excluded contract came from existing tracker history."

$appliedOldCddRow = New-OrderedJobRecord @{
    status         = "applied"
    job_title      = "Web Analyst"
    company_name   = "Applied Company"
    location       = "Paris"
    contract_type  = "CDD"
    match_score    = "80"
    match_level    = "High"
    job_url_raw    = "https://example.test/applied-cdd"
    platform       = "LinkedIn"
    published_date = ([DateTimeOffset]::Now.AddDays(-30).ToString("yyyy-MM-dd"))
}
$appliedDecision = Get-JobPipelineEligibility -Row $appliedOldCddRow -Stage "pre_export"
Assert-Pipeline -Condition ($appliedDecision.IsEligible -and $appliedDecision.KeepForever) -Message "Expected application history rows to be preserved even when old or excluded by current crawl rules."

$invariant = Test-JobPipelineInvariants -Rows @($recentRow, $cddRow) -Stage "test"
Assert-Pipeline -Condition (-not $invariant.IsValid -and @($invariant.Issues).Count -eq 1 -and $invariant.Issues[0].Reason -eq "excluded_contract") -Message "Expected pre-export invariant check to detect invalid non-application rows."

Write-Host "Pipeline guard tests passed."
