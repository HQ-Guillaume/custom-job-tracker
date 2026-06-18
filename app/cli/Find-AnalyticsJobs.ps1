[CmdletBinding()]
param(
    [int]$DaysBack = 7,
    [string]$Location = "France",
    [string]$TrackerPath = "",
    [string]$WelcomeKitApiKey = $env:WK_API_KEY,
    [string]$FranceTravailClientId = $env:FRANCE_TRAVAIL_CLIENT_ID,
    [string]$FranceTravailClientSecret = $env:FRANCE_TRAVAIL_CLIENT_SECRET,
    [string]$FranceTravailScope = $(if ([string]::IsNullOrWhiteSpace($env:FRANCE_TRAVAIL_SCOPE)) { "api_offresdemploiv2 o2dsoffre" } else { $env:FRANCE_TRAVAIL_SCOPE }),
    [string]$AdzunaAppId = $env:ADZUNA_APP_ID,
    [string]$AdzunaAppKey = $env:ADZUNA_APP_KEY,
    [string]$CrawlMode = "Default",
    [int]$MaxLinkedInSearchPages = 3,
    [int]$MaxLinkedInDetails = 0,
    [int]$MaxFranceTravailPages = 2,
    [int]$MaxAdzunaPages = 1,
    [int]$MaxApecPages = 2,
    [int]$MaxHelloWorkPages = 1,
    [int]$MaxHelloWorkCardsPerQuery = 20,
    [int]$MaxHelloWorkDetails = 50,
    [int]$MaxWelcomeKitPages = 10,
    [int]$MaxWttjCandidatePages = 120,
    [int]$MaxBackups = 5,
    [string]$Profile = "",
    [string[]]$EnableSource = @(),
    [string[]]$SkipSource = @(),
    [switch]$SkipFranceTravail,
    [switch]$SkipAdzuna,
    [switch]$SkipApec,
    [switch]$SkipHelloWork,
    [switch]$SkipLinkedIn,
    [switch]$SkipWttj,
    [switch]$EnableFranceTravail,
    [switch]$EnableAdzuna,
    [switch]$EnableWelcomeKit,
    [switch]$DisableWelcomeKit,
    [switch]$DisableWttjPublicFallback,
    [switch]$DisableCache,
    [int]$CacheTtlHours = 24,
    [string]$ConfigDirectory = "config",
    [switch]$DryRun,
    [switch]$DiagnosticMode,
    [switch]$ValidateConfig,
    [switch]$SelfTest
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$CoreRoot = Join-Path $ProjectRoot "app\core"
$SourcesRoot = Join-Path $ProjectRoot "app\sources"

. (Join-Path $CoreRoot "JobTracker.Common.ps1")
. (Join-Path $CoreRoot "JobTracker.Config.ps1")
. (Join-Path $CoreRoot "JobTracker.Context.ps1")
. (Join-Path $CoreRoot "JobTracker.Runtime.ps1")
. (Join-Path $CoreRoot "JobTracker.SourceAdapter.ps1")
. (Join-Path $CoreRoot "JobTracker.Scoring.ps1")
. (Join-Path $CoreRoot "JobTracker.Deduplication.ps1")
. (Join-Path $CoreRoot "JobTracker.Excel.ps1")
. (Join-Path $CoreRoot "JobTracker.Pipeline.ps1")
. (Join-Path $CoreRoot "JobTracker.SelfTest.ps1")

$LoadedSourceAdapters = @(Get-JobCrawlerSourceAdapterFiles -SourcesRoot $SourcesRoot)
foreach ($sourceAdapter in $LoadedSourceAdapters) {
    . $sourceAdapter.Path
}

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$configPath = Resolve-JobCrawlerPath -BasePath $ProjectRoot -Path $ConfigDirectory
$JobCrawlerConfig = Get-JobCrawlerConfig -ConfigDirectory $configPath -ProfileId $Profile
$JobCrawlerRuntimeConfig = $JobCrawlerConfig.Runtime
$JobCrawlerSourcesConfig = $JobCrawlerConfig.Sources
$JobCrawlerMatchingRules = $JobCrawlerConfig.MatchingRules
$JobCrawlerWorkbookConfig = $JobCrawlerConfig.Workbook

$configValidation = Test-JobCrawlerConfig -Config $JobCrawlerConfig
if (-not $configValidation.IsValid) {
    throw ("Invalid crawler config:`n- {0}" -f (($configValidation.Issues) -join "`n- "))
}
if ($ValidateConfig) {
    Write-Host "Crawler config validation passed."
    return
}

if (-not $PSBoundParameters.ContainsKey("DaysBack")) { $DaysBack = [int](Get-ConfigPathValue -Object $JobCrawlerRuntimeConfig -Path "defaults.days_back" -DefaultValue 7) }
if (-not $PSBoundParameters.ContainsKey("Location")) { $Location = [string](Get-ConfigPathValue -Object $JobCrawlerRuntimeConfig -Path "defaults.location" -DefaultValue "France") }
if (-not $PSBoundParameters.ContainsKey("CrawlMode")) { $CrawlMode = [string](Get-ConfigPathValue -Object $JobCrawlerRuntimeConfig -Path "defaults.crawl_mode" -DefaultValue "Default") }
if (-not $PSBoundParameters.ContainsKey("MaxBackups")) { $MaxBackups = [int](Get-ConfigPathValue -Object $JobCrawlerRuntimeConfig -Path "defaults.max_backups" -DefaultValue 5) }
if (-not $PSBoundParameters.ContainsKey("CacheTtlHours")) { $CacheTtlHours = [int](Get-ConfigPathValue -Object $JobCrawlerRuntimeConfig -Path "defaults.cache_ttl_hours" -DefaultValue 24) }

if (-not $PSBoundParameters.ContainsKey("WelcomeKitApiKey")) { $WelcomeKitApiKey = Get-JobCrawlerCredentialValue -SourcesConfig $JobCrawlerSourcesConfig -SourceKey "welcome_kit" -CredentialKey "api_key" -FallbackValue $WelcomeKitApiKey }
if (-not $PSBoundParameters.ContainsKey("FranceTravailClientId")) { $FranceTravailClientId = Get-JobCrawlerCredentialValue -SourcesConfig $JobCrawlerSourcesConfig -SourceKey "france_travail" -CredentialKey "client_id" -FallbackValue $FranceTravailClientId }
if (-not $PSBoundParameters.ContainsKey("FranceTravailClientSecret")) { $FranceTravailClientSecret = Get-JobCrawlerCredentialValue -SourcesConfig $JobCrawlerSourcesConfig -SourceKey "france_travail" -CredentialKey "client_secret" -FallbackValue $FranceTravailClientSecret }
if (-not $PSBoundParameters.ContainsKey("FranceTravailScope")) {
    $scopeDefault = [string](Get-ConfigPathValue -Object $JobCrawlerSourcesConfig -Path "credentials.france_travail.scope.default" -DefaultValue "api_offresdemploiv2 o2dsoffre")
    $FranceTravailScope = Get-JobCrawlerCredentialValue -SourcesConfig $JobCrawlerSourcesConfig -SourceKey "france_travail" -CredentialKey "scope" -FallbackValue $scopeDefault
}
if (-not $PSBoundParameters.ContainsKey("AdzunaAppId")) { $AdzunaAppId = Get-JobCrawlerCredentialValue -SourcesConfig $JobCrawlerSourcesConfig -SourceKey "adzuna" -CredentialKey "app_id" -FallbackValue $AdzunaAppId }
if (-not $PSBoundParameters.ContainsKey("AdzunaAppKey")) { $AdzunaAppKey = Get-JobCrawlerCredentialValue -SourcesConfig $JobCrawlerSourcesConfig -SourceKey "adzuna" -CredentialKey "app_key" -FallbackValue $AdzunaAppKey }

function ConvertTo-JobCrawlerSourceKeySet {
    param([AllowNull()][string[]]$SourceKeys)

    $set = @{}
    foreach ($sourceKeyValue in @($SourceKeys)) {
        foreach ($sourceKeyPart in @(([string]$sourceKeyValue) -split ",")) {
            $sourceKey = ConvertTo-JobCrawlerProfileId $sourceKeyPart
            if (-not [string]::IsNullOrWhiteSpace($sourceKey)) {
                $set[$sourceKey] = $true
            }
        }
    }

    return $set
}

$SourceDefinitions = @(Get-JobCrawlerSourceDefinitions -SourcesConfig $JobCrawlerSourcesConfig)
$sourceContract = Assert-JobCrawlerSourceContract -SourceDefinitions $SourceDefinitions -LoadedFiles $LoadedSourceAdapters
$EnabledSourceKeys = ConvertTo-JobCrawlerSourceKeySet -SourceKeys $EnableSource
$SkippedSourceKeys = ConvertTo-JobCrawlerSourceKeySet -SourceKeys $SkipSource
$KnownSourceKeys = @{}
foreach ($source in $SourceDefinitions) {
    $KnownSourceKeys[(ConvertTo-JobCrawlerProfileId ([string]$source.Key))] = $true
}
$requestedSourceKeys = @((@($EnabledSourceKeys.Keys) + @($SkippedSourceKeys.Keys)) | Select-Object -Unique)
foreach ($sourceKey in $requestedSourceKeys) {
    if (-not $KnownSourceKeys.ContainsKey($sourceKey)) {
        Write-Warning ("Unknown source key '{0}'. Configure it in config\sources.json before using -EnableSource/-SkipSource." -f $sourceKey)
    }
}

$SourceEnabled = @{}
foreach ($source in $SourceDefinitions) {
    $normalizedSourceKey = ConvertTo-JobCrawlerProfileId ([string]$source.Key)
    $enabled = [bool]$source.EnabledByDefault
    if ($EnabledSourceKeys.ContainsKey($normalizedSourceKey)) {
        $enabled = $true
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$source.EnableSwitch) -and $PSBoundParameters.ContainsKey([string]$source.EnableSwitch)) {
        $enabled = $true
    }
    if ($SkipWttj -and ([string]$source.Key -in @("wttj_public", "welcome_kit"))) {
        $enabled = $false
    }
    if ($SkippedSourceKeys.ContainsKey($normalizedSourceKey)) {
        $enabled = $false
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$source.SkipSwitch)) {
        $skipVariable = Get-Variable -Name ([string]$source.SkipSwitch) -ErrorAction SilentlyContinue
        if ($null -ne $skipVariable -and [bool]$skipVariable.Value) {
            $enabled = $false
        }
    }

    $SourceEnabled[[string]$source.Key] = $enabled
}

$welcomeKitSourceEnabled = [bool]$SourceEnabled["welcome_kit"]
if (-not $welcomeKitSourceEnabled) { $WelcomeKitApiKey = "" }

$modeConfig = Get-ConfigPathValue -Object $JobCrawlerConfig.CrawlModes -Path ("modes.{0}" -f $CrawlMode) -DefaultValue $null
if ($null -eq $modeConfig) {
    throw "Unknown crawl mode '$CrawlMode'. Configure it in config\crawl_modes.json."
}
$crawlModeParameterMap = @{
    MaxLinkedInSearchPages   = "max_linkedin_search_pages"
    MaxLinkedInDetails       = "max_linkedin_details"
    MaxFranceTravailPages    = "max_france_travail_pages"
    MaxAdzunaPages           = "max_adzuna_pages"
    MaxApecPages             = "max_apec_pages"
    MaxHelloWorkPages        = "max_hellowork_pages"
    MaxHelloWorkCardsPerQuery = "max_hellowork_cards_per_query"
    MaxHelloWorkDetails      = "max_hellowork_details"
    MaxWelcomeKitPages       = "max_welcome_kit_pages"
    MaxWttjCandidatePages    = "max_wttj_candidate_pages"
}
foreach ($parameterName in $crawlModeParameterMap.Keys) {
    if (-not $PSBoundParameters.ContainsKey($parameterName)) {
        $configuredValue = Get-ConfigProperty -Object $modeConfig -Name $crawlModeParameterMap[$parameterName] -DefaultValue $null
        if ($null -ne $configuredValue) {
            Set-Variable -Name $parameterName -Value ([int]$configuredValue) -Scope Local
        }
    }
}

$BrowserUserAgent = [string](Get-ConfigPathValue -Object $JobCrawlerRuntimeConfig -Path "http.user_agent" -DefaultValue "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")
$Cutoff = [DateTimeOffset]::Now.AddDays(-[Math]::Abs($DaysBack))
$CutoffDate = $Cutoff.ToString("yyyy-MM-dd")
$RunDate = Get-Date -Format "yyyy-MM-dd"
$RunStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$DefaultTrackerPath = Resolve-JobCrawlerPath -BasePath $ProjectRoot -Path ([string](Get-ConfigPathValue -Object $JobCrawlerRuntimeConfig -Path "defaults.tracker_path" -DefaultValue "output\jobs_tracker.xlsx"))
$CacheDirectory = Resolve-JobCrawlerPath -BasePath $ProjectRoot -Path ([string](Get-ConfigPathValue -Object $JobCrawlerRuntimeConfig -Path "defaults.cache_directory" -DefaultValue "output\cache"))
$RunHistoryPath = Resolve-JobCrawlerPath -BasePath $ProjectRoot -Path ([string](Get-ConfigPathValue -Object $JobCrawlerRuntimeConfig -Path "output.run_history_path" -DefaultValue "output\run_history.jsonl"))
$RunHistoryMaxEntries = [int](Get-ConfigPathValue -Object $JobCrawlerRuntimeConfig -Path "output.run_history_max_entries" -DefaultValue 250)

if ([string]::IsNullOrWhiteSpace($TrackerPath)) {
    $TrackerPath = $DefaultTrackerPath
}
if ([IO.Path]::GetExtension($TrackerPath).ToLowerInvariant() -ne ".xlsx") {
    throw "This crawler uses only the XLSX tracker file. Use output\jobs_tracker.xlsx for -TrackerPath."
}

$SeenResultKeys = @{}
$LinkedInDelayMilliseconds = [int](Get-ConfigPathValue -Object $JobCrawlerRuntimeConfig -Path "delays_ms.linkedin_detail" -DefaultValue 1200)
$AdzunaDelayMilliseconds = [int](Get-ConfigPathValue -Object $JobCrawlerRuntimeConfig -Path "delays_ms.adzuna_search" -DefaultValue 2500)
$ApecDelayMilliseconds = [int](Get-ConfigPathValue -Object $JobCrawlerRuntimeConfig -Path "delays_ms.apec_search" -DefaultValue 300)
$HelloWorkSearchDelayMilliseconds = [int](Get-ConfigPathValue -Object $JobCrawlerRuntimeConfig -Path "delays_ms.hellowork_search" -DefaultValue 350)
$HelloWorkDetailDelayMilliseconds = [int](Get-ConfigPathValue -Object $JobCrawlerRuntimeConfig -Path "delays_ms.hellowork_detail" -DefaultValue 450)
$MinimumMatchScore = [int](Get-ConfigPathValue -Object $JobCrawlerMatchingRules -Path "thresholds.minimum_match_score" -DefaultValue 35)
$JobCrawlerPreferences = $null
$FeedbackLearningProfile = $null
$SourceRunStats = New-Object System.Collections.Generic.List[object]

$WttjUrlCandidatePattern = [string](Get-ConfigPathValue -Object $JobCrawlerSourcesConfig -Path "patterns.wttj_url_candidate" -DefaultValue "")
$LinkedInQueries = @(Get-JobCrawlerSourceQueryList -SourcesConfig $JobCrawlerSourcesConfig -SourceKey "linkedin" -FallbackKeys @("api"))
$HelloWorkQueries = @(Get-JobCrawlerSourceQueryList -SourcesConfig $JobCrawlerSourcesConfig -SourceKey "hellowork" -FallbackKeys @("api"))
$ApecQueries = @(Get-JobCrawlerSourceQueryList -SourcesConfig $JobCrawlerSourcesConfig -SourceKey "apec" -FallbackKeys @("api"))
$FranceTravailQueries = @(Get-JobCrawlerSourceQueryList -SourcesConfig $JobCrawlerSourcesConfig -SourceKey "france_travail" -FallbackKeys @("api"))
$AdzunaQueries = @(Get-JobCrawlerSourceQueryList -SourcesConfig $JobCrawlerSourcesConfig -SourceKey "adzuna" -FallbackKeys @("api"))
$ApiSearchQueries = @(Get-ConfigStringArray (Get-ConfigPathValue -Object $JobCrawlerSourcesConfig -Path "queries.api" -DefaultValue @()))







































































$MasterColumns = Get-JobTrackerMasterColumns
$ColumnLabels = Get-JobTrackerColumnLabels


















































































$JobCrawlerPreferences = Get-JobCrawlerPreferences

$JobCrawlerContext = New-JobCrawlerContext -ProjectRoot $ProjectRoot -ConfigDirectory $configPath -Config $JobCrawlerConfig -Runtime @{
    BrowserUserAgent = $BrowserUserAgent
    Cutoff = $Cutoff
    CutoffDate = $CutoffDate
    RunDate = $RunDate
    RunStamp = $RunStamp
    TrackerPath = $TrackerPath
    DefaultTrackerPath = $DefaultTrackerPath
    CacheDirectory = $CacheDirectory
    CacheTtlHours = $CacheTtlHours
    DisableCache = [bool]$DisableCache
    CrawlMode = $CrawlMode
    Location = $Location
    MinimumMatchScore = $MinimumMatchScore
    SourceDefinitions = $SourceDefinitions
    SourceEnabled = $SourceEnabled
    LoadedSourceAdapters = $LoadedSourceAdapters
}
Set-JobCrawlerScriptContext -Context $JobCrawlerContext | Out-Null

if ($SelfTest) {
    Invoke-ScoringSelfTest
    return
}

Set-RunWindowTitle "Custom Job Tracker - Starting"
Write-RunStatus ("Starting crawl for jobs published since {0}." -f $CutoffDate)
Write-RunStatus ("Profile: {0} ({1})" -f $JobCrawlerConfig.Profile.Label, $JobCrawlerConfig.Profile.Id)
Write-RunStatus ("Tracker file: {0}" -f $TrackerPath)
Write-RunStatus ("Loaded {0} source adapter file(s)." -f @($sourceContract.LoadedFiles).Count)

$cachePruneResult = [PSCustomObject]@{ RemovedFiles = 0; RemovedBytes = 0; RemainingBytes = 0 }
if (-not $DisableCache) {
    $cachePruneEnabled = ConvertTo-ConfigBoolean -Value (Get-ConfigPathValue -Object $JobCrawlerRuntimeConfig -Path "defaults.cache_prune_enabled" -DefaultValue $true) -DefaultValue $true
    $cachePruneResult = Invoke-JobCrawlerCachePrune -Path $CacheDirectory -Enabled:$cachePruneEnabled
    if ([int]$cachePruneResult.RemovedFiles -gt 0) {
        Write-RunStatus ("Cache pruning removed {0} file(s), {1:N1} MB; remaining cache {2:N1} MB." -f $cachePruneResult.RemovedFiles, ([double]$cachePruneResult.RemovedBytes / 1MB), ([double]$cachePruneResult.RemainingBytes / 1MB))
    }
}

$existingTrackerRows = @(Import-TrackerRows -Path $TrackerPath)
Write-RunStatus ("Loaded {0} existing tracker row(s)." -f $existingTrackerRows.Count)
$script:FeedbackLearningProfile = New-FeedbackLearningProfile -Rows $existingTrackerRows
Write-RunStatus ("Feedback profile: {0} positive row(s), {1} ignored row(s)." -f $script:FeedbackLearningProfile.PositiveRows, $script:FeedbackLearningProfile.IgnoredRows)

$allResults = New-Object System.Collections.Generic.List[object]
$sourceResultCounts = @{}
foreach ($source in $SourceDefinitions) {
    $sourceKey = [string]$source.Key
    if (-not [bool]$SourceEnabled[$sourceKey]) {
        continue
    }

    if ($sourceKey -eq "welcome_kit" -and [string]::IsNullOrWhiteSpace($WelcomeKitApiKey)) {
        Write-RunStatus "WelcomeKit source enabled but WK_API_KEY is not set; WTTJ public fallback remains available."
        continue
    }

    if ($sourceKey -eq "wttj_public" -and $welcomeKitSourceEnabled -and -not [string]::IsNullOrWhiteSpace($WelcomeKitApiKey)) {
        Write-RunStatus "Skipping WTTJ public fallback because WelcomeKit API is enabled and configured."
        continue
    }

    $sourceResultCounts[$sourceKey] = Invoke-ConfiguredCrawlerSource -SourceDefinition $source -Target $allResults
}

$sortedCrawlResults = $allResults |
    Sort-Object -Property @{ Expression = "match_score"; Descending = $true }, @{ Expression = "published_date"; Descending = $true }, platform, job_title
$postEnrichmentGate = Invoke-JobPipelineEligibilityGate -Rows @($sortedCrawlResults) -Stage "post_enrichment"
$filteredCrawlResults = @($postEnrichmentGate.KeptRows)
$excludedOldCount = Get-JobPipelineReasonCount -GateResult $postEnrichmentGate -Reason @("outside_published_window", "missing_published_date")
$excludedContractCount = Get-JobPipelineReasonCount -GateResult $postEnrichmentGate -Reason @("excluded_contract")
$excludedInvalidLocationCount = Get-JobPipelineReasonCount -GateResult $postEnrichmentGate -Reason @("invalid_location")

if ($excludedOldCount -gt 0) {
    Write-RunStatus ("Excluded {0} job(s) outside the published-date window after detail enrichment." -f $excludedOldCount)
}

if ($excludedContractCount -gt 0) {
    Write-RunStatus ("Excluded {0} CDD/apprenticeship/internship/freelance job(s) from this crawl." -f $excludedContractCount)
}

if ($excludedInvalidLocationCount -gt 0) {
    Write-RunStatus ("Excluded {0} job(s) with invalid or unsupported location after detail enrichment." -f $excludedInvalidLocationCount)
}

$diagnosticPath = ""
if ($DiagnosticMode) {
    $diagnosticDirectory = Resolve-JobCrawlerPath -BasePath $ProjectRoot -Path ([string](Get-ConfigPathValue -Object $JobCrawlerRuntimeConfig -Path "output.diagnostics_directory" -DefaultValue "output\diagnostics"))
    if (-not (Test-Path -LiteralPath $diagnosticDirectory)) {
        New-Item -ItemType Directory -Force -Path $diagnosticDirectory | Out-Null
    }
    $diagnosticPath = Join-Path $diagnosticDirectory ("crawl_diagnostics_{0}.csv" -f $RunStamp)
    $diagnosticRows = @($postEnrichmentGate.Decisions | ForEach-Object {
        $row = $_.Row
        $contractType = Get-RowValue -Row $row -Name "contract_type"
        $diagnosticStatus = $(if ($_.IsEligible) { "kept_for_merge" } else { $_.Reason })
        [PSCustomObject]@{
            diagnostic_status = $diagnosticStatus
            platform          = Get-RowValue -Row $row -Name "platform"
            match_level       = Get-RowValue -Row $row -Name "match_level"
            match_score       = Get-RowValue -Row $row -Name "match_score"
            job_title         = Get-RowValue -Row $row -Name "job_title"
            company_name      = Get-RowValue -Row $row -Name "company_name"
            location          = Get-RowValue -Row $row -Name "location"
            contract_type     = $contractType
            published_date    = Get-RowValue -Row $row -Name "published_date"
            matched_keywords  = Get-RowValue -Row $row -Name "matched_keywords"
            job_url           = Get-RowValue -Row $row -Name "job_url_raw"
        }
    })
    $diagnosticRows | Export-Csv -LiteralPath $diagnosticPath -NoTypeInformation -Encoding UTF8
    Write-RunStatus ("Diagnostics written: {0}" -f $diagnosticPath)
}

$feedbackAdjustedResults = @(Apply-FeedbackScoring -Rows $filteredCrawlResults -ExistingRows $existingTrackerRows)
$mergeResult = Merge-JobsWithTracker -CurrentRows $feedbackAdjustedResults -ExistingRows $existingTrackerRows -Path $TrackerPath -SkipBackup:$DryRun
$finalResults = @($mergeResult.TrackerRows)
Test-JobPipelineInvariants -Rows $finalResults -Stage "pre_export" -ThrowOnIssue | Out-Null

$crawlCaps = @(
    "LinkedIn pages $MaxLinkedInSearchPages"
    "LinkedIn details $MaxLinkedInDetails"
    "France Travail pages $MaxFranceTravailPages"
    "Adzuna pages $MaxAdzunaPages"
    "APEC pages $MaxApecPages"
    "HelloWork pages $MaxHelloWorkPages"
    "HelloWork cards/query $MaxHelloWorkCardsPerQuery"
    "HelloWork details $MaxHelloWorkDetails"
    "WelcomeKit pages $MaxWelcomeKitPages"
    "WTTJ candidates $MaxWttjCandidatePages"
) -join " | "

$crawlSummary = @{
    Profile = ("{0} ({1})" -f $JobCrawlerConfig.Profile.Label, $JobCrawlerConfig.Profile.Id)
    TotalMatched = @($sortedCrawlResults).Count
    ExcludedContractCount = $excludedContractCount
    ExcludedOldCount = $excludedOldCount
    ExcludedInvalidLocationCount = $excludedInvalidLocationCount
    PipelineExclusionSummary = Format-JobPipelineReasonSummary -GateResult $postEnrichmentGate
    CurrentCount = @($mergeResult.CurrentRows).Count
    TrackerCount = @($finalResults).Count
    DuplicateCount = $mergeResult.DuplicateCount
    RemovedCount = $mergeResult.RemovedCount
    PreservedAppliedCount = $mergeResult.PreservedAppliedCount
    SourceDiagnostics = Get-SourceStatsSummaryText
    BackupPath = $mergeResult.BackupPath
    CrawlCaps = $crawlCaps
    DryRun = $(if ($DryRun) { "yes" } else { "no" })
    DiagnosticMode = $(if ($DiagnosticMode) { "yes" } else { "no" })
    DiagnosticPath = $diagnosticPath
    CachePrunedFiles = $cachePruneResult.RemovedFiles
    CachePrunedMB = [Math]::Round(([double]$cachePruneResult.RemovedBytes / 1MB), 1)
    CacheRemainingMB = [Math]::Round(([double]$cachePruneResult.RemainingBytes / 1MB), 1)
    RunHistoryPath = $RunHistoryPath
}

if (-not $DryRun) {
    Export-TrackerWorkbook -Rows $finalResults -Path $TrackerPath -Summary $crawlSummary
}

Write-RunHistoryEntry -Summary $crawlSummary -Path $RunHistoryPath -MaxEntries $RunHistoryMaxEntries

Set-RunWindowTitle "Custom Job Tracker - Finished"
Write-Host ""
if ($DryRun) {
    Write-RunStatus ("Dry run complete. Workbook was not written. Simulated tracker rows: {0}; current jobs: {1}; preserved application rows: {2}; removed by retention: {3}." -f @($finalResults).Count, @($mergeResult.CurrentRows).Count, $mergeResult.PreservedAppliedCount, $mergeResult.RemovedCount)
}
else {
    Write-RunStatus ("Wrote tracker with {0} row(s): {1} current job(s), {2} preserved application row(s), {3} removed by retention." -f @($finalResults).Count, @($mergeResult.CurrentRows).Count, $mergeResult.PreservedAppliedCount, $mergeResult.RemovedCount)
}
Write-RunStatus ("Crawl mode: {0}. Source diagnostics: {1}" -f $CrawlMode, (Get-SourceStatsSummaryText))
if (-not [string]::IsNullOrWhiteSpace($mergeResult.BackupPath)) {
    Write-RunStatus ("Backup: {0}" -f $mergeResult.BackupPath)
}
Write-RunStatus ("Tracker: {0}" -f (Resolve-Path $TrackerPath).Path)

if (@($finalResults).Count -gt 0) {
    $finalResults | Format-Table -AutoSize
}
