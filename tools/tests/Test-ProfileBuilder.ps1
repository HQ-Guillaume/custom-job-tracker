[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
. (Join-Path $projectRoot "app\core\JobTracker.Common.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Config.ps1")

function Assert-ProfileBuilder {
    param(
        [bool]$Condition,
        [string]$Message
    )

    if (-not $Condition) {
        throw "Profile builder test failed: $Message"
    }
}

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ("custom-job-tracker-profile-{0}" -f ([Guid]::NewGuid().ToString("N")))
$tempConfig = Join-Path $tempRoot "config"

New-Item -ItemType Directory -Path $tempRoot | Out-Null
try {
    Copy-Item -LiteralPath (Join-Path $projectRoot "config") -Destination $tempConfig -Recurse -Force
    Remove-Item -LiteralPath (Join-Path $tempConfig "local") -Recurse -Force -ErrorAction SilentlyContinue
    Get-ChildItem -LiteralPath $tempConfig -Filter "local*.json" -File -ErrorAction SilentlyContinue |
        Remove-Item -Force
    Get-ChildItem -LiteralPath (Join-Path $tempConfig "profiles") -Filter "*.json" -File -ErrorAction SilentlyContinue |
        Remove-Item -Force

    $publicConfig = Get-JobCrawlerConfig -ConfigDirectory $tempConfig
    $publicValidation = Test-JobCrawlerConfig -Config $publicConfig
    Assert-ProfileBuilder -Condition $publicValidation.IsValid -Message "Expected clean public config without profiles to remain structurally valid."
    Assert-ProfileBuilder -Condition (-not (Test-JobCrawlerProfileConfigured -Config $publicConfig)) -Message "Expected clean public config to require profile creation before crawling."

    $compactProfile = New-JobCrawlerProfileFromBuilder `
        -Label "CRM Analyst" `
        -Description "CRM and lifecycle analytics roles." `
        -TargetTitles @("CRM analyst", "Lifecycle analyst") `
        -ImportantSkills @("Braze", "HubSpot", "SQL") `
        -ExclusionKeywords @("data engineer", "backend") `
        -SearchQueries @("crm analyst braze", "lifecycle analyst hubspot") `
        -TargetLocations @("France", "Paris") `
        -ExcludedLocations @("London", "New York") `
        -ExcludedContracts @("CDD", "Internship", "Freelance") `
        -EmployerPreference "neutral" `
        -Compact

    Assert-ProfileBuilder -Condition ($compactProfile["id"] -eq "crm_analyst") -Message "Expected stable generated profile id."
    Assert-ProfileBuilder -Condition ($null -ne $compactProfile["profile_builder"]) -Message "Expected compact profile builder payload."
    Assert-ProfileBuilder -Condition (-not $compactProfile.Contains("matching_rules")) -Message "Compact profiles should not store generated matching rules."
    Assert-ProfileBuilder -Condition (-not $compactProfile.Contains("preferences")) -Message "Compact profiles should not store generated preferences."

    $savedPath = Save-JobCrawlerLocalProfile -ConfigDirectory $tempConfig -Profile $compactProfile
    Assert-ProfileBuilder -Condition (Test-Path -LiteralPath $savedPath) -Message "Expected local profile file to be saved."
    [void](Set-JobCrawlerDefaultProfile -ConfigDirectory $tempConfig -ProfileId "crm_analyst")

    $loadedConfig = Get-JobCrawlerConfig -ConfigDirectory $tempConfig
    Assert-ProfileBuilder -Condition ($loadedConfig.Profile.Id -eq "crm_analyst") -Message "Expected local default profile to load."
    $loadedTrackerPath = Get-JobCrawlerTrackerPath -ProjectRoot $tempRoot -Config $loadedConfig
    Assert-ProfileBuilder -Condition ($loadedTrackerPath -like "*output*profiles*crm_analyst*jobs_tracker.xlsx") -Message "Expected default tracker path to be profile-specific."
    Assert-ProfileBuilder -Condition (@(Get-ConfigStringArray (Get-ConfigPathValue -Object $loadedConfig.Sources -Path "queries.linkedin" -DefaultValue @())).Count -gt 0) -Message "Expected compact profile to expand LinkedIn queries."
    Assert-ProfileBuilder -Condition (-not [string]::IsNullOrWhiteSpace([string](Get-ConfigPathValue -Object $loadedConfig.MatchingRules -Path "contexts.profile_skill_context" -DefaultValue ""))) -Message "Expected compact profile to expand generic profile skill context."
    Assert-ProfileBuilder -Condition (@(Get-ConfigPathValue -Object $loadedConfig.MatchingRules -Path "positive_signals" -DefaultValue @()).Count -gt 0) -Message "Expected compact profile to expand positive signals."
    Assert-ProfileBuilder -Condition (@(Get-ConfigPathValue -Object $loadedConfig.Preferences -Path "target_location_patterns" -DefaultValue @()).Count -gt 0) -Message "Expected compact profile to expand location preferences."

    $singleQueryProfile = New-JobCrawlerProfileFromBuilder `
        -Label "One Query Analyst" `
        -TargetTitles @("One Query Analyst") `
        -SearchQueries @("one query analyst") `
        -TargetLocations @("France") `
        -Compact
    $singleQuerySavedQueries = @(Get-ConfigStringArray (Get-ConfigProperty -Object $singleQueryProfile["profile_builder"] -Name "search_queries" -DefaultValue @()))
    Assert-ProfileBuilder -Condition ($singleQuerySavedQueries.Count -gt 1) -Message "Expected a one-query profile to be expanded with safer query suggestions."
    $singleQueryPath = Save-JobCrawlerLocalProfile -ConfigDirectory $tempConfig -Profile $singleQueryProfile
    Assert-ProfileBuilder -Condition (Test-Path -LiteralPath $singleQueryPath) -Message "Expected one-query local profile file to be saved."
    $singleQueryConfig = Get-JobCrawlerConfig -ConfigDirectory $tempConfig -ProfileId "one_query_analyst"
    $singleQueryValidation = Test-JobCrawlerConfig -Config $singleQueryConfig
    Assert-ProfileBuilder -Condition $singleQueryValidation.IsValid -Message "Expected one-query profile validation not to fail on scalar Count behavior."
    Assert-ProfileBuilder -Condition (@(Get-ConfigStringArray (Get-ConfigPathValue -Object $singleQueryConfig.Sources -Path "queries.linkedin" -DefaultValue @())).Count -gt 1) -Message "Expected expanded LinkedIn queries for one-query profiles."

    $weakQuality = Get-JobCrawlerProfileQuality `
        -Label "One Query Analyst" `
        -TargetTitles @("One Query Analyst") `
        -SearchQueries @("one query analyst") `
        -TargetLocations @("France")
    Assert-ProfileBuilder -Condition ($weakQuality.Score -lt 70) -Message "Expected sparse profiles to receive a low quality score."
    Assert-ProfileBuilder -Condition (@($weakQuality.QuerySuggestions).Count -gt 1) -Message "Expected sparse profile quality result to include query suggestions."

    $strongQuality = Get-JobCrawlerProfileQuality `
        -Label "CRM Analyst" `
        -TargetTitles @("CRM analyst", "Lifecycle analyst", "Marketing automation analyst") `
        -SearchQueries @("crm analyst", "lifecycle analytics", "crm analyst braze", "marketing automation hubspot", "customer analytics") `
        -ImportantSkills @("Braze", "HubSpot", "SQL", "dashboarding", "segmentation", "campaign analytics") `
        -ExclusionKeywords @("backend", "data engineer", "internship") `
        -TargetLocations @("France") `
        -ExcludedContracts @("CDD", "Internship", "Freelance")
    Assert-ProfileBuilder -Condition ($strongQuality.Score -ge 70) -Message "Expected richer profiles to receive a good quality score."

    $digitalSuggestions = @(Get-JobCrawlerSearchQuerySuggestions -Label "Digital analyst" -TargetTitles @("digital analyst") -ImportantSkills @())
    $mergedDigitalQueries = @(Merge-JobCrawlerProfileLineArrays -Primary "digital analyst" -Secondary $digitalSuggestions -MaxItems 24)
    Assert-ProfileBuilder -Condition ($mergedDigitalQueries[0] -eq "digital analyst") -Message "Expected query merge to keep the existing scalar query as its own line."
    Assert-ProfileBuilder -Condition ($mergedDigitalQueries -contains "digital analytics") -Message "Expected query merge to add suggested role variants."
    Assert-ProfileBuilder -Condition (-not ($mergedDigitalQueries[0] -like "digital analystdigital*")) -Message "Expected query merge not to concatenate scalar and array values."

    $removedProfilePath = Remove-JobCrawlerLocalProfile -ConfigDirectory $tempConfig -ProfileId "one_query_analyst"
    Assert-ProfileBuilder -Condition (-not (Test-Path -LiteralPath $removedProfilePath)) -Message "Expected local profile deletion to remove the profile file."
    [void](Set-JobCrawlerDefaultProfile -ConfigDirectory $tempConfig -ProfileId "crm_analyst")
    $removedDefaultPath = Remove-JobCrawlerLocalProfile -ConfigDirectory $tempConfig -ProfileId "crm_analyst"
    Assert-ProfileBuilder -Condition (-not (Test-Path -LiteralPath $removedDefaultPath)) -Message "Expected deleting the default profile to remove the profile file."
    $afterDeleteConfig = Get-JobCrawlerConfig -ConfigDirectory $tempConfig
    Assert-ProfileBuilder -Condition (-not (Test-JobCrawlerProfileConfigured -Config $afterDeleteConfig)) -Message "Expected deleting the last local profile to leave the project unconfigured instead of pointing at a stale default."
}
finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "Profile builder tests passed."
