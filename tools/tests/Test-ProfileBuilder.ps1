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
    Assert-ProfileBuilder -Condition (@(Get-ConfigStringArray (Get-ConfigPathValue -Object $loadedConfig.Sources -Path "queries.linkedin" -DefaultValue @())).Count -gt 0) -Message "Expected compact profile to expand LinkedIn queries."
    Assert-ProfileBuilder -Condition (-not [string]::IsNullOrWhiteSpace([string](Get-ConfigPathValue -Object $loadedConfig.MatchingRules -Path "contexts.profile_skill_context" -DefaultValue ""))) -Message "Expected compact profile to expand generic profile skill context."
    Assert-ProfileBuilder -Condition (@(Get-ConfigPathValue -Object $loadedConfig.MatchingRules -Path "positive_signals" -DefaultValue @()).Count -gt 0) -Message "Expected compact profile to expand positive signals."
    Assert-ProfileBuilder -Condition (@(Get-ConfigPathValue -Object $loadedConfig.Preferences -Path "target_location_patterns" -DefaultValue @()).Count -gt 0) -Message "Expected compact profile to expand location preferences."
}
finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "Profile builder tests passed."
