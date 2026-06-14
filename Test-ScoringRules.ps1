[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "JobTracker.Common.ps1")

function Assert-Rule {
    param(
        [bool]$Condition,
        [string]$Message
    )

    if (-not $Condition) {
        throw "Scoring rule test failed: $Message"
    }
}

Assert-Rule -Condition ((Get-IgnoreReasonFromNotes "ignore_reason=too_data_engineering; detail=dbt") -eq "too_data_engineering") -Message "Structured data-engineering ignore reason was not parsed."
Assert-Rule -Condition ((Get-IgnoreReasonFromNotes "mostly SEO SEA acquisition") -eq "too_seo_sea_marketing") -Message "Free-text SEO/SEA ignore reason was not inferred."
Assert-Rule -Condition (Test-Path -LiteralPath (Join-Path $PSScriptRoot "config\preferences.json")) -Message "Preference config file is missing."

& (Join-Path $PSScriptRoot "Find-AnalyticsJobs.ps1") -SelfTest

Write-Host "Scoring rule tests passed."
