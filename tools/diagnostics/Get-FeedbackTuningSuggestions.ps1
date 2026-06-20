[CmdletBinding()]
param(
    [string]$TrackerPath = "",
    [string]$ConfigDirectory = "config",
    [string]$Profile = "",
    [string]$OutputPath = ""
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$coreRoot = Join-Path $projectRoot "app\core"

. (Join-Path $coreRoot "JobTracker.Common.ps1")
. (Join-Path $coreRoot "JobTracker.Config.ps1")
. (Join-Path $coreRoot "JobTracker.Runtime.ps1")
. (Join-Path $coreRoot "JobTracker.Scoring.ps1")
. (Join-Path $coreRoot "JobTracker.Deduplication.ps1")
. (Join-Path $coreRoot "JobTracker.Excel.ps1")

$configPath = Resolve-JobCrawlerPath -BasePath $projectRoot -Path $ConfigDirectory
$script:JobCrawlerConfig = Get-JobCrawlerConfig -ConfigDirectory $configPath -ProfileId $Profile
$script:JobCrawlerRuntimeConfig = $script:JobCrawlerConfig.Runtime
$script:JobCrawlerSourcesConfig = $script:JobCrawlerConfig.Sources
$script:JobCrawlerMatchingRules = $script:JobCrawlerConfig.MatchingRules
$script:JobCrawlerWorkbookConfig = $script:JobCrawlerConfig.Workbook
$script:JobCrawlerPreferences = Get-JobCrawlerPreferences
$MasterColumns = Get-JobTrackerMasterColumns
$ColumnLabels = Get-JobTrackerColumnLabels

if ([string]::IsNullOrWhiteSpace($TrackerPath)) {
    $TrackerPath = Get-JobCrawlerTrackerPath -ProjectRoot $projectRoot -Config $script:JobCrawlerConfig
}
if (-not (Test-Path -LiteralPath $TrackerPath)) {
    throw "Tracker workbook not found: $TrackerPath"
}

$rows = @(Import-TrackerRows -Path $TrackerPath)
$learning = New-FeedbackLearningProfile -Rows $rows
$suggestions = New-Object System.Collections.Generic.List[object]

function Add-Suggestion {
    param(
        [string]$Area,
        [string]$Suggestion,
        [int]$EvidenceCount,
        [string]$Action
    )

    $suggestions.Add([PSCustomObject]@{
        Area = $Area
        Suggestion = $Suggestion
        EvidenceCount = $EvidenceCount
        RecommendedAction = $Action
    }) | Out-Null
}

foreach ($entry in @($learning.PositiveSignalCounts.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 8)) {
    if ([int]$entry.Value -ge 2) {
        Add-Suggestion -Area "Positive signals" -Suggestion ("Keep or strengthen signal '{0}'" -f $entry.Key) -EvidenceCount ([int]$entry.Value) -Action "Review profile important_skills/search_queries before changing weights."
    }
}

foreach ($entry in @($learning.IgnoreReasonCounts.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 10)) {
    if ([int]$entry.Value -ge 2) {
        Add-Suggestion -Area "Negative signals" -Suggestion ("Ignored reason appears often: {0}" -f $entry.Key) -EvidenceCount ([int]$entry.Value) -Action "Consider adding or strengthening a negative rule only if the ignored rows are consistently poor matches."
    }
}

$ignoredWithoutReason = @($rows | Where-Object {
    ((Get-RowValue -Row $_ -Name "status").ToLowerInvariant() -eq "ignored") -and [string]::IsNullOrWhiteSpace((Get-IgnoreReasonFromNotes (Get-RowValue -Row $_ -Name "notes")))
}).Count
if ($ignoredWithoutReason -gt 0) {
    Add-Suggestion -Area "Feedback quality" -Suggestion "Some ignored rows have no structured ignore_reason." -EvidenceCount $ignoredWithoutReason -Action "Fill Apply notes with ignore_reason=... so future scoring learns safely."
}

$agencyIgnored = [int](Get-ConfigProperty -Object $learning.IgnoreReasonCounts -Name "agency_consulting_esn" -DefaultValue 0)
if ($agencyIgnored -ge 2) {
    Add-Suggestion -Area "Employer preference" -Suggestion "Agency/consulting/ESN is a frequent ignore reason." -EvidenceCount $agencyIgnored -Action "Keep employer preference as a soft penalty, not a hard filter, to avoid missing strong matches."
}

$resultRows = @($suggestions.ToArray())
if ($resultRows.Count -eq 0) {
    $resultRows = @([PSCustomObject]@{
        Area = "Feedback"
        Suggestion = "Not enough structured feedback yet."
        EvidenceCount = 0
        RecommendedAction = "Keep using status and ignore_reason notes; rerun this report after more decisions."
    })
}

if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $fullOutputPath = Resolve-JobCrawlerPath -BasePath $projectRoot -Path $OutputPath
    $directory = Split-Path -Parent $fullOutputPath
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }
    $resultRows | Export-Csv -LiteralPath $fullOutputPath -NoTypeInformation -Encoding UTF8
    Write-Host ("Wrote feedback tuning suggestions: {0}" -f $fullOutputPath)
}
else {
    $resultRows | Format-Table -AutoSize
}
