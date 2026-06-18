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

function Assert-OpenXmlWorkbook {
    param(
        [bool]$Condition,
        [string]$Message
    )

    if (-not $Condition) {
        throw "OpenXML workbook test failed: $Message"
    }
}

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
$script:LinkedInQueries = @("web analyst")
$script:HelloWorkQueries = @("digital analyst")
$script:ApecQueries = @("tracking")
$script:FranceTravailQueries = @("ga4")
$script:AdzunaQueries = @("gtm")
$script:ApiSearchQueries = @("analytics")

$today = Get-Date -Format "yyyy-MM-dd"
$row = New-OrderedJobRecord @{
    review_priority       = "New High"
    status                = "interesting"
    job_title             = "Web Analyst"
    company_name          = "Radio France"
    employer_type         = "annonceur"
    location              = "Paris"
    contract_type         = "CDI"
    platform              = "LinkedIn"
    source_count          = "1"
    published_date        = $today
    days_since_published  = "0"
    job_url               = "Open"
    job_url_raw           = "https://example.com/job"
    applied_date          = ""
    notes                 = "Looks relevant"
    match_level           = "High"
    match_score           = "92"
    matched_keywords      = "Role: web analyst | Tool: GA4"
    role_score            = "80"
    employer_fit          = "12"
    location_fit          = "10"
    seniority_fit         = "0"
    contract_fit          = "5"
    fit_notes             = "role score 80"
    seen_in_current_crawl = "yes"
    first_seen_date       = $today
    last_seen_date        = $today
    is_new                = "yes"
    duplicate_reason      = ""
    feedback_adjustment   = ""
    job_id                = "test-job"
    alternate_urls        = ""
    seen_before           = "no"
    days_since_first_seen = "0"
    days_since_last_seen  = "0"
}

$tempPath = Join-Path ([IO.Path]::GetTempPath()) ("custom-job-tracker-openxml-test-{0}.xlsx" -f ([Guid]::NewGuid().ToString("N")))
try {
    Export-TrackerWorkbook -Rows @($row) -Path $tempPath -Summary @{
        Profile               = "Digital Analytics (digital_analytics)"
        TotalMatched          = 1
        ExcludedContractCount = 0
        CurrentCount          = 1
        TrackerCount          = 1
        DuplicateCount        = 0
        RemovedCount          = 0
        PreservedAppliedCount = 0
        SourceDiagnostics     = "OpenXML workbook test"
        DryRun                = "no"
        DiagnosticMode        = "no"
    }

    Assert-OpenXmlWorkbook -Condition (Test-Path -LiteralPath $tempPath) -Message "Expected XLSX file to be created."
    $rows = @(Import-TrackerRowsFromOpenXmlXlsx -Path $tempPath)
    Assert-OpenXmlWorkbook -Condition ($rows.Count -eq 1) -Message "Expected one imported row."
    Assert-OpenXmlWorkbook -Condition ((Get-RowValue -Row $rows[0] -Name "job_title") -eq "Web Analyst") -Message "Expected job title to survive round trip."
    Assert-OpenXmlWorkbook -Condition ((Get-RowValue -Row $rows[0] -Name "status") -eq "interesting") -Message "Expected status to survive round trip."
    Assert-OpenXmlWorkbook -Condition ((Get-RowValue -Row $rows[0] -Name "job_url_raw") -eq "https://example.com/job") -Message "Expected raw URL to survive round trip."

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($tempPath)
    try {
        $sheetEntry = $archive.GetEntry("xl/worksheets/sheet1.xml")
        $reader = New-Object IO.StreamReader($sheetEntry.Open())
        try {
            $sheetXml = $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }
    }
    finally {
        $archive.Dispose()
    }
    Assert-OpenXmlWorkbook -Condition ($sheetXml -notmatch '<conditionalFormatting sqref="A2:[A-Z]{2,}\d+') -Message "Expected OpenXML workbook to avoid whole-row conditional formatting fills."
}
finally {
    Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
}

Write-Host "OpenXML workbook tests passed."
