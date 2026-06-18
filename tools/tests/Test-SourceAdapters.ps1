[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
. (Join-Path $projectRoot "app\core\JobTracker.Common.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Config.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Runtime.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.SourceAdapter.ps1")

function Assert-SourceAdapter {
    param(
        [bool]$Condition,
        [string]$Message
    )

    if (-not $Condition) {
        throw "Source adapter test failed: $Message"
    }
}

$config = Get-JobCrawlerConfig -ConfigDirectory (Join-Path $projectRoot "config")
$definitions = @(Get-JobCrawlerSourceDefinitions -SourcesConfig $config.Sources)
$loaded = @(Get-JobCrawlerSourceAdapterFiles -SourcesRoot (Join-Path $projectRoot "app\sources"))
foreach ($sourceAdapter in $loaded) {
    . $sourceAdapter.Path
}
$contract = Test-JobCrawlerSourceContract -SourceDefinitions $definitions -LoadedFiles $loaded

Assert-SourceAdapter -Condition ($loaded.Count -ge 6) -Message "Expected source auto-loader to load every Source.*.ps1 adapter."
Assert-SourceAdapter -Condition $contract.IsValid -Message (($contract.Issues) -join "; ")
Assert-SourceAdapter -Condition (@($definitions | Where-Object { $_.Key -eq "linkedin" -and (Get-Command $_.CrawlFunction -CommandType Function -ErrorAction SilentlyContinue) }).Count -eq 1) -Message "Expected LinkedIn adapter function to be loaded dynamically."

$badDefinition = [PSCustomObject]@{
    Key = "broken_board"
    CrawlFunction = "Get-DefinitelyMissingJobs"
}
$badContract = Test-JobCrawlerSourceContract -SourceDefinitions @($badDefinition) -LoadedFiles $loaded
Assert-SourceAdapter -Condition (-not $badContract.IsValid -and (($badContract.Issues -join " ") -match "Get-DefinitelyMissingJobs")) -Message "Expected contract check to report missing crawl functions."

Write-Host "Source adapter tests passed."
