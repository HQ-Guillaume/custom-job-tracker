function Get-JobCrawlerSourceAdapterFiles {
    param([string]$SourcesRoot)

    if ([string]::IsNullOrWhiteSpace($SourcesRoot) -or -not (Test-Path -LiteralPath $SourcesRoot)) {
        throw "Source adapter directory not found: $SourcesRoot"
    }

    $loaded = New-Object System.Collections.Generic.List[object]
    foreach ($file in @(Get-ChildItem -LiteralPath $SourcesRoot -Filter "Source.*.ps1" -File | Sort-Object Name)) {
        $loaded.Add([PSCustomObject]@{
            Name = $file.Name
            Path = $file.FullName
        }) | Out-Null
    }

    return @($loaded.ToArray())
}

function Import-JobCrawlerSourceAdapters {
    param([string]$SourcesRoot)

    # PowerShell dot-sourcing inside a function is local to that function.
    # Use Get-JobCrawlerSourceAdapterFiles and dot-source the returned paths from
    # the caller script when adapter functions need caller-scope globals.
    return @(Get-JobCrawlerSourceAdapterFiles -SourcesRoot $SourcesRoot)
}

function Test-JobCrawlerSourceContract {
    param(
        [AllowNull()][object[]]$SourceDefinitions,
        [AllowNull()][object[]]$LoadedFiles = @()
    )

    $issues = New-Object System.Collections.Generic.List[string]
    $loadedFileCount = @($LoadedFiles).Count
    if ($loadedFileCount -eq 0) {
        $issues.Add("No source adapter files were loaded from app\sources.") | Out-Null
    }

    foreach ($source in @($SourceDefinitions)) {
        if ($null -eq $source) {
            continue
        }

        $sourceKey = [string]$source.Key
        $functionName = [string]$source.CrawlFunction
        if ([string]::IsNullOrWhiteSpace($sourceKey)) {
            $issues.Add("A configured source has an empty key.") | Out-Null
            continue
        }
        if ([string]::IsNullOrWhiteSpace($functionName)) {
            $issues.Add("Source '$sourceKey' has no crawl function.") | Out-Null
            continue
        }

        $command = Get-Command $functionName -CommandType Function -ErrorAction SilentlyContinue
        if ($null -eq $command) {
            $issues.Add("Source '$sourceKey' points to missing crawl function '$functionName'.") | Out-Null
        }
    }

    return [PSCustomObject]@{
        IsValid     = ($issues.Count -eq 0)
        Issues      = @($issues.ToArray())
        LoadedFiles = @($LoadedFiles)
    }
}

function Assert-JobCrawlerSourceContract {
    param(
        [AllowNull()][object[]]$SourceDefinitions,
        [AllowNull()][object[]]$LoadedFiles = @()
    )

    $result = Test-JobCrawlerSourceContract -SourceDefinitions $SourceDefinitions -LoadedFiles $LoadedFiles
    if (-not $result.IsValid) {
        throw ("Invalid source adapter contract:`n- {0}" -f (($result.Issues) -join "`n- "))
    }

    return $result
}

function Invoke-ConfiguredCrawlerSource {
    param(
        [object]$SourceDefinition,
        [System.Collections.Generic.List[object]]$Target
    )

    if ($null -eq $SourceDefinition -or [string]::IsNullOrWhiteSpace([string]$SourceDefinition.CrawlFunction)) {
        return 0
    }

    $command = Get-Command ([string]$SourceDefinition.CrawlFunction) -CommandType Function -ErrorAction SilentlyContinue
    if ($null -eq $command) {
        Write-RunStatus ("Source '{0}' skipped because crawl function '{1}' was not found." -f $SourceDefinition.Key, $SourceDefinition.CrawlFunction) "WARN"
        return 0
    }

    $rows = @(& $command.Name)
    foreach ($row in $rows) {
        $Target.Add($row) | Out-Null
    }

    return $rows.Count
}
