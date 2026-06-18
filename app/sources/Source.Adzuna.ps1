function Get-AdzunaContractType {
    param([AllowNull()]$Job)

    $contractType = ConvertTo-MatchText (Get-ObjectPropertyValue -Object $Job -Names @("contract_type"))
    $contractTime = ConvertTo-MatchText (Get-ObjectPropertyValue -Object $Job -Names @("contract_time"))

    if ($contractType -match "permanent") {
        return "Permanent"
    }
    if ($contractType -match "contract|freelance") {
        return "Freelance"
    }
    if ($contractTime -match "full\s*time|full_time") {
        return "Full-time"
    }
    if ($contractTime -match "part\s*time|part_time") {
        return "Part-time"
    }

    return ""
}

function Get-AdzunaLocation {
    param([AllowNull()]$Job)

    $location = Get-ObjectPropertyValue -Object $Job -Names @("location")
    $displayName = Get-ObjectPropertyValue -Object $location -Names @("display_name")
    if (-not [string]::IsNullOrWhiteSpace($displayName)) {
        return [string]$displayName
    }

    $area = Get-ObjectPropertyValue -Object $location -Names @("area")
    return ConvertTo-LocationText $area
}

function Get-AdzunaCompanyName {
    param([AllowNull()]$Job)

    $company = Get-ObjectPropertyValue -Object $Job -Names @("company")
    $companyName = Get-ObjectPropertyValue -Object $company -Names @("display_name", "name")
    if (-not [string]::IsNullOrWhiteSpace($companyName)) {
        return [string]$companyName
    }

    return ""
}

function Get-AdzunaJobs {
    if ([string]::IsNullOrWhiteSpace($AdzunaAppId) -or [string]::IsNullOrWhiteSpace($AdzunaAppKey)) {
        Write-RunStatus "Adzuna credentials not set; skipping Adzuna source. Set ADZUNA_APP_ID and ADZUNA_APP_KEY to enable it."
        return @()
    }

    Set-RunWindowTitle "Custom Job Tracker - Adzuna"
    Write-RunStatus "Collecting Adzuna jobs through the official API..."
    Write-RunStatus ("Adzuna plan: {0} query/queries, up to {1} page(s) each." -f $AdzunaQueries.Count, $MaxAdzunaPages)
    $stats = Start-SourceStats "Adzuna"
    $results = New-Object System.Collections.Generic.List[object]
    $queryIndex = 0

    foreach ($query in $AdzunaQueries) {
        $queryIndex++
        Write-RunStatus ("Adzuna query {0}/{1}: {2}" -f $queryIndex, $AdzunaQueries.Count, $query)
        for ($page = 1; $page -le $MaxAdzunaPages; $page++) {
            $params = @{
                app_id           = $AdzunaAppId
                app_key          = $AdzunaAppKey
                results_per_page = "25"
                what             = $query
                where            = $Location
                max_days_old     = [string][Math]::Abs($DaysBack)
                sort_by          = "date"
                "content-type"   = "application/json"
            }
            $template = [string](Get-ConfigPathValue -Object $script:JobCrawlerSourcesConfig -Path "endpoints.adzuna_jobs" -DefaultValue "https://api.adzuna.com/v1/api/jobs/fr/search/{page}")
            $url = "{0}?{1}" -f $template.Replace("{page}", [string]$page), (ConvertTo-QueryString $params)

            try {
                Add-SourceMetric -Stats $stats -Name "SearchRequests"
                $response = Invoke-RestMethod -Uri $url -Method Get -Headers @{ "Accept" = "application/json" } -TimeoutSec 45
            }
            catch {
                Add-SourceMetric -Stats $stats -Name "Errors"
                Write-Warning ("Adzuna search failed for '{0}' page {1}: {2}" -f $query, $page, $_.Exception.Message)
                break
            }

            $jobArray = @()
            if ($null -ne $response -and @($response.PSObject.Properties.Name) -contains "results") {
                $jobArray = @($response.results)
            }

            if ($jobArray.Count -eq 0) {
                break
            }

            foreach ($job in $jobArray) {
                Add-SourceMetric -Stats $stats -Name "Candidates"
                $publishedAt = ConvertTo-DateTimeOffsetOrNull (Get-ObjectPropertyValue -Object $job -Names @("created"))
                if (-not (Test-IsRecent $publishedAt)) {
                    Add-SourceMetric -Stats $stats -Name "SkippedOld"
                    continue
                }

                $title = [string](Get-ObjectPropertyValue -Object $job -Names @("title"))
                $description = [string](Get-ObjectPropertyValue -Object $job -Names @("description"))
                $companyName = Get-AdzunaCompanyName $job
                $jobLocation = Get-AdzunaLocation $job
                $contractType = Get-AdzunaContractType $job
                $jobUrl = ConvertTo-CleanUrl ([string](Get-ObjectPropertyValue -Object $job -Names @("redirect_url", "adref")))
                $sourceText = Join-CleanTextParts @($title, $description, $contractType)
                if (Test-ShouldSkipEarlyByContract -ContractType $contractType -Text $sourceText -Reliable) {
                    Add-SourceMetric -Stats $stats -Name "SkippedContract"
                    continue
                }

                $match = Get-JobMatch -Title $title -Text $sourceText
                if (-not $match.IsMatch) {
                    Add-SourceMetric -Stats $stats -Name "SkippedNoMatch"
                    continue
                }

                $result = New-JobResult -Title $title -CompanyName $companyName -JobLocation $jobLocation -ContractType $contractType -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url $jobUrl -Platform "Adzuna" -PublishedAt $publishedAt -SourceText $sourceText
                if ($null -ne $result) {
                    $results.Add($result) | Out-Null
                    Add-SourceMetric -Stats $stats -Name "Matches"
                }
            }

            Write-CountProgress -Activity ("Adzuna query {0}/{1}" -f $queryIndex, $AdzunaQueries.Count) -Current $page -Total $MaxAdzunaPages -Found $results.Count -Every 1
            Start-Sleep -Milliseconds $AdzunaDelayMilliseconds
        }
    }

    Write-RunStatus ("Adzuna complete: {0} matching jobs." -f $results.Count)
    Complete-SourceStats $stats
    return $results.ToArray()
}

