function Get-LinkedInContractType {
    param(
        [AllowNull()][string]$Title,
        [AllowNull()][string]$DetailText
    )

    $titleContract = Get-ContractTypeFromText -Text $Title
    if (-not [string]::IsNullOrWhiteSpace($titleContract)) {
        return $titleContract
    }

    $detailMatchText = ConvertTo-MatchText $DetailText
    if ($detailMatchText -match "employment\s+type\s+full-time|type\s+d.?emploi\s+temps\s+plein") {
        return "Full-time"
    }
    if ($detailMatchText -match "employment\s+type\s+internship|type\s+d.?emploi\s+stage") {
        return "Internship"
    }
    if ($detailMatchText -match "employment\s+type\s+temporary|employment\s+type\s+contract") {
        return "CDD"
    }

    return Get-ContractTypeFromText -Text $DetailText
}

function Get-LinkedInLocationFromHtml {
    param([AllowNull()][string]$Html)

    if ([string]::IsNullOrWhiteSpace($Html)) {
        return ""
    }

    foreach ($pattern in @(
        '(?is)<span[^>]*class="[^"]*job-search-card__location[^"]*"[^>]*>(?<location>.*?)</span>',
        '(?is)<span[^>]*class="[^"]*topcard__flavor[^"]*topcard__flavor--bullet[^"]*"[^>]*>(?<location>.*?)</span>',
        '(?is)<span[^>]*class="[^"]*jobs-unified-top-card__bullet[^"]*"[^>]*>(?<location>.*?)</span>'
    )) {
        $match = [regex]::Match($Html, $pattern)
        if ($match.Success) {
            $location = ConvertFrom-HtmlText $match.Groups["location"].Value
            if (-not [string]::IsNullOrWhiteSpace($location)) {
                return $location
            }
        }
    }

    return Get-LocationFromStructuredHtml $Html
}

function Get-LinkedInJobs {
    Set-RunWindowTitle "Custom Job Tracker - LinkedIn"
    Write-RunStatus "Collecting LinkedIn jobs from public guest endpoints..."
    Write-RunStatus ("LinkedIn plan: {0} search query/queries, up to {1} page(s) each, then up to {2} ranked detail page(s)." -f $LinkedInQueries.Count, $MaxLinkedInSearchPages, $(if ($MaxLinkedInDetails -gt 0) { $MaxLinkedInDetails } else { "all" }))
    $stats = Start-SourceStats "LinkedIn"
    $results = New-Object System.Collections.Generic.List[object]
    $candidateById = @{}
    $seconds = [Math]::Max(86400, [int]([Math]::Abs($DaysBack) * 86400))

    $queryIndex = 0
    foreach ($query in $LinkedInQueries) {
        $queryIndex++
        Write-RunStatus ("LinkedIn query {0}/{1}: {2}" -f $queryIndex, $LinkedInQueries.Count, $query)
        for ($page = 0; $page -lt $MaxLinkedInSearchPages; $page++) {
            $start = $page * 25
            $params = @{
                keywords = $query
                location = $Location
                f_TPR    = "r$seconds"
                start    = [string]$start
            }
            $searchBaseUrl = [string](Get-ConfigPathValue -Object $script:JobCrawlerSourcesConfig -Path "endpoints.linkedin_search" -DefaultValue "https://www.linkedin.com/jobs-guest/jobs/api/seeMoreJobPostings/search")
            $url = "{0}?{1}" -f $searchBaseUrl, (ConvertTo-QueryString $params)

            try {
                Add-SourceMetric -Stats $stats -Name "SearchRequests"
                $html = Invoke-TextRequest $url -Headers @{ "Accept" = "text/html,*/*" } -TimeoutSec 30
            }
            catch {
                Start-Sleep -Seconds 8
                try {
                    Add-SourceMetric -Stats $stats -Name "SearchRequests"
                    $html = Invoke-TextRequest $url -Headers @{ "Accept" = "text/html,*/*" } -TimeoutSec 30
                }
                catch {
                    Add-SourceMetric -Stats $stats -Name "Errors"
                    Write-Warning ("LinkedIn search failed for '{0}' page {1}: {2}" -f $query, ($page + 1), $_.Exception.Message)
                    break
                }
            }

            if ([string]::IsNullOrWhiteSpace($html) -or $html -notmatch "jobPosting") {
                Write-RunStatus ("LinkedIn query {0}/{1}, page {2}/{3}: no result cards returned." -f $queryIndex, $LinkedInQueries.Count, ($page + 1), $MaxLinkedInSearchPages)
                break
            }

            $cards = [regex]::Matches($html, '(?is)<li>.*?data-entity-urn="urn:li:jobPosting:(?<id>\d+)".*?</li>')
            if ($cards.Count -eq 0) {
                Write-RunStatus ("LinkedIn query {0}/{1}, page {2}/{3}: no readable cards found." -f $queryIndex, $LinkedInQueries.Count, ($page + 1), $MaxLinkedInSearchPages)
                break
            }
            Write-RunStatus ("LinkedIn query {0}/{1}, page {2}/{3}: {4} card(s) found; {5} unique candidate(s) so far." -f $queryIndex, $LinkedInQueries.Count, ($page + 1), $MaxLinkedInSearchPages, $cards.Count, $candidateById.Count)

            $cardIndex = 0
            foreach ($card in $cards) {
                $cardIndex++
                Add-SourceMetric -Stats $stats -Name "Candidates"
                $cardHtml = $card.Value
                $id = $card.Groups["id"].Value

                $titleMatch = [regex]::Match($cardHtml, '(?is)<h3[^>]*class="[^"]*base-search-card__title[^"]*"[^>]*>(?<title>.*?)</h3>')
                $urlMatch = [regex]::Match($cardHtml, '(?is)<a[^>]+href="(?<url>https?://[^"]+)"')
                $dateMatch = [regex]::Match($cardHtml, '(?is)<time[^>]+datetime="(?<date>[^"]+)"')

                if (-not $titleMatch.Success -or -not $urlMatch.Success -or -not $dateMatch.Success) {
                    continue
                }

                $title = ConvertFrom-HtmlText $titleMatch.Groups["title"].Value
                $companyMatch = [regex]::Match($cardHtml, '(?is)<h4[^>]*class="[^"]*base-search-card__subtitle[^"]*"[^>]*>(?<company>.*?)</h4>')
                $companyName = ""
                if ($companyMatch.Success) {
                    $companyName = ConvertFrom-HtmlText $companyMatch.Groups["company"].Value
                }
                $jobLocation = Get-LinkedInLocationFromHtml $cardHtml

                $jobUrl = ConvertTo-CleanUrl $urlMatch.Groups["url"].Value
                $publishedAt = ConvertTo-DateTimeOffsetOrNull $dateMatch.Groups["date"].Value
                if (-not (Test-IsRecent $publishedAt)) {
                    Add-SourceMetric -Stats $stats -Name "SkippedOld"
                    continue
                }

                $cardText = Join-CleanTextParts @($title, $companyName, $jobLocation, $jobUrl, (ConvertFrom-HtmlText $cardHtml))
                if (Test-ShouldSkipEarlyByContract -Text $cardText) {
                    Add-SourceMetric -Stats $stats -Name "SkippedContract"
                    continue
                }

                $cardMatch = Get-JobMatch -Title $title -Text $cardText
                $cardScore = 10
                if ($cardMatch.IsMatch) {
                    $cardScore = [int]$cardMatch.Score
                }
                elseif ($cardText -match $WttjUrlCandidatePattern) {
                    $cardScore = 45
                }

                $candidate = [PSCustomObject]@{
                    Id           = $id
                    Title        = $title
                    Company      = $companyName
                    Location     = $jobLocation
                    Url          = $jobUrl
                    PublishedAt  = $publishedAt
                    CardText     = $cardText
                    CardScore    = [int]$cardScore
                    QueryIndex   = $queryIndex
                    Page         = $page
                    CardPosition = $cardIndex
                }

                if (-not $candidateById.ContainsKey($id) -or [int]$candidate.CardScore -gt [int]$candidateById[$id].CardScore) {
                    $candidateById[$id] = $candidate
                }
            }

            Start-Sleep -Milliseconds $LinkedInDelayMilliseconds
        }
    }

    $orderedCandidates = @($candidateById.Values |
        Sort-Object -Property `
            @{ Expression = "CardScore"; Descending = $true },
            @{ Expression = "PublishedAt"; Descending = $true },
            @{ Expression = "QueryIndex"; Descending = $false },
            @{ Expression = "Page"; Descending = $false },
            @{ Expression = "CardPosition"; Descending = $false })
    if ($MaxLinkedInDetails -gt 0) {
        $selectedCandidates = @($orderedCandidates | Select-Object -First $MaxLinkedInDetails)
    }
    else {
        $selectedCandidates = @($orderedCandidates)
    }

    Add-SourceMetric -Stats $stats -Name "SelectedDetails" -Amount $selectedCandidates.Count
    Add-SourceMetric -Stats $stats -Name "SkippedByCap" -Amount ([Math]::Max(0, $orderedCandidates.Count - $selectedCandidates.Count))
    Write-RunStatus ("LinkedIn candidates selected: {0} detail page(s) from {1} unique candidate(s)." -f $selectedCandidates.Count, $orderedCandidates.Count)

    $candidateIndex = 0
    foreach ($candidate in $selectedCandidates) {
        $candidateIndex++
        Write-CountProgress -Activity "LinkedIn detail pages" -Current $candidateIndex -Total $selectedCandidates.Count -Found $results.Count -Every 10

        $detailTemplate = [string](Get-ConfigPathValue -Object $script:JobCrawlerSourcesConfig -Path "endpoints.linkedin_detail" -DefaultValue "https://www.linkedin.com/jobs-guest/jobs/api/jobPosting/{id}")
        $detailUrl = $detailTemplate.Replace("{id}", [Uri]::EscapeDataString([string]$candidate.Id))
        $detailHtml = ""
        try {
            $cacheHitsBefore = [int]$stats["CacheHits"]
            $detailHtml = Invoke-CachedTextRequest -Url $detailUrl -CacheScope "linkedin-detail" -Headers @{ "Accept" = "text/html,*/*" } -TimeoutSec 30 -Stats $stats
            if ([int]$stats["CacheHits"] -eq $cacheHitsBefore) {
                Start-Sleep -Milliseconds $LinkedInDelayMilliseconds
            }
        }
        catch {
            Start-Sleep -Seconds 5
            try {
                $cacheHitsBefore = [int]$stats["CacheHits"]
                $detailHtml = Invoke-CachedTextRequest -Url $detailUrl -CacheScope "linkedin-detail" -Headers @{ "Accept" = "text/html,*/*" } -TimeoutSec 30 -Stats $stats
                if ([int]$stats["CacheHits"] -eq $cacheHitsBefore) {
                    Start-Sleep -Milliseconds $LinkedInDelayMilliseconds
                }
            }
            catch {
                Add-SourceMetric -Stats $stats -Name "Errors"
                $detailHtml = ""
            }
        }

        $detailText = ConvertFrom-HtmlText $detailHtml
        $combined = Join-CleanTextParts @($candidate.Title, $candidate.Url, $candidate.CardText, $detailText)
        $match = Get-JobMatch -Title $candidate.Title -Text $combined
        if (-not $match.IsMatch) {
            Add-SourceMetric -Stats $stats -Name "SkippedNoMatch"
            continue
        }

        $companyName = $candidate.Company
        if ([string]::IsNullOrWhiteSpace($companyName)) {
            $detailCompanyMatch = [regex]::Match($detailHtml, '(?is)<a[^>]*class="[^"]*topcard__org-name-link[^"]*"[^>]*>(?<company>.*?)</a>')
            if ($detailCompanyMatch.Success) {
                $companyName = ConvertFrom-HtmlText $detailCompanyMatch.Groups["company"].Value
            }
        }

        $jobLocation = $candidate.Location
        if ([string]::IsNullOrWhiteSpace($jobLocation)) {
            $jobLocation = Get-LinkedInLocationFromHtml $detailHtml
        }

        $contractType = Get-LinkedInContractType -Title $candidate.Title -DetailText $detailText
        if (Test-IsExcludedContractType $contractType) {
            Add-SourceMetric -Stats $stats -Name "SkippedContract"
            continue
        }

        $result = New-JobResult -Title $candidate.Title -CompanyName $companyName -JobLocation $jobLocation -ContractType $contractType -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url $candidate.Url -Platform "LinkedIn" -PublishedAt $candidate.PublishedAt -SourceText $combined
        if ($null -ne $result) {
            $results.Add($result) | Out-Null
            Add-SourceMetric -Stats $stats -Name "Matches"
        }
    }

    Write-RunStatus ("LinkedIn complete: {0} matching jobs." -f $results.Count)
    Complete-SourceStats $stats
    return $results.ToArray()
}

