function Get-HelloWorkSearchUrl {
    param(
        [string]$Query,
        [int]$Page
    )

    $params = @{
        k = $Query
    }
    if (-not [string]::IsNullOrWhiteSpace($Location) -and $Location -notmatch "(?i)^france$") {
        $params["l"] = $Location
    }
    if ($Page -gt 1) {
        $params["p"] = [string]$Page
    }

    $baseUrl = [string](Get-ConfigPathValue -Object $script:JobCrawlerSourcesConfig -Path "endpoints.hellowork_search" -DefaultValue "https://www.hellowork.com/fr-fr/emploi/recherche.html")
    return "{0}?{1}" -f $baseUrl, (ConvertTo-QueryString $params)
}

function Get-HelloWorkJsonObjects {
    param([AllowNull()][string]$Html)

    $objects = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrWhiteSpace($Html)) {
        return $objects.ToArray()
    }

    $scripts = [regex]::Matches($Html, '(?is)<script[^>]*type=["'']application/ld\+json["''][^>]*>(?<json>.*?)</script>')
    foreach ($script in $scripts) {
        $jsonText = [System.Net.WebUtility]::HtmlDecode($script.Groups["json"].Value).Trim()
        if ([string]::IsNullOrWhiteSpace($jsonText)) {
            continue
        }

        try {
            $parsed = $jsonText | ConvertFrom-Json -ErrorAction Stop
        }
        catch {
            continue
        }

        if ($parsed -is [System.Collections.IEnumerable] -and $parsed -isnot [string] -and $parsed -isnot [pscustomobject]) {
            foreach ($item in @($parsed)) {
                if ($null -ne $item) {
                    $objects.Add($item) | Out-Null
                }
            }
        }
        else {
            $objects.Add($parsed) | Out-Null
        }
    }

    return $objects.ToArray()
}

function Get-HelloWorkJobMetadata {
    param([AllowNull()][string]$Html)

    $title = ""
    $company = ""
    $location = ""
    $contract = ""
    $description = ""
    $datePosted = $null
    $employmentType = ""

    foreach ($jsonObject in (Get-HelloWorkJsonObjects -Html $Html)) {
        $objectType = [string](Get-ObjectPropertyValue -Object $jsonObject -Names @("@type", "type"))
        if ($objectType -eq "JobPosting") {
            $title = Get-PreferredValue (Repair-DisplayText ([string](Get-ObjectPropertyValue -Object $jsonObject -Names @("title", "name")))) $title
            $description = Get-PreferredValue (ConvertFrom-HtmlText ([string](Get-ObjectPropertyValue -Object $jsonObject -Names @("description")))) $description
            $datePostedValue = Get-ObjectPropertyValue -Object $jsonObject -Names @("datePosted")
            if ($null -eq $datePosted -and $null -ne $datePostedValue) {
                $datePosted = ConvertTo-DateTimeOffsetOrNull $datePostedValue
            }

            $organization = Get-ObjectPropertyValue -Object $jsonObject -Names @("hiringOrganization")
            if ($null -ne $organization) {
                $company = Get-PreferredValue (Repair-DisplayText ([string](Get-ObjectPropertyValue -Object $organization -Names @("name")))) $company
            }

            $jobLocationValue = Get-ObjectPropertyValue -Object $jsonObject -Names @("jobLocation")
            $location = Get-PreferredValue (ConvertTo-LocationText $jobLocationValue) $location
            $employmentType = Get-PreferredValue (Repair-DisplayText ([string](Get-ObjectPropertyValue -Object $jsonObject -Names @("employmentType")))) $employmentType
        }

        $title = Get-PreferredValue (Repair-DisplayText ([string](Get-ObjectPropertyValue -Object $jsonObject -Names @("JobTitle")))) $title
        $company = Get-PreferredValue (Repair-DisplayText ([string](Get-ObjectPropertyValue -Object $jsonObject -Names @("Company")))) $company
        $location = Get-PreferredValue (Repair-DisplayText ([string](Get-ObjectPropertyValue -Object $jsonObject -Names @("Localisation")))) $location
        $contract = Get-PreferredValue (Repair-DisplayText ([string](Get-ObjectPropertyValue -Object $jsonObject -Names @("ContractType")))) $contract
        $description = Get-PreferredValue (ConvertFrom-HtmlText ([string](Get-ObjectPropertyValue -Object $jsonObject -Names @("Description")))) $description
    }

    if ([string]::IsNullOrWhiteSpace($title)) {
        $title = Get-TitleFromHtml $Html
    }

    [PSCustomObject]@{
        Title          = $title
        Company        = $company
        Location       = $location
        Contract       = $contract
        Description    = $description
        DatePosted     = $datePosted
        EmploymentType = $employmentType
    }
}

function Get-HelloWorkCardCandidates {
    param(
        [string]$Html,
        [string]$SearchUrl,
        [string]$Query,
        [AllowNull()]$Stats = $null
    )

    $candidates = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrWhiteSpace($Html)) {
        return $candidates.ToArray()
    }

    $cards = [regex]::Matches($Html, '(?is)<li\b[^>]*>.*?data-cy=["'']serpCard["''].*?</li>')
    $cardIndex = 0
    foreach ($card in $cards) {
        $cardIndex++
        Add-SourceMetric -Stats $Stats -Name "Candidates"
        if ($cardIndex -gt $MaxHelloWorkCardsPerQuery) {
            Add-SourceMetric -Stats $Stats -Name "SkippedByCap" -Amount ([Math]::Max(0, $cards.Count - $MaxHelloWorkCardsPerQuery))
            break
        }

        $cardHtml = $card.Value
        $linkMatch = [regex]::Match($cardHtml, '(?is)<a\b[^>]*data-cy=["'']offerTitle["''][^>]*>.*?</a>')
        if (-not $linkMatch.Success) {
            $linkMatch = [regex]::Match($cardHtml, '(?is)<a\b[^>]*href=["'']/fr-fr/emplois/\d+\.html[^>]*>.*?</a>')
        }
        if (-not $linkMatch.Success) {
            continue
        }

        $linkHtml = $linkMatch.Value
        $jobUrl = ConvertTo-AbsoluteUrl -BaseUrl $SearchUrl -Href (Get-HtmlAttributeValue -Html $linkHtml -Name "href")
        if ([string]::IsNullOrWhiteSpace($jobUrl)) {
            continue
        }

        $title = ""
        $companyName = ""
        $titleAttribute = Get-HtmlAttributeValue -Html $linkHtml -Name "title"
        if ($titleAttribute -match "^(?<title>.+?)\s+-\s+(?<company>.+)$") {
            $title = Repair-DisplayText $matches["title"]
            $companyName = Repair-DisplayText $matches["company"]
        }
        if ([string]::IsNullOrWhiteSpace($title)) {
            $titleMatch = [regex]::Match($linkHtml, '(?is)<p[^>]*class=["''][^"'']*typo-l[^"'']*["''][^>]*>(?<title>.*?)</p>')
            if ($titleMatch.Success) {
                $title = ConvertFrom-HtmlText $titleMatch.Groups["title"].Value
            }
        }
        if ([string]::IsNullOrWhiteSpace($companyName)) {
            $paragraphs = @([regex]::Matches($linkHtml, '(?is)<p[^>]*>(?<text>.*?)</p>'))
            if ($paragraphs.Count -gt 1) {
                $companyName = ConvertFrom-HtmlText $paragraphs[1].Groups["text"].Value
            }
        }

        $location = ""
        $locationMatch = [regex]::Match($cardHtml, '(?is)data-cy=["'']localisationCard["''][^>]*>\s*(?<value>.*?)\s*</div>')
        if ($locationMatch.Success) {
            $location = ConvertFrom-HtmlText $locationMatch.Groups["value"].Value
        }

        $contractType = ""
        $contractMatch = [regex]::Match($cardHtml, '(?is)data-cy=["'']contractCard["''][^>]*>\s*(?<value>.*?)\s*</div>')
        if ($contractMatch.Success) {
            $contractType = ConvertFrom-HtmlText $contractMatch.Groups["value"].Value
        }

        $cardText = ConvertFrom-HtmlText $cardHtml
        if (Test-ShouldSkipEarlyByContract -ContractType $contractType -Text (Join-CleanTextParts @($title, $cardText)) -Reliable) {
            Add-SourceMetric -Stats $Stats -Name "SkippedContract"
            continue
        }

        $publishedAt = ConvertFrom-FrenchRelativeDateText $cardText
        if ($null -ne $publishedAt -and -not (Test-IsRecent $publishedAt)) {
            Add-SourceMetric -Stats $Stats -Name "SkippedOld"
            continue
        }

        $actualCandidateText = Join-CleanTextParts @($title, $companyName, $location, $contractType, $cardText)
        $rankingText = Join-CleanTextParts @($actualCandidateText, ("search query {0}" -f $Query))
        $actualMatch = Get-JobMatch -Title $title -Text $actualCandidateText
        $rankingMatch = Get-JobMatch -Title $title -Text $rankingText
        if (-not $actualMatch.IsMatch -and $actualCandidateText -notmatch $WttjUrlCandidatePattern -and $Query -notmatch $WttjUrlCandidatePattern) {
            Add-SourceMetric -Stats $Stats -Name "SkippedNoMatch"
            continue
        }

        $cardScore = $(if ($actualMatch.IsMatch) { $actualMatch.Score } elseif ($rankingMatch.IsMatch) { [Math]::Min(45, [int]$rankingMatch.Score) } else { 10 })
        $candidates.Add([PSCustomObject]@{
            Url          = $jobUrl
            Title        = $title
            Company      = $companyName
            Location     = $location
            Contract     = $contractType
            PublishedAt  = $publishedAt
            CardText     = $actualCandidateText
            Query        = $Query
            CardScore    = [int]$cardScore
            CardPosition = $cardIndex
        }) | Out-Null
    }

    return $candidates.ToArray()
}

function Get-HelloWorkJobs {
    Set-RunWindowTitle "Custom Job Tracker - HelloWork"
    Write-RunStatus "Collecting HelloWork jobs from public search pages..."
    Write-RunStatus ("HelloWork plan: {0} query/queries, {1} page(s) each, then at most {2} unique detail page(s)." -f $HelloWorkQueries.Count, $MaxHelloWorkPages, $MaxHelloWorkDetails)

    $stats = Start-SourceStats "HelloWork"
    $results = New-Object System.Collections.Generic.List[object]
    $candidateByUrl = @{}
    $queryIndex = 0

    foreach ($query in $HelloWorkQueries) {
        $queryIndex++
        Write-RunStatus ("HelloWork query {0}/{1}: {2}" -f $queryIndex, $HelloWorkQueries.Count, $query)
        for ($page = 1; $page -le $MaxHelloWorkPages; $page++) {
            $searchUrl = Get-HelloWorkSearchUrl -Query $query -Page $page
            try {
                Add-SourceMetric -Stats $stats -Name "SearchRequests"
                $html = Invoke-TextRequest $searchUrl -Headers @{ "Accept" = "text/html,application/xhtml+xml" } -TimeoutSec 30
            }
            catch {
                Add-SourceMetric -Stats $stats -Name "Errors"
                Write-Warning ("HelloWork search failed for '{0}' page {1}: {2}" -f $query, $page, $_.Exception.Message)
                break
            }

            $candidates = @(Get-HelloWorkCardCandidates -Html $html -SearchUrl $searchUrl -Query $query -Stats $stats)
            foreach ($candidate in $candidates) {
                if (-not $candidateByUrl.ContainsKey($candidate.Url) -or [int]$candidate.CardScore -gt [int]$candidateByUrl[$candidate.Url].CardScore) {
                    $candidateByUrl[$candidate.Url] = $candidate
                }
            }

            Write-CountProgress -Activity ("HelloWork search query {0}/{1}" -f $queryIndex, $HelloWorkQueries.Count) -Current $page -Total $MaxHelloWorkPages -Found $candidateByUrl.Count -Every 1
            if ($candidates.Count -eq 0) {
                break
            }
            Start-Sleep -Milliseconds $HelloWorkSearchDelayMilliseconds
        }
    }

    $selectedCandidates = @($candidateByUrl.Values |
        Sort-Object -Property `
            @{ Expression = "CardScore"; Descending = $true },
            @{ Expression = { if ($null -ne $_.PublishedAt) { $_.PublishedAt } else { [DateTimeOffset]::MinValue } }; Descending = $true },
            @{ Expression = "CardPosition"; Descending = $false } |
        Select-Object -First $MaxHelloWorkDetails)
    Add-SourceMetric -Stats $stats -Name "SelectedDetails" -Amount $selectedCandidates.Count
    Add-SourceMetric -Stats $stats -Name "SkippedByCap" -Amount ([Math]::Max(0, $candidateByUrl.Count - $selectedCandidates.Count))

    Write-RunStatus ("HelloWork candidates selected: {0} unique detail page(s) from {1} candidate(s)." -f $selectedCandidates.Count, $candidateByUrl.Count)
    $detailIndex = 0
    foreach ($candidate in $selectedCandidates) {
        $detailIndex++
        Write-CountProgress -Activity "HelloWork detail pages" -Current $detailIndex -Total $selectedCandidates.Count -Found $results.Count -Every 5

        try {
            $html = Invoke-CachedTextRequest -Url $candidate.Url -CacheScope "hellowork-detail" -Headers @{ "Accept" = "text/html,application/xhtml+xml" } -TimeoutSec 30 -Stats $stats
        }
        catch {
            Add-SourceMetric -Stats $stats -Name "Errors"
            Write-Warning ("HelloWork detail failed for '{0}': {1}" -f $candidate.Url, $_.Exception.Message)
            continue
        }

        $metadata = Get-HelloWorkJobMetadata -Html $html
        $title = Get-PreferredValue $metadata.Title $candidate.Title
        $companyName = Get-PreferredValue $metadata.Company $candidate.Company
        $jobLocation = Get-PreferredValue $metadata.Location $candidate.Location
        $pageTitle = Get-TitleFromHtml $html
        $sourceText = Join-CleanTextParts @($title, $companyName, $jobLocation, $metadata.Contract, $metadata.Description, $candidate.CardText, $pageTitle)
        $contractType = Get-ContractTypeFromText -Text $sourceText
        if ([string]::IsNullOrWhiteSpace($contractType)) {
            $contractType = Get-PreferredValue $metadata.Contract $candidate.Contract
        }
        if ([string]::IsNullOrWhiteSpace($contractType)) {
            $contractType = Get-ContractTypeFromText -Text $sourceText -RawContractType $metadata.EmploymentType
        }

        $publishedAt = $metadata.DatePosted
        if ($null -eq $publishedAt) {
            $publishedAt = $candidate.PublishedAt
        }
        if (-not (Test-IsRecent $publishedAt)) {
            Add-SourceMetric -Stats $stats -Name "SkippedOld"
            continue
        }

        $match = Get-JobMatch -Title $title -Text $sourceText
        if (-not $match.IsMatch) {
            Add-SourceMetric -Stats $stats -Name "SkippedNoMatch"
            continue
        }

        $result = New-JobResult -Title $title -CompanyName $companyName -JobLocation $jobLocation -ContractType $contractType -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url $candidate.Url -Platform "HelloWork" -PublishedAt $publishedAt -SourceText $sourceText
        if ($null -ne $result) {
            $results.Add($result) | Out-Null
            Add-SourceMetric -Stats $stats -Name "Matches"
        }

        Start-Sleep -Milliseconds $HelloWorkDetailDelayMilliseconds
    }

    Write-RunStatus ("HelloWork complete: {0} matching jobs." -f $results.Count)
    Complete-SourceStats $stats
    return $results.ToArray()
}

