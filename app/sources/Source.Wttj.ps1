# Auto-extracted from Find-AnalyticsJobs.ps1. Keep dot-sourced execution order in the main script.

function Get-WttjLocationFromUrl {
    param([AllowNull()][string]$Url)

    if ([string]::IsNullOrWhiteSpace($Url)) {
        return ""
    }

    $path = $Url
    try {
        $path = ([Uri]$Url).AbsolutePath
    }
    catch {
    }

    $slugMatch = [regex]::Match($path, "/jobs/(?<slug>[^/?#]+)", [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $slugMatch.Success) {
        return ""
    }

    $slugParts = @(([string]$slugMatch.Groups["slug"].Value) -split "_")
    if ($slugParts.Count -lt 2) {
        return ""
    }

    foreach ($locationSlug in @($slugParts | Select-Object -Skip 1)) {
        if ([string]::IsNullOrWhiteSpace($locationSlug)) {
            continue
        }
        if ($locationSlug -match "^(h|f|m|x|nb|stage|internship|cdi|cdd)$") {
            continue
        }
        if ($locationSlug -cmatch "^[A-Z0-9]{3,12}$") {
            continue
        }

        $location = ConvertFrom-SlugToTitle $locationSlug
        if (-not (Test-IsJunkLocationText $location)) {
            return $location
        }
    }

    return ""
}

function ConvertFrom-WttjJsonStringValue {
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ""
    }

    try {
        return Repair-DisplayText ([string](('"{0}"' -f ($Value -replace '"', '\"')) | ConvertFrom-Json))
    }
    catch {
        return Repair-DisplayText $Value
    }
}

function Get-WttjInitialDataJsonText {
    param([AllowNull()][string]$Html)

    if ([string]::IsNullOrWhiteSpace($Html)) {
        return ""
    }

    $match = [regex]::Match($Html, '(?s)window\.__INITIAL_DATA__\s*=\s*"(?<data>(?:\\.|[^"\\])*)"\s*;')
    if (-not $match.Success) {
        return $Html
    }

    try {
        return [string](('"{0}"' -f $match.Groups["data"].Value) | ConvertFrom-Json)
    }
    catch {
        return $Html
    }
}

function Get-WttjJsonStringField {
    param(
        [AllowNull()][string]$Text,
        [string]$Name
    )

    if ([string]::IsNullOrWhiteSpace($Text) -or [string]::IsNullOrWhiteSpace($Name)) {
        return ""
    }

    $pattern = '(?s)(?:"|\\")' + [regex]::Escape($Name) + '(?:"|\\")\s*:\s*(?:"|\\")(?<value>.*?)(?:"|\\")'
    $match = [regex]::Match($Text, $pattern, [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $match.Success) {
        return ""
    }

    return ConvertFrom-WttjJsonStringValue $match.Groups["value"].Value
}

function ConvertTo-WttjCountryLabel {
    param([AllowNull()][string]$CountryCode)

    $code = (ConvertTo-MatchText $CountryCode).ToUpperInvariant()
    switch ($code) {
        "FR" { return "France" }
        "GB" { return "United Kingdom" }
        "UK" { return "United Kingdom" }
        "US" { return "United States" }
        "MA" { return "Morocco" }
        "ES" { return "Spain" }
        "CA" { return "Canada" }
        "BE" { return "Belgium" }
        "CH" { return "Switzerland" }
        "DE" { return "Germany" }
        "NL" { return "Netherlands" }
        "IE" { return "Ireland" }
        "LU" { return "Luxembourg" }
        default { return $CountryCode }
    }
}

function Get-WttjLocationFromInitialData {
    param([AllowNull()][string]$Html)

    $jsonText = Get-WttjInitialDataJsonText $Html
    if ([string]::IsNullOrWhiteSpace($jsonText)) {
        return ""
    }

    $officesMatch = [regex]::Match($jsonText, '(?s)(?:"|\\")offices(?:"|\\")\s*:\s*\[(?<offices>.*?)\]', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $officesMatch.Success) {
        return ""
    }

    $locations = New-Object System.Collections.Generic.List[string]
    foreach ($officeMatch in [regex]::Matches($officesMatch.Groups["offices"].Value, '(?s)\{(?<office>.*?)\}')) {
        $officeText = $officeMatch.Groups["office"].Value
        $city = Get-WttjJsonStringField -Text $officeText -Name "city"
        if ([string]::IsNullOrWhiteSpace($city)) {
            $city = Get-WttjJsonStringField -Text $officeText -Name "local_city"
        }

        $countryCode = Get-WttjJsonStringField -Text $officeText -Name "country_code"
        $country = ConvertTo-WttjCountryLabel $countryCode
        if ([string]::IsNullOrWhiteSpace($country)) {
            $country = Get-WttjJsonStringField -Text $officeText -Name "country"
        }

        $location = Join-CleanTextParts @($city, $country)
        if (-not [string]::IsNullOrWhiteSpace($location) -and -not $locations.Contains($location)) {
            $locations.Add($location) | Out-Null
        }
    }

    return Join-CleanTextParts ($locations.ToArray() | Select-Object -First 3)
}

function Test-IsWttjFranceTarget {
    $target = ConvertTo-MatchText $Location
    return ([string]::IsNullOrWhiteSpace($target) -or $target -match "france|paris|ile\s*de\s*france")
}

function Test-IsWttjKnownForeignLocation {
    param([AllowNull()][string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $false
    }

    $foreignPatterns = Get-JobCrawlerPreferenceArray -Preferences $script:JobCrawlerPreferences -Name "foreign_location_patterns"
    return (Test-AnyPatternMatch -Text (ConvertTo-MatchText $Text) -Patterns $foreignPatterns)
}

function Test-IsWttjLocationAllowed {
    param(
        [AllowNull()][string]$JobLocation,
        [AllowNull()][string]$Url,
        [AllowNull()][string]$Text = ""
    )

    if (-not (Test-IsWttjFranceTarget)) {
        return $true
    }

    $urlLocation = Get-WttjLocationFromUrl $Url
    if (Test-IsWttjKnownForeignLocation $urlLocation) {
        return $false
    }

    if ([string]::IsNullOrWhiteSpace($JobLocation)) {
        return $false
    }

    if (Test-IsWttjKnownForeignLocation $JobLocation) {
        return $false
    }

    $category = Get-LocationFitCategory -Location $JobLocation -Preferences $script:JobCrawlerPreferences
    if ($category -eq "foreign") {
        return $false
    }

    return $true
}

function Get-WttjLocation {
    param(
        [AllowNull()][string]$Html,
        [AllowNull()][string]$Url,
        [AllowNull()][string]$Title
    )

    $location = Get-LocationFromStructuredHtml $Html
    if (-not [string]::IsNullOrWhiteSpace($location)) {
        return $location
    }

    $location = Get-WttjLocationFromInitialData $Html
    if (-not [string]::IsNullOrWhiteSpace($location)) {
        return $location
    }

    $location = Get-LocationFromText $Title
    if (-not [string]::IsNullOrWhiteSpace($location)) {
        return $location
    }

    return Get-WttjLocationFromUrl $Url
}

function Get-WelcomeKitLocation {
    param(
        [AllowNull()]$Job,
        [AllowNull()][string]$JobUrl
    )

    foreach ($propertyName in @("location", "locations", "office", "offices", "address", "addresses", "workplace", "workplaces")) {
        $location = ConvertTo-LocationText (Get-ObjectPropertyValue -Object $Job -Names @($propertyName))
        if (-not [string]::IsNullOrWhiteSpace($location)) {
            return $location
        }
    }

    return Get-WttjLocationFromUrl $JobUrl
}

function Get-WttjCompanyNameFromUrl {
    param([string]$Url)

    $companyMatch = [regex]::Match($Url, "/companies/(?<slug>[^/]+)/jobs/", [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($companyMatch.Success) {
        return ConvertFrom-SlugToTitle $companyMatch.Groups["slug"].Value
    }

    return ""
}

function Get-WttjCandidateScore {
    param([string]$Url)

    if ((Test-IsWttjFranceTarget) -and (Test-IsWttjKnownForeignLocation (Get-WttjLocationFromUrl $Url))) {
        return -1000
    }

    $score = 0
    if ($Url -match "(?i)(web[-_\s]*analyst|digital[-_\s]*analyst|web[-_\s]*analytics|digital[-_\s]*analytics|tracking|taggage|tagging|(^|[-_\s])ga4($|[-_\s])|(^|[-_\s])gtm($|[-_\s])|google[-_\s]*(analytics|tag[-_\s]*manager)|piano|contentsquare|content[-_\s]*square|(^|[-_\s])cro($|[-_\s]))") {
        $score += 100
    }
    if ($Url -match "(?i)(data[-_\s]*analyst|analytics[-_\s]*(consultant|specialist|manager|engineer))") {
        $score += 40
    }
    if ($Location -match "(?i)france|paris|french|remote" -and $Url -match "(?i)(/fr/|_paris|_puteaux|_levallois|_boulogne|_lille|_lyon|_fr\b)") {
        $score += 75
    }
    elseif ($Url -match "/fr/") {
        $score += 10
    }

    return $score
}

function Get-WelcomeKitCompanyName {
    param($Job, [string]$JobUrl)

    if ($Job.PSObject.Properties.Name -contains "organization" -and $null -ne $Job.organization) {
        if ($Job.organization.PSObject.Properties.Name -contains "name" -and -not [string]::IsNullOrWhiteSpace($Job.organization.name)) {
            return [string]$Job.organization.name
        }
        if ($Job.organization.PSObject.Properties.Name -contains "slug" -and -not [string]::IsNullOrWhiteSpace($Job.organization.slug)) {
            return ConvertFrom-SlugToTitle $Job.organization.slug
        }
    }

    return Get-WttjCompanyNameFromUrl $JobUrl
}

function Get-WelcomeKitJobUrl {
    param($Job)

    if ($Job.PSObject.Properties.Name -contains "websites" -and $null -ne $Job.websites) {
        foreach ($site in @($Job.websites)) {
            if ($site.PSObject.Properties.Name -contains "url" -and $site.url -match "welcometothejungle") {
                return [string]$site.url
            }
        }

        foreach ($site in @($Job.websites)) {
            if ($site.PSObject.Properties.Name -contains "url" -and -not [string]::IsNullOrWhiteSpace($site.url)) {
                return [string]$site.url
            }
        }
    }

    if ($Job.PSObject.Properties.Name -contains "apply_url" -and -not [string]::IsNullOrWhiteSpace($Job.apply_url)) {
        return [string]$Job.apply_url
    }

    if ($Job.PSObject.Properties.Name -contains "reference" -and -not [string]::IsNullOrWhiteSpace($Job.reference)) {
        $template = [string](Get-ConfigPathValue -Object $script:JobCrawlerSourcesConfig -Path "endpoints.wttj_job_page" -DefaultValue "https://www.welcometothejungle.com/fr/jobs/{reference}")
        return $template.Replace("{reference}", [Uri]::EscapeDataString([string]$Job.reference))
    }

    return ""
}

function Get-WelcomeKitJobDetails {
    param(
        [string]$Reference,
        [hashtable]$Headers
    )

    if ([string]::IsNullOrWhiteSpace($Reference)) {
        return $null
    }

    $params = @{
        websites     = "true"
        organization = "true"
    }
    $template = [string](Get-ConfigPathValue -Object $script:JobCrawlerSourcesConfig -Path "endpoints.welcome_kit_job_detail" -DefaultValue "https://www.welcomekit.co/api/v1/external/jobs/{reference}")
    $url = "{0}?{1}" -f $template.Replace("{reference}", [Uri]::EscapeDataString($Reference)), (ConvertTo-QueryString $params)

    try {
        return Invoke-RestMethod -Uri $url -Headers $Headers -Method Get -TimeoutSec 45
    }
    catch {
        return $null
    }
}

function Get-WelcomeKitJobs {
    if ([string]::IsNullOrWhiteSpace($WelcomeKitApiKey)) {
        Write-RunStatus "WelcomeKit API key not set; using WTTJ public sitemap fallback."
        return @()
    }

    Set-RunWindowTitle "Analytics Job Crawler - WTTJ API"
    Write-RunStatus "Collecting Welcome to the Jungle jobs through the official WelcomeKit API..."
    $stats = Start-SourceStats "WelcomeKit"
    $results = New-Object System.Collections.Generic.List[object]
    $headers = @{
        "Authorization" = "Bearer $WelcomeKitApiKey"
        "Accept"        = "application/json"
    }

    for ($page = 1; $page -le $MaxWelcomeKitPages; $page++) {
        Write-RunStatus ("WelcomeKit API page {0}/{1}; {2} matches so far." -f $page, $MaxWelcomeKitPages, $results.Count)
        $params = @{
            status          = "published"
            websites        = "true"
            per_page        = "100"
            page            = [string]$page
            published_after = $CutoffDate
        }
        $baseUrl = [string](Get-ConfigPathValue -Object $script:JobCrawlerSourcesConfig -Path "endpoints.welcome_kit_jobs" -DefaultValue "https://www.welcomekit.co/api/v1/external/jobs/all")
        $url = "{0}?{1}" -f $baseUrl, (ConvertTo-QueryString $params)

        try {
            Add-SourceMetric -Stats $stats -Name "SearchRequests"
            $jobs = Invoke-RestMethod -Uri $url -Headers $headers -Method Get -TimeoutSec 45
        }
        catch {
            Add-SourceMetric -Stats $stats -Name "Errors"
            Write-Warning ("WelcomeKit API call failed on page {0}: {1}" -f $page, $_.Exception.Message)
            break
        }

        $jobArray = @($jobs)
        if ($jobArray.Count -eq 0) {
            break
        }

        foreach ($job in $jobArray) {
            Add-SourceMetric -Stats $stats -Name "Candidates"
            $publishedAt = ConvertTo-DateTimeOffsetOrNull $job.published_at
            if (-not (Test-IsRecent $publishedAt)) {
                Add-SourceMetric -Stats $stats -Name "SkippedOld"
                continue
            }

            $combined = ConvertFrom-HtmlText ("{0} {1} {2} {3}" -f $job.name, $job.profile, $job.description, $job.company_description)
            $contractType = Get-ContractTypeFromText -Text $combined -RawContractType ([string]$job.contract_type)
            if (Test-ShouldSkipEarlyByContract -ContractType $contractType -Text $combined -Reliable) {
                Add-SourceMetric -Stats $stats -Name "SkippedContract"
                continue
            }

            $match = Get-JobMatch -Title ([string]$job.name) -Text $combined
            if (-not $match.IsMatch) {
                Add-SourceMetric -Stats $stats -Name "SkippedNoMatch"
                continue
            }

            $jobUrl = Get-WelcomeKitJobUrl $job
            $jobDetails = $null
            if ($job.PSObject.Properties.Name -contains "reference") {
                Add-SourceMetric -Stats $stats -Name "DetailRequests"
                $jobDetails = Get-WelcomeKitJobDetails -Reference ([string]$job.reference) -Headers $headers
            }
            if ($null -ne $jobDetails) {
                $jobUrl = Get-WelcomeKitJobUrl $jobDetails
                $job = $jobDetails
            }

            $companyName = Get-WelcomeKitCompanyName -Job $job -JobUrl $jobUrl
            $jobLocation = Get-WelcomeKitLocation -Job $job -JobUrl $jobUrl
            if (-not (Test-IsWttjLocationAllowed -JobLocation $jobLocation -Url $jobUrl -Text $combined)) {
                Add-SourceMetric -Stats $stats -Name "SkippedNoMatch"
                continue
            }

            $result = New-JobResult -Title ([string]$job.name) -CompanyName $companyName -JobLocation $jobLocation -ContractType $contractType -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url $jobUrl -Platform "Welcome to the Jungle" -PublishedAt $publishedAt -SourceText $combined
            if ($null -ne $result) {
                $results.Add($result) | Out-Null
                Add-SourceMetric -Stats $stats -Name "Matches"
            }
        }

        if ($jobArray.Count -lt 100) {
            break
        }
    }

    Write-RunStatus ("WelcomeKit API complete: {0} matching jobs." -f $results.Count)
    Complete-SourceStats $stats
    return $results.ToArray()
}

function Get-WttjPublicFallbackJobs {
    if ($DisableWttjPublicFallback) {
        return @()
    }

    Set-RunWindowTitle "Analytics Job Crawler - WTTJ"
    Write-RunStatus "Collecting Welcome to the Jungle jobs from public sitemaps..."
    $stats = Start-SourceStats "WTTJ public"
    $results = New-Object System.Collections.Generic.List[object]
    $candidateSeen = @{}
    $candidates = New-Object System.Collections.Generic.List[object]

    try {
        Add-SourceMetric -Stats $stats -Name "SearchRequests"
        $indexUrl = [string](Get-ConfigPathValue -Object $script:JobCrawlerSourcesConfig -Path "endpoints.wttj_sitemap_index" -DefaultValue "https://www.welcometothejungle.com/sitemaps/index.xml.gz")
        $indexXml = Invoke-CurlTextRequest $indexUrl
    }
    catch {
        Add-SourceMetric -Stats $stats -Name "Errors"
        Write-Warning ("Could not read WTTJ sitemap index: {0}" -f $_.Exception.Message)
        Complete-SourceStats $stats
        return @()
    }

    $sitemapMatches = [regex]::Matches($indexXml, "<loc>(?<url>https://www\.welcometothejungle\.com/sitemaps/job-listings\.\d+\.xml\.gz)</loc>", [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    Write-RunStatus ("WTTJ sitemap scan: {0} job sitemap(s) found." -f $sitemapMatches.Count)
    $sitemapCount = 0
    foreach ($sitemapMatch in $sitemapMatches) {
        $sitemapCount++
        Write-CountProgress -Activity "WTTJ sitemap scan" -Current $sitemapCount -Total $sitemapMatches.Count -Found $candidates.Count -Every 10
        $sitemapUrl = $sitemapMatch.Groups["url"].Value
        try {
            Add-SourceMetric -Stats $stats -Name "SearchRequests"
            $xml = Invoke-CurlTextRequest $sitemapUrl
        }
        catch {
            Add-SourceMetric -Stats $stats -Name "Errors"
            Write-Warning ("Could not read WTTJ sitemap {0}: {1}" -f $sitemapUrl, $_.Exception.Message)
            continue
        }

        $urlMatches = [regex]::Matches($xml, "(?is)<url>.*?</url>")
        foreach ($urlMatch in $urlMatches) {
            $block = $urlMatch.Value
            $locMatch = [regex]::Match($block, "<loc>(?<loc>.*?)</loc>", [Text.RegularExpressions.RegexOptions]::IgnoreCase)
            $lastmodMatch = [regex]::Match($block, "<lastmod>(?<lastmod>.*?)</lastmod>", [Text.RegularExpressions.RegexOptions]::IgnoreCase)
            if (-not $locMatch.Success -or -not $lastmodMatch.Success) {
                continue
            }

            $loc = ConvertFrom-HtmlAttribute $locMatch.Groups["loc"].Value
            if ($candidateSeen.ContainsKey($loc)) {
                continue
            }
            Add-SourceMetric -Stats $stats -Name "Candidates"

            $lastmod = ConvertTo-DateTimeOffsetOrNull $lastmodMatch.Groups["lastmod"].Value
            if (-not (Test-IsRecent $lastmod)) {
                Add-SourceMetric -Stats $stats -Name "SkippedOld"
                continue
            }

            if ($loc -notmatch $WttjUrlCandidatePattern) {
                Add-SourceMetric -Stats $stats -Name "SkippedNoMatch"
                continue
            }

            $slugTitle = Get-TitleFromWttjUrl $loc
            $urlText = "{0} {1}" -f $loc, (($slugTitle -replace "[-_]", " "))
            if (Test-ShouldSkipEarlyByContract -Text $urlText) {
                Add-SourceMetric -Stats $stats -Name "SkippedContract"
                continue
            }

            if ((Test-IsWttjFranceTarget) -and (Test-IsWttjKnownForeignLocation (Get-WttjLocationFromUrl $loc))) {
                Add-SourceMetric -Stats $stats -Name "SkippedNoMatch"
                continue
            }

            $candidateSeen[$loc] = $true
            $candidates.Add([PSCustomObject]@{
                Url       = $loc
                LastMod   = $lastmod
                SlugTitle = $slugTitle
                Score     = Get-WttjCandidateScore $loc
            }) | Out-Null
        }
    }

    $selectedCandidates = $candidates |
        Sort-Object -Property @{ Expression = "Score"; Descending = $true }, @{ Expression = "LastMod"; Descending = $true } |
        Select-Object -First $MaxWttjCandidatePages

    $selectedCandidateCount = @($selectedCandidates).Count
    Add-SourceMetric -Stats $stats -Name "SelectedDetails" -Amount $selectedCandidateCount
    Add-SourceMetric -Stats $stats -Name "SkippedByCap" -Amount ([Math]::Max(0, $candidates.Count - $selectedCandidateCount))
    Write-RunStatus ("WTTJ candidates selected: {0} page(s) to inspect." -f $selectedCandidateCount)
    $count = 0
    foreach ($candidate in $selectedCandidates) {
        $count++
        Write-CountProgress -Activity "WTTJ candidate pages" -Current $count -Total $selectedCandidateCount -Found $results.Count -Every 10

        $urlText = "{0} {1}" -f $candidate.Url, (($candidate.SlugTitle -replace "[-_]", " "))
        $urlMatchResult = Get-JobMatch -Title $candidate.SlugTitle -Text $urlText
        $urlOnlyMatch = $urlMatchResult.IsMatch

        try {
            $html = Get-CachedText -Scope "wttj-detail" -Key $candidate.Url
            if ($null -ne $html) {
                Add-SourceMetric -Stats $stats -Name "CacheHits"
            }
            else {
                Add-SourceMetric -Stats $stats -Name "DetailRequests"
                $html = Invoke-CurlTextRequest $candidate.Url
                Set-CachedText -Scope "wttj-detail" -Key $candidate.Url -Text $html
            }
        }
        catch {
            Add-SourceMetric -Stats $stats -Name "Errors"
            if ($urlOnlyMatch) {
                $jobLocation = Get-WttjLocation -Html "" -Url $candidate.Url -Title $candidate.SlugTitle
                if (-not (Test-IsWttjLocationAllowed -JobLocation $jobLocation -Url $candidate.Url -Text $urlText)) {
                    Add-SourceMetric -Stats $stats -Name "SkippedNoMatch"
                    continue
                }

                $fallbackResult = New-JobResult -Title $candidate.SlugTitle -CompanyName (Get-WttjCompanyNameFromUrl $candidate.Url) -JobLocation $jobLocation -ContractType (Get-ContractTypeFromText -Text $urlText) -MatchScore $urlMatchResult.Score -MatchLevel $urlMatchResult.Level -MatchedKeywords $urlMatchResult.Keywords -Url $candidate.Url -Platform "Welcome to the Jungle" -PublishedAt $candidate.LastMod -SourceText $urlText
                if ($null -ne $fallbackResult) {
                    $results.Add($fallbackResult) | Out-Null
                    Add-SourceMetric -Stats $stats -Name "Matches"
                }
            }
            continue
        }

        if ($html -match "(?i)<title>ERROR: The request could not be satisfied</title>|AwsWafIntegration|challenge-container") {
            if ($urlOnlyMatch) {
                $jobLocation = Get-WttjLocation -Html $html -Url $candidate.Url -Title $candidate.SlugTitle
                if (-not (Test-IsWttjLocationAllowed -JobLocation $jobLocation -Url $candidate.Url -Text $urlText)) {
                    Add-SourceMetric -Stats $stats -Name "SkippedNoMatch"
                    continue
                }

                $fallbackResult = New-JobResult -Title $candidate.SlugTitle -CompanyName (Get-WttjCompanyNameFromUrl $candidate.Url) -JobLocation $jobLocation -ContractType (Get-ContractTypeFromText -Text $urlText) -MatchScore $urlMatchResult.Score -MatchLevel $urlMatchResult.Level -MatchedKeywords $urlMatchResult.Keywords -Url $candidate.Url -Platform "Welcome to the Jungle" -PublishedAt $candidate.LastMod -SourceText $urlText
                if ($null -ne $fallbackResult) {
                    $results.Add($fallbackResult) | Out-Null
                    Add-SourceMetric -Stats $stats -Name "Matches"
                }
            }
            continue
        }

        $title = Get-TitleFromHtml $html
        if ([string]::IsNullOrWhiteSpace($title)) {
            $title = $candidate.SlugTitle
        }

        $publishedAt = $candidate.LastMod
        $publishedMatch = [regex]::Match($html, '(?i)"published_at"\s*:\s*"(?<date>[^"]+)"')
        if ($publishedMatch.Success) {
            $parsedPublished = ConvertTo-DateTimeOffsetOrNull $publishedMatch.Groups["date"].Value
            if ($null -ne $parsedPublished) {
                $publishedAt = $parsedPublished
            }
        }

        if (-not (Test-IsRecent $publishedAt)) {
            Add-SourceMetric -Stats $stats -Name "SkippedOld"
            continue
        }

        $pageText = ConvertFrom-HtmlText $html
        $combined = "{0} {1} {2}" -f $title, $candidate.Url, $pageText
        $match = Get-JobMatch -Title $title -Text $combined
        if (-not $match.IsMatch -and -not $urlOnlyMatch) {
            Add-SourceMetric -Stats $stats -Name "SkippedNoMatch"
            continue
        }
        if (-not $match.IsMatch) {
            $match = $urlMatchResult
        }

        $jobLocation = Get-WttjLocation -Html $html -Url $candidate.Url -Title $title
        if (-not (Test-IsWttjLocationAllowed -JobLocation $jobLocation -Url $candidate.Url -Text $combined)) {
            Add-SourceMetric -Stats $stats -Name "SkippedNoMatch"
            continue
        }

        $result = New-JobResult -Title $title -CompanyName (Get-WttjCompanyNameFromUrl $candidate.Url) -JobLocation $jobLocation -ContractType (Get-ContractTypeFromText -Text $combined) -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url $candidate.Url -Platform "Welcome to the Jungle" -PublishedAt $publishedAt -SourceText $combined
        if ($null -ne $result) {
            $results.Add($result) | Out-Null
            Add-SourceMetric -Stats $stats -Name "Matches"
        }

        Start-Sleep -Milliseconds 250
    }

    Write-RunStatus ("WTTJ public fallback complete: {0} matching jobs." -f $results.Count)
    Complete-SourceStats $stats
    return $results.ToArray()
}

