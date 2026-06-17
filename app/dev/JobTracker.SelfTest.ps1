# Auto-extracted from Find-AnalyticsJobs.ps1. Keep dot-sourced execution order in the main script.

function Assert-ScoringCondition {
    param(
        [bool]$Condition,
        [string]$Message
    )

    if (-not $Condition) {
        throw "Scoring self-test failed: $Message"
    }
}

function Invoke-ScoringSelfTest {
    $script:JobCrawlerPreferences = Get-JobCrawlerPreferences
    $script:SeenResultKeys = @{}
    $script:FeedbackLearningProfile = $null

    $mojibakeInterim = "Int" + [string][char]0x00C3 + [string][char]0x00A9 + "rim - 6 Mois"
    $expectedInterim = "Int" + [string][char]0x00E9 + "rim - 6 Mois"
    Assert-ScoringCondition -Condition ((Repair-DisplayText $mojibakeInterim) -eq $expectedInterim) -Message "Expected UTF-8 mojibake to be repaired with French accents."
    $terminalMojibakeInterim = "Int" + [string][char]0x251C + [string][char]0x00AE + "rim - 6 Mois"
    Assert-ScoringCondition -Condition ((Repair-DisplayText $terminalMojibakeInterim) -eq $expectedInterim) -Message "Expected terminal mojibake to be repaired with French accents."
    $oemMojibakeLocation = "CDI " + [string][char]0x251C + [string][char]0x00E1 + " Paris"
    $expectedLocation = "CDI " + [string][char]0x00E0 + " Paris"
    Assert-ScoringCondition -Condition ((Repair-DisplayText $oemMojibakeLocation) -eq $expectedLocation) -Message "Expected OEM mojibake to be repaired with French accents."
    Assert-ScoringCondition -Condition (Test-IsExcludedContractType "Freelance") -Message "Expected freelance contracts to be excluded."

    $annonceurMatch = Get-JobMatch -Title "Web Analyst CRO" -Text "Google Tag Manager GA4 ContentSquare dataLayer tagging plan"
    Assert-ScoringCondition -Condition $annonceurMatch.IsMatch -Message "Expected a Web Analyst CRO role with web analytics tools to match."
    $expandedToolMatch = Get-JobMatch -Title "Tracking Specialist" -Text "Tag Commander Commanders Act Tealium server-side tracking RGPD"
    Assert-ScoringCondition -Condition $expandedToolMatch.IsMatch -Message "Expected Tag Commander, Commanders Act, Tealium, server-side, and RGPD signals to match."
    Assert-ScoringCondition -Condition ($expandedToolMatch.Keywords -match "Tag Commander" -and $expandedToolMatch.Keywords -match "Tealium" -and $expandedToolMatch.Keywords -match "server-side" -and $expandedToolMatch.Keywords -match "RGPD") -Message "Expected expanded tool/mission keywords to be reported."
    $positiveFeedbackRow = New-OrderedJobRecord @{
        status           = "interesting"
        job_title        = "Web Analyst"
        matched_keywords = "Tealium; server-side tracking"
    }
    $ignoredFeedbackRow = New-OrderedJobRecord @{
        status    = "ignored"
        job_title = "SEO Manager"
        notes     = "ignore_reason=too_seo_sea_marketing; detail=too marketing"
    }
    $script:FeedbackLearningProfile = New-FeedbackLearningProfile -Rows @($positiveFeedbackRow, $ignoredFeedbackRow)
    $positiveLearning = Get-FeedbackLearningAdjustment -FullText "tealium server side tracking" -HasCoreTitleSignal:$true -HasWebAnalyticsToolSignal:$true -HasDigitalAnalyticsContext:$true
    Assert-ScoringCondition -Condition ([int]$positiveLearning.Adjustment -gt 0 -and (($positiveLearning.Reasons -join ";") -match "Tealium")) -Message "Expected positive saved tracker feedback to boost similar tool signals."
    $negativeLearning = Get-FeedbackLearningAdjustment -FullText "seo sea paid media campaign" -HasCoreTitleSignal:$false -HasWebAnalyticsToolSignal:$false -HasDigitalAnalyticsContext:$false
    Assert-ScoringCondition -Condition ([int]$negativeLearning.Adjustment -lt 0 -and (($negativeLearning.Reasons -join ";") -match "SEO/SEA")) -Message "Expected ignored saved tracker feedback to penalize similar marketing-only signals."
    $script:FeedbackLearningProfile = $null
    $annonceurResult = New-JobResult `
        -Title "Web Analyst CRO" `
        -CompanyName "Radio France" `
        -JobLocation "Paris" `
        -ContractType "CDI" `
        -MatchScore $annonceurMatch.Score `
        -MatchLevel $annonceurMatch.Level `
        -MatchedKeywords $annonceurMatch.Keywords `
        -Url "https://example.test/jobs/radio-france-web-analyst" `
        -Platform "Test" `
        -PublishedAt ([DateTimeOffset]::Now) `
        -SourceText "Google Tag Manager GA4 ContentSquare dataLayer tagging plan"
    Assert-ScoringCondition -Condition ((Get-RowValue -Row $annonceurResult -Name "employer_type") -eq "annonceur") -Message "Expected Radio France to be classified as annonceur."
    Assert-ScoringCondition -Condition ((Get-IntegerRowValue -Row $annonceurResult -Name "match_score") -gt (Get-IntegerRowValue -Row $annonceurResult -Name "role_score")) -Message "Expected annonceur/Paris/CDI fit to boost the role score."

    $consultingMatch = Get-JobMatch -Title "Digital Analytics Consultant" -Text "GA4 Google Tag Manager Piano Analytics ContentSquare"
    $consultingResult = New-JobResult `
        -Title "Digital Analytics Consultant" `
        -CompanyName "fifty-five" `
        -JobLocation "Paris" `
        -ContractType "CDI" `
        -MatchScore $consultingMatch.Score `
        -MatchLevel $consultingMatch.Level `
        -MatchedKeywords $consultingMatch.Keywords `
        -Url "https://example.test/jobs/fifty-five-digital-analytics-consultant" `
        -Platform "Test" `
        -PublishedAt ([DateTimeOffset]::Now) `
        -SourceText "GA4 Google Tag Manager Piano Analytics ContentSquare"
    Assert-ScoringCondition -Condition ((Get-RowValue -Row $consultingResult -Name "employer_type") -eq "consulting") -Message "Expected fifty-five to be classified as consulting."
    Assert-ScoringCondition -Condition ((Get-IntegerRowValue -Row $consultingResult -Name "employer_fit") -lt 0) -Message "Expected consulting employer type to be demoted, not excluded."

    $dataEngineeringMatch = Get-JobMatch -Title "Data Analyst" -Text "python dbt snowflake airflow data warehouse data pipeline"
    Assert-ScoringCondition -Condition (-not $dataEngineeringMatch.IsMatch) -Message "Expected warehouse/python data analyst role without web analytics signals to stay below the match threshold."
    $companyNameOnlyToolMatch = Get-JobMatch -Title "People Business Partner" -Text "Contentsquare Paris Full-time"
    Assert-ScoringCondition -Condition (-not $companyNameOnlyToolMatch.IsMatch) -Message "Expected a non-analytics role not to match only because the company name is an analytics tool."

    $titleOnlyExcludedContract = New-JobResult `
        -Title "Alternance Assistant web analytics" `
        -CompanyName "Example Company" `
        -JobLocation "Paris" `
        -ContractType "" `
        -MatchScore $annonceurMatch.Score `
        -MatchLevel $annonceurMatch.Level `
        -MatchedKeywords $annonceurMatch.Keywords `
        -Url "https://example.test/jobs/alternance-web-analytics" `
        -Platform "Test" `
        -PublishedAt ([DateTimeOffset]::Now) `
        -SourceText "Google Analytics tagging plan"
    Assert-ScoringCondition -Condition ((Get-RowValue -Row $titleOnlyExcludedContract -Name "contract_type") -eq "Apprenticeship") -Message "Expected title-only alternance to be mapped to Apprenticeship."
    Assert-ScoringCondition -Condition (Test-IsExcludedContractType (Get-RowValue -Row $titleOnlyExcludedContract -Name "contract_type")) -Message "Expected title-only alternance to be excluded by contract filtering."
    $titleOverridesGenericContract = New-JobResult `
        -Title "STAGE - Communication digitale et web analytics" `
        -CompanyName "Example Company" `
        -JobLocation "Paris" `
        -ContractType "Full-time" `
        -MatchScore $annonceurMatch.Score `
        -MatchLevel $annonceurMatch.Level `
        -MatchedKeywords $annonceurMatch.Keywords `
        -Url "https://example.test/jobs/stage-web-analytics" `
        -Platform "Test" `
        -PublishedAt ([DateTimeOffset]::Now) `
        -SourceText "Google Analytics tagging plan"
    Assert-ScoringCondition -Condition ((Get-RowValue -Row $titleOverridesGenericContract -Name "contract_type") -eq "Internship") -Message "Expected explicit STAGE title to override generic Full-time contract."

    $junkLocation = Get-WttjLocationFromUrl "https://www.welcometothejungle.com/fr/companies/acme/jobs/web-analyst_5Kvvowa"
    Assert-ScoringCondition -Condition ([string]::IsNullOrWhiteSpace($junkLocation)) -Message "Expected random WTTJ URL suffixes not to become city names."
    $parisLocation = Get-WttjLocationFromUrl "https://www.welcometothejungle.com/fr/companies/acme/jobs/web-analyst_paris"
    Assert-ScoringCondition -Condition ($parisLocation -eq "Paris") -Message "Expected readable WTTJ city suffix to be kept."
    $multiPartUrlLocation = Get-WttjLocationFromUrl "https://www.welcometothejungle.com/en/companies/acme/jobs/web-analyst_london_ACME_3zgazPX"
    Assert-ScoringCondition -Condition ($multiPartUrlLocation -eq "London") -Message "Expected WTTJ URL city suffixes before reference tokens to be parsed."
    $wttjInitialDataHtml = 'window.__INITIAL_DATA__ = "{\"queries\":[{\"state\":{\"data\":{\"offices\":[{\"city\":\"Saint-Denis\",\"country_code\":\"FR\"}]}}}]}";'
    $wttjInitialDataLocation = Get-WttjLocation -Html $wttjInitialDataHtml -Url "https://www.welcometothejungle.com/fr/companies/acme/jobs/web-analyst_5Kvvowa" -Title "Web Analyst"
    Assert-ScoringCondition -Condition ($wttjInitialDataLocation -eq "Saint-Denis, France") -Message "Expected WTTJ embedded office city and country code to be parsed."
    Assert-ScoringCondition -Condition (-not (Test-IsWttjLocationAllowed -JobLocation "New York, United States" -Url "https://www.welcometothejungle.com/fr/companies/acme/jobs/web-analyst_new-york" -Text "Web Analyst")) -Message "Expected foreign WTTJ locations to be rejected for France crawls."
    Assert-ScoringCondition -Condition (-not (Test-IsWttjLocationAllowed -JobLocation "" -Url "https://www.welcometothejungle.com/fr/companies/acme/jobs/web-analyst_5Kvvowa" -Text "Web Analyst")) -Message "Expected blank WTTJ locations to be rejected for France crawls."
    $invalidExistingWttjRow = New-OrderedJobRecord @{
        status         = "ignored"
        job_title      = "Web Analyst"
        company_name   = "Example Company"
        location       = ""
        contract_type  = "CDI"
        match_score    = "80"
        match_level    = "Good"
        job_url_raw    = "https://www.welcometothejungle.com/en/companies/acme/jobs/web-analyst_london_ACME_3zgazPX"
        platform       = "Welcome to the Jungle"
        published_date = ([DateTimeOffset]::Now.ToString("yyyy-MM-dd"))
    }
    $foreignExistingWttjRow = New-OrderedJobRecord @{
        status         = "ignored"
        job_title      = "Web Analyst Casablanca"
        company_name   = "Example Company"
        location       = "Casablanca"
        contract_type  = "CDI"
        match_score    = "80"
        match_level    = "Good"
        job_url_raw    = "https://www.welcometothejungle.com/fr/companies/acme/jobs/web-analyst_casablanca"
        platform       = "Welcome to the Jungle"
        published_date = ([DateTimeOffset]::Now.ToString("yyyy-MM-dd"))
    }
    $invalidMerge = Merge-JobsWithTracker -CurrentRows @() -ExistingRows @($invalidExistingWttjRow, $foreignExistingWttjRow) -Path "selftest.xlsx" -SkipBackup
    Assert-ScoringCondition -Condition ([int]$invalidMerge.RemovedCount -eq 2 -and @($invalidMerge.TrackerRows).Count -eq 0) -Message "Expected invalid existing non-applied WTTJ rows to be removed during merge."

    $franceTravailMock = [PSCustomObject]@{
        id                  = "123ABC"
        intitule            = "Web Analyst"
        description         = "Google Analytics GA4 Google Tag Manager dataLayer"
        dateCreation        = ([DateTimeOffset]::Now.ToString("o"))
        typeContrat         = "CDI"
        typeContratLibelle  = "CDI"
        urlPostulation      = "https://candidat.francetravail.fr/offres/recherche/detail/123ABC"
        lieuTravail         = [PSCustomObject]@{ libelle = "75 - Paris" }
        entreprise          = [PSCustomObject]@{ nom = "Example Annonceur" }
    }
    Assert-ScoringCondition -Condition ((Get-FranceTravailContractType $franceTravailMock) -eq "CDI") -Message "Expected France Travail CDI contract mapping."
    Assert-ScoringCondition -Condition ((Get-FranceTravailCompanyName $franceTravailMock) -eq "Example Annonceur") -Message "Expected France Travail company mapping."
    Assert-ScoringCondition -Condition ((Get-FranceTravailLocation $franceTravailMock) -eq "75 - Paris") -Message "Expected France Travail location mapping."

    $adzunaMock = [PSCustomObject]@{
        title         = "Digital Analyst"
        description   = "Piano Analytics ContentSquare Google Analytics"
        created       = ([DateTimeOffset]::Now.ToString("o"))
        redirect_url  = "https://www.adzuna.fr/details/123"
        contract_type = "permanent"
        contract_time = "full_time"
        company       = [PSCustomObject]@{ display_name = "Example Retailer" }
        location      = [PSCustomObject]@{ display_name = "Paris, Ile-de-France" }
    }
    Assert-ScoringCondition -Condition ((Get-AdzunaContractType $adzunaMock) -eq "Permanent") -Message "Expected Adzuna permanent contract mapping."
    Assert-ScoringCondition -Condition ((Get-AdzunaCompanyName $adzunaMock) -eq "Example Retailer") -Message "Expected Adzuna company mapping."
    Assert-ScoringCondition -Condition ((Get-AdzunaLocation $adzunaMock) -eq "Paris, Ile-de-France") -Message "Expected Adzuna location mapping."

    $apecMock = [PSCustomObject]@{
        id              = 123456789
        numeroOffre     = "123456789W"
        intitule        = "Web Analyst F/H"
        nomCommercial   = "Example Retailer"
        lieuTexte       = "Paris - 75"
        typeContrat     = 101888
        texteOffre      = "Google Analytics GA4 Google Tag Manager ContentSquare"
        datePublication = ([DateTimeOffset]::Now.ToString("yyyy-MM-ddTHH:mm:ss.000+0000"))
    }
    Assert-ScoringCondition -Condition ((Get-ApecContractType $apecMock) -eq "CDI") -Message "Expected APEC CDI contract mapping."
    Assert-ScoringCondition -Condition ((Get-ApecJobUrl $apecMock) -match "/detail-offre/123456789W$") -Message "Expected APEC detail URL mapping."

    $helloWorkMockHtml = @'
<script type="application/ld+json">{"@context":"https://schema.org","@type":"JobPosting","title":"Web Analyst H/F","description":"Google Tag Manager GA4 dataLayer","datePosted":"2026-06-16T09:38:15Z","employmentType":"FULL_TIME","hiringOrganization":{"@type":"Organization","name":"Example Retailer"},"jobLocation":{"@type":"Place","address":{"@type":"PostalAddress","addressLocality":"Paris","addressRegion":"Ile-de-France","addressCountry":"FR"}}}</script>
<script type="application/ld+json">{"JobTitle":"Web Analyst H/F","Company":"Example Retailer","Localisation":"Paris - 75","ContractType":"CDI","Description":"Piano Analytics ContentSquare"}</script>
'@
    $helloWorkMetadata = Get-HelloWorkJobMetadata -Html $helloWorkMockHtml
    Assert-ScoringCondition -Condition ($helloWorkMetadata.Title -eq "Web Analyst H/F") -Message "Expected HelloWork title metadata mapping."
    Assert-ScoringCondition -Condition ($helloWorkMetadata.Company -eq "Example Retailer") -Message "Expected HelloWork company metadata mapping."
    Assert-ScoringCondition -Condition ($helloWorkMetadata.Location -eq "Paris - 75") -Message "Expected HelloWork custom location metadata to be preferred."
    Assert-ScoringCondition -Condition ($helloWorkMetadata.Contract -eq "CDI") -Message "Expected HelloWork contract metadata mapping."
    Assert-ScoringCondition -Condition ($helloWorkMetadata.Description -match "Piano Analytics") -Message "Expected HelloWork custom description metadata mapping."
    Assert-ScoringCondition -Condition (Test-IsRecent (ConvertFrom-FrenchRelativeDateText "il y a 2 jours")) -Message "Expected French relative dates to parse as recent."

    $crossPlatformMatch = Get-JobMatch -Title "Web Analyst" -Text "GA4 Google Tag Manager ContentSquare dataLayer"
    $crossPlatformRows = @(
        (New-JobResult -Title "Web Analyst" -CompanyName "Radio France" -JobLocation "Paris" -ContractType "CDI" -MatchScore $crossPlatformMatch.Score -MatchLevel $crossPlatformMatch.Level -MatchedKeywords $crossPlatformMatch.Keywords -Url "https://www.linkedin.com/jobs/view/111" -Platform "LinkedIn" -PublishedAt ([DateTimeOffset]::Now) -SourceText "GA4 Google Tag Manager ContentSquare dataLayer"),
        (New-JobResult -Title "Analyste Web H/F" -CompanyName "Radio France" -JobLocation "75 - Paris" -ContractType "CDI" -MatchScore $crossPlatformMatch.Score -MatchLevel $crossPlatformMatch.Level -MatchedKeywords $crossPlatformMatch.Keywords -Url "https://candidat.francetravail.fr/offres/recherche/detail/111" -Platform "France Travail" -PublishedAt ([DateTimeOffset]::Now) -SourceText "GA4 Google Tag Manager ContentSquare dataLayer"),
        (New-JobResult -Title "Web Analyst" -CompanyName "Radio France" -JobLocation "Paris, Ile-de-France" -ContractType "Permanent" -MatchScore $crossPlatformMatch.Score -MatchLevel $crossPlatformMatch.Level -MatchedKeywords $crossPlatformMatch.Keywords -Url "https://www.adzuna.fr/details/111" -Platform "Adzuna" -PublishedAt ([DateTimeOffset]::Now) -SourceText "GA4 Google Tag Manager ContentSquare dataLayer"),
        (New-JobResult -Title "Web Analyst F/H" -CompanyName "Radio France" -JobLocation "Paris - 75" -ContractType "CDI" -MatchScore $crossPlatformMatch.Score -MatchLevel $crossPlatformMatch.Level -MatchedKeywords $crossPlatformMatch.Keywords -Url "https://www.apec.fr/candidat/recherche-emploi.html/emploi/detail-offre/111W" -Platform "APEC" -PublishedAt ([DateTimeOffset]::Now) -SourceText "GA4 Google Tag Manager ContentSquare dataLayer"),
        (New-JobResult -Title "Web Analyst H/F" -CompanyName "Radio France" -JobLocation "Paris - 75" -ContractType "CDI" -MatchScore $crossPlatformMatch.Score -MatchLevel $crossPlatformMatch.Level -MatchedKeywords $crossPlatformMatch.Keywords -Url "https://www.hellowork.com/fr-fr/emplois/111.html" -Platform "HelloWork" -PublishedAt ([DateTimeOffset]::Now) -SourceText "GA4 Google Tag Manager ContentSquare dataLayer")
    )
    $crossPlatformKeys = @($crossPlatformRows | ForEach-Object { Get-JobDedupeKeyFromRow $_ } | Select-Object -Unique)
    Assert-ScoringCondition -Condition ($crossPlatformKeys.Count -eq 1) -Message "Expected same company/title role from several platforms to share one dedupe key."
    $crossPlatformMerged = Merge-SimilarJobRows -Rows $crossPlatformRows -Reason "test cross-platform duplicate"
    $crossPlatformSources = Get-RowValue -Row $crossPlatformMerged -Name "platform"
    Assert-ScoringCondition -Condition ($crossPlatformSources -match "LinkedIn" -and $crossPlatformSources -match "France Travail" -and $crossPlatformSources -match "Adzuna" -and $crossPlatformSources -match "APEC" -and $crossPlatformSources -match "HelloWork") -Message "Expected merged cross-platform row to keep all source names."
    Assert-ScoringCondition -Condition ((Get-RowValue -Row $crossPlatformMerged -Name "source_count") -eq "5") -Message "Expected source_count to count unique platforms."
    Assert-ScoringCondition -Condition ((Get-RowValue -Row $crossPlatformMerged -Name "job_url_raw") -match "apec") -Message "Expected APEC URL to be preferred over LinkedIn, France Travail, HelloWork, and Adzuna for this merge."
    Assert-ScoringCondition -Condition ((Get-RowValue -Row $crossPlatformMerged -Name "alternate_urls") -match "linkedin" -and (Get-RowValue -Row $crossPlatformMerged -Name "alternate_urls") -match "adzuna" -and (Get-RowValue -Row $crossPlatformMerged -Name "alternate_urls") -match "hellowork" -and (Get-RowValue -Row $crossPlatformMerged -Name "alternate_urls") -match "francetravail") -Message "Expected alternate URLs to keep non-primary cross-platform links."

    Write-Host "Scoring self-test passed."
}

