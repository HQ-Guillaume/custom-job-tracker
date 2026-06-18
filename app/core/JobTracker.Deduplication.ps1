function ConvertTo-IdentityText {
    param(
        [AllowNull()][string]$Text,
        [switch]$Title
    )

    $clean = ConvertTo-MatchText $Text
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return ""
    }

    if ($Title) {
        $clean = $clean -replace "\b(cdi|cdd|stage|internship|intern|alternance|apprentissage|apprenticeship|contrat|full[- ]?time|permanent|h|f|hf|h/f|f/h|m|x|nb)\b", " "
        $clean = $clean -replace "\b(senior|junior|jr|sr|lead|manager|confirme|confirmee|experimente|experimentee)\b", " "
        $clean = $clean -replace "\b(paris|lyon|lille|bordeaux|nantes|rennes|montpellier|marseille|toulouse|puteaux|levallois|boulogne|casablanca|france)\b", " "
    }
    else {
        $clean = $clean -replace "\b(france|fr|group|groupe|sa|sas|sasu|ltd|limited|inc|plc)\b", " "
    }

    $clean = $clean -replace "[^a-z0-9]+", " "
    return ([regex]::Replace($clean, "\s+", " ")).Trim()
}

function Split-NormalizedTokens {
    param([AllowNull()][string]$Text)

    $clean = ConvertTo-IdentityText -Text $Text
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return @()
    }

    return @($clean -split "\s+" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}

function Test-IsGenericJobBoardName {
    param([AllowNull()][string]$Name)

    $text = ConvertTo-MatchText $Name
    $text = $text -replace "[^a-z0-9]+", " "
    $text = ([regex]::Replace($text, "\s+", " ")).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $false
    }

    return $text -match "\b(france travail|pole emploi|poles emploi|adzuna|linkedin|indeed|hellowork|hello work|meteojob|jobijoba|monster|apec|jobgether|talent com|confidential|confidentiel|licorne|recrutement)\b"
}

function Get-DedupeCompanyNoiseTokens {
    return @(
        "the", "and", "et", "de", "du", "des", "la", "le", "les",
        "sa", "sas", "sasu", "ltd", "limited", "inc", "plc", "fr", "france",
        "group", "groupe", "company", "companies", "media", "digital",
        "consulting", "consultants", "technology", "technologies", "solutions"
    )
}

function Get-DedupeWeakCompanyTokens {
    return @("confidential", "confidentiel", "jobgether", "licorne", "recrutement", "talent", "emploi", "travail", "adzuna", "linkedin", "indeed", "hellowork", "meteojob", "jobijoba", "monster", "apec")
}

function Get-DedupeCompanyDescriptorTokens {
    return @(
        "assurance", "assurances", "insurance", "mutuelle", "bank", "banque",
        "consulting", "consultants", "consultant", "conseil", "agency", "agence",
        "esn", "ssii", "recrutement", "media", "digital", "technology", "technologies",
        "tech", "solutions", "services", "service", "group", "groupe", "company", "companies",
        "france", "international", "global", "partners", "partner"
    )
}

function Get-DedupeCompanyStrongTokens {
    param([AllowNull()][string]$CompanyName)

    if (Test-IsGenericJobBoardName $CompanyName) {
        return @()
    }

    $tokens = @(Split-NormalizedTokens $CompanyName)
    if ($tokens.Count -eq 0) {
        return @()
    }

    $noise = @(Get-DedupeCompanyNoiseTokens)
    $weakCompanyTokens = @(Get-DedupeWeakCompanyTokens)
    return @($tokens | Where-Object { $_.Length -gt 1 -and $noise -notcontains $_ -and $weakCompanyTokens -notcontains $_ })
}

function Get-DedupeCompanyKey {
    param([AllowNull()][string]$CompanyName)

    $strongTokens = @(Get-DedupeCompanyStrongTokens $CompanyName)
    if ($strongTokens.Count -eq 0) {
        return ""
    }

    if ($strongTokens.Count -eq 1) {
        return $strongTokens[0]
    }

    return ($strongTokens | Select-Object -First 2) -join " "
}

function Get-ConfiguredCompanyAliasCanonicalKey {
    param([AllowNull()][string]$CompanyName)

    if ([string]::IsNullOrWhiteSpace($CompanyName) -or -not (Get-Variable -Name JobCrawlerMatchingRules -Scope Script -ErrorAction SilentlyContinue)) {
        return ""
    }

    $companyText = ConvertTo-IdentityText -Text $CompanyName
    if ([string]::IsNullOrWhiteSpace($companyText)) {
        return ""
    }

    foreach ($aliasGroup in @(Get-ConfigPathValue -Object $script:JobCrawlerMatchingRules -Path "deduplication.company_aliases" -DefaultValue @())) {
        $canonical = [string](Get-ConfigProperty -Object $aliasGroup -Name "canonical" -DefaultValue "")
        $aliases = @(Get-ConfigStringArray (Get-ConfigProperty -Object $aliasGroup -Name "aliases" -DefaultValue @()))
        $candidateValues = @($canonical) + $aliases
        foreach ($candidate in @($candidateValues)) {
            $candidateText = ConvertTo-IdentityText -Text $candidate
            if (-not [string]::IsNullOrWhiteSpace($candidateText) -and $companyText -eq $candidateText) {
                $canonicalKey = Get-DedupeCompanyKey $canonical
                if (-not [string]::IsNullOrWhiteSpace($canonicalKey)) {
                    return $canonicalKey
                }
            }
        }
    }

    return ""
}

function Get-DedupeCompanyAliasKeys {
    param([AllowNull()][string]$CompanyName)

    $keys = New-Object System.Collections.Generic.List[string]
    $configuredCanonical = Get-ConfiguredCompanyAliasCanonicalKey $CompanyName
    if (-not [string]::IsNullOrWhiteSpace($configuredCanonical)) {
        $keys.Add($configuredCanonical) | Out-Null
    }

    $primaryKey = Get-DedupeCompanyKey $CompanyName
    if (-not [string]::IsNullOrWhiteSpace($primaryKey)) {
        $keys.Add($primaryKey) | Out-Null
    }

    $strongTokens = @(Get-DedupeCompanyStrongTokens $CompanyName)
    $descriptorTokens = @(Get-DedupeCompanyDescriptorTokens)
    if ($strongTokens.Count -gt 1) {
        $rootToken = [string]$strongTokens[0]
        $remainingTokens = @($strongTokens | Select-Object -Skip 1)
        $remainingAreDescriptors = @($remainingTokens | Where-Object { $descriptorTokens -notcontains $_ }).Count -eq 0
        if ($rootToken.Length -ge 5 -and $remainingAreDescriptors) {
            $keys.Add($rootToken) | Out-Null
        }
    }

    return @($keys.ToArray() | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
}

function ConvertTo-DedupeTitleToken {
    param([AllowNull()][string]$Token)

    if ([string]::IsNullOrWhiteSpace($Token)) {
        return ""
    }

    switch -Regex ($Token) {
        "^(analyste|analystes)$" { return "analyst" }
        "^(consultante|consultants|consultantes)$" { return "consultant" }
        "^(digitale|digitaux|digitales)$" { return "digital" }
        "^(analytics|analytic|analytique|analytiques)$" { return "analytics" }
        "^(chargee|charges|chargees)$" { return "charge" }
        "^(performances)$" { return "performance" }
        default { return $Token }
    }
}

function Get-DedupeLocationKey {
    param([AllowNull()][string]$Location)

    $text = ConvertTo-MatchText $Location
    $text = $text -replace "[^a-z0-9]+", " "
    $text = ([regex]::Replace($text, "\s+", " ")).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return "unknown"
    }

    if ($text -match "\b(ile de france|greater paris|ville de paris|paris|puteaux|courbevoie|nanterre|levallois|boulogne|clichy|issy|issy les moulineaux|neuilly|la defense|saint denis|montreuil|gentilly|ivry|bagnolet|suresnes|rueil|meudon|chatillon|montrouge|vincennes)\b") {
        return "paris_metro"
    }

    if ($text -match "\b(remote|teletravail|france|fr)\b" -and $text -notmatch "\b(lyon|lille|bordeaux|nantes|rennes|montpellier|marseille|toulouse|nice|strasbourg|grenoble|dijon|angers|annecy|niort|caen|aix|limoges|poissy|armentieres|champagne|haute savoie|carros|casablanca)\b") {
        return "france"
    }

    $text = $text -replace "\b(et peripherie|metropolitan area|area|region|france|remote|hybrid|on site|sur site)\b", " "
    $tokens = @($text -split "\s+" | Where-Object { $_.Length -gt 1 })
    if ($tokens.Count -eq 0) {
        return "unknown"
    }

    return ($tokens | Select-Object -First 2) -join " "
}

function Get-DedupeTitleKey {
    param(
        [AllowNull()][string]$Title,
        [AllowNull()][string]$CompanyName
    )

    $text = ConvertTo-MatchText $Title
    $text = $text -replace "[^a-z0-9]+", " "
    $text = ([regex]::Replace($text, "\s+", " ")).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return ""
    }

    $companyTokens = @(Split-NormalizedTokens $CompanyName | Where-Object { $_.Length -gt 2 })
    foreach ($token in $companyTokens) {
        $text = $text -replace ("\b{0}\b" -f [regex]::Escape($token)), " "
    }

    $noise = @(
        "cdi", "cdd", "stage", "stagiaire", "internship", "intern", "alternance", "apprentissage",
        "contrat", "full", "time", "permanent", "temps", "plein", "h", "f", "hf", "fh", "fm", "mf", "x", "nb",
        "paris", "lyon", "lille", "bordeaux", "nantes", "rennes", "montpellier", "marseille", "toulouse",
        "puteaux", "courbevoie", "nanterre", "levallois", "boulogne", "clichy", "issy", "france",
        "ile", "de", "a", "au", "aux", "en", "la", "le", "les", "du", "des", "et", "emea",
        "media", "sa", "sas", "sasu", "groupe", "group"
    )
    $tokens = @(
        $text -split "\s+" |
            Where-Object { $_.Length -gt 1 -and $noise -notcontains $_ } |
            ForEach-Object { ConvertTo-DedupeTitleToken $_ } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
    )
    if ($tokens.Count -eq 0) {
        return ""
    }

    return ($tokens -join " ")
}

function Test-IsStrongDedupeTitle {
    param([AllowNull()][string]$TitleKey)

    if ([string]::IsNullOrWhiteSpace($TitleKey)) {
        return $false
    }

    $tokens = @($TitleKey -split "\s+" | Where-Object { $_.Length -gt 1 })
    if ($tokens.Count -lt 2) {
        return $false
    }

    return $TitleKey -match "\b(analyst|analyste|analytics|web|digital|tracking|tagging|taggage|data|traffic|performance|cro|conversion|consultant|manager|lead|product|gtm|seo|sea)\b"
}

function Test-UseLocationInDedupeKey {
    param([AllowNull()][string]$TitleKey)

    $tokens = @(([string]$TitleKey) -split "\s+" | Where-Object { $_.Length -gt 1 })
    if ($tokens.Count -ge 3) {
        return $false
    }

    $hasCoreAnalyticsTitle =
        (($TitleKey -match "\bweb\b") -and ($TitleKey -match "\banalyst\b|\banalytics\b")) -or
        (($TitleKey -match "\bdigital\b") -and ($TitleKey -match "\banalyst\b|\banalytics\b")) -or
        (($TitleKey -match "\bdata\b") -and ($TitleKey -match "\banalyst\b") -and ($TitleKey -match "\bweb\b|\bdigital\b|\banalytics\b")) -or
        ($TitleKey -match "\b(tracking|tagging|taggage|gtm|cro)\b")

    if ($hasCoreAnalyticsTitle) {
        return $false
    }

    return $true
}

function Get-CanonicalJobUrl {
    param([AllowNull()][string]$Url)

    if ([string]::IsNullOrWhiteSpace($Url)) {
        return ""
    }

    $clean = ([string]$Url).Trim().ToLowerInvariant()
    $clean = $clean -replace "#.*$", ""
    $clean = $clean -replace "\?.*$", ""
    $clean = $clean -replace "/+$", ""
    $clean = $clean -replace "^http://", "https://"
    $clean = $clean -replace "^https://(www\.)?", "https://"
    return $clean
}

function Get-SourceJobIdKeyFromUrl {
    param([AllowNull()][string]$Url)

    $canonicalUrl = Get-CanonicalJobUrl $Url
    if ([string]::IsNullOrWhiteSpace($canonicalUrl)) {
        return ""
    }

    switch -Regex ($canonicalUrl) {
        "linkedin\.com/jobs/view/(?:[^/]*-)?(?<id>\d+)$" { return "source-id|linkedin|{0}" -f $matches.Id }
        "linkedin\.com/jobs/view/(?<id>\d+)" { return "source-id|linkedin|{0}" -f $matches.Id }
        "welcometothejungle\.com/.+/jobs/(?<id>[^/?#]+)$" { return "source-id|wttj|{0}" -f $matches.Id }
        "apec\.fr/.+/detail-offre/(?<id>[^/?#]+)$" { return "source-id|apec|{0}" -f $matches.Id }
        "hellowork\.com/.+/emplois/(?<id>[^/?#]+)\.html$" { return "source-id|hellowork|{0}" -f $matches.Id }
        "francetravail\.fr/.+/detail/(?<id>[^/?#]+)$" { return "source-id|france-travail|{0}" -f $matches.Id }
        "pole-emploi\.fr/.+/detail/(?<id>[^/?#]+)$" { return "source-id|france-travail|{0}" -f $matches.Id }
        "adzuna\.fr/details/(?<id>\d+)$" { return "source-id|adzuna|{0}" -f $matches.Id }
    }

    return ""
}

function Get-DedupeTitleTokens {
    param([AllowNull()][string]$TitleKey)

    if ([string]::IsNullOrWhiteSpace($TitleKey)) {
        return @()
    }

    return @(([string]$TitleKey) -split "\s+" | Where-Object { $_.Length -gt 1 } | Sort-Object -Unique)
}

function Get-DedupeTokenSimilarity {
    param(
        [string[]]$Left,
        [string[]]$Right
    )

    $set = @{}
    foreach ($token in @($Left)) {
        if (-not [string]::IsNullOrWhiteSpace($token)) {
            $set[$token] = 1
        }
    }
    foreach ($token in @($Right)) {
        if ([string]::IsNullOrWhiteSpace($token)) {
            continue
        }
        if ($set.ContainsKey($token)) {
            $set[$token] = 3
        }
        else {
            $set[$token] = 2
        }
    }

    $union = @($set.Keys).Count
    if ($union -eq 0) {
        return 0
    }

    $intersection = @($set.Keys | Where-Object { $set[$_] -eq 3 }).Count
    return [Math]::Round($intersection / $union, 3)
}

function Get-DedupeTitleSimilarity {
    param(
        [AllowNull()][string]$LeftTitleKey,
        [AllowNull()][string]$RightTitleKey
    )

    return Get-DedupeTokenSimilarity -Left (Get-DedupeTitleTokens $LeftTitleKey) -Right (Get-DedupeTitleTokens $RightTitleKey)
}

function Get-DedupeRoleFamily {
    param([AllowNull()][string]$TitleKey)

    if ([string]::IsNullOrWhiteSpace($TitleKey)) {
        return ""
    }

    if ($TitleKey -match "\b(front|frontend|javascript|react|vue|angular|developpeur|developer)\b") {
        return "frontend_engineering"
    }
    if ($TitleKey -match "\b(tracking|tagging|taggage|gtm|tag)\b") {
        return "tracking"
    }
    if ($TitleKey -match "\bcro\b|conversion") {
        return "cro"
    }
    if ($TitleKey -match "\bproduct\b" -and $TitleKey -match "\b(analyst|analytics)\b") {
        return "product_analytics"
    }
    if ($TitleKey -match "\b(web|digital)\b" -and $TitleKey -match "\b(analyst|analytics)\b") {
        return "digital_analytics"
    }
    if ($TitleKey -match "\bdata\b" -and $TitleKey -match "\b(analyst|analytics)\b") {
        return "data_analytics"
    }
    if ($TitleKey -match "\bseo\b|\bsea\b|traffic|performance|marketing") {
        return "marketing_performance"
    }

    $tokens = @(Get-DedupeTitleTokens $TitleKey)
    if ($tokens.Count -gt 0) {
        return ($tokens | Select-Object -First ([Math]::Min(2, $tokens.Count))) -join "_"
    }

    return ""
}

function Test-DedupeLocationsCompatible {
    param(
        [AllowNull()][string]$LeftLocation,
        [AllowNull()][string]$RightLocation
    )

    $leftKey = Get-DedupeLocationKey $LeftLocation
    $rightKey = Get-DedupeLocationKey $RightLocation
    if ($leftKey -eq $rightKey) {
        return $true
    }
    if ($leftKey -in @("unknown", "france") -or $rightKey -in @("unknown", "france")) {
        return $true
    }
    if ($leftKey -eq "paris_metro" -or $rightKey -eq "paris_metro") {
        return $false
    }

    $leftTokens = @($leftKey -split "\s+" | Where-Object { $_.Length -gt 2 })
    $rightTokens = @($rightKey -split "\s+" | Where-Object { $_.Length -gt 2 })
    if ((Get-DedupeTokenSimilarity -Left $leftTokens -Right $rightTokens) -gt 0) {
        return $true
    }

    $leftText = ConvertTo-MatchText $LeftLocation
    $rightText = ConvertTo-MatchText $RightLocation
    $locationFamilies = @(
        @("bordeaux", "gironde", "nouvelle aquitaine"),
        @("lyon", "rhone", "auvergne rhone alpes"),
        @("lille", "nord", "hauts de france"),
        @("nantes", "loire atlantique", "pays de la loire"),
        @("toulouse", "haute garonne", "occitanie"),
        @("marseille", "bouches du rhone", "provence alpes cote azur")
    )
    foreach ($family in $locationFamilies) {
        $leftMatches = @($family | Where-Object { $leftText -match ("\b{0}\b" -f [regex]::Escape($_)) }).Count -gt 0
        $rightMatches = @($family | Where-Object { $rightText -match ("\b{0}\b" -f [regex]::Escape($_)) }).Count -gt 0
        if ($leftMatches -and $rightMatches) {
            return $true
        }
    }

    return $false
}

function Get-JobDuplicateKeysFromRow {
    param([AllowNull()]$Row)

    $keys = New-Object System.Collections.Generic.List[string]
    $url = Get-RowValue -Row $Row -Name "job_url_raw"
    $canonicalUrl = Get-CanonicalJobUrl $url
    if (-not [string]::IsNullOrWhiteSpace($canonicalUrl)) {
        $keys.Add(("hard-url|{0}" -f $canonicalUrl)) | Out-Null
    }

    $sourceIdKey = Get-SourceJobIdKeyFromUrl $url
    if (-not [string]::IsNullOrWhiteSpace($sourceIdKey)) {
        $keys.Add($sourceIdKey) | Out-Null
    }

    $title = Get-RowValue -Row $Row -Name "job_title"
    $company = Get-RowValue -Row $Row -Name "company_name"
    $location = Get-RowValue -Row $Row -Name "location"
    $titleKey = Get-DedupeTitleKey -Title $title -CompanyName $company
    $locationKey = Get-DedupeLocationKey $location
    foreach ($companyKey in @(Get-DedupeCompanyAliasKeys $company)) {
        if ([string]::IsNullOrWhiteSpace($companyKey) -or [string]::IsNullOrWhiteSpace($titleKey)) {
            continue
        }
        if (Test-IsStrongDedupeTitle $titleKey) {
            if (Test-UseLocationInDedupeKey $titleKey) {
                $keys.Add(("strong-company-title-location|{0}|{1}|{2}" -f $companyKey, $titleKey, $locationKey)) | Out-Null
            }
            else {
                $keys.Add(("strong-company-title|{0}|{1}" -f $companyKey, $titleKey)) | Out-Null
            }
        }

        $roleFamily = Get-DedupeRoleFamily $titleKey
        if (-not [string]::IsNullOrWhiteSpace($roleFamily) -and $roleFamily -notin @("data_analytics", "marketing_performance")) {
            $keys.Add(("role-family-location|{0}|{1}|{2}" -f $companyKey, $roleFamily, $locationKey)) | Out-Null
        }
    }

    return @($keys.ToArray() | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
}

function Test-JobRowsAreDuplicates {
    param(
        [AllowNull()]$Left,
        [AllowNull()]$Right,
        [switch]$AllowProbable
    )

    if ($null -eq $Left -or $null -eq $Right) {
        return $false
    }

    $leftUrl = Get-CanonicalJobUrl (Get-RowValue -Row $Left -Name "job_url_raw")
    $rightUrl = Get-CanonicalJobUrl (Get-RowValue -Row $Right -Name "job_url_raw")
    if (-not [string]::IsNullOrWhiteSpace($leftUrl) -and $leftUrl -eq $rightUrl) {
        return $true
    }

    $leftSourceId = Get-SourceJobIdKeyFromUrl (Get-RowValue -Row $Left -Name "job_url_raw")
    $rightSourceId = Get-SourceJobIdKeyFromUrl (Get-RowValue -Row $Right -Name "job_url_raw")
    if (-not [string]::IsNullOrWhiteSpace($leftSourceId) -and $leftSourceId -eq $rightSourceId) {
        return $true
    }

    $leftCompanies = @(Get-DedupeCompanyAliasKeys (Get-RowValue -Row $Left -Name "company_name"))
    $rightCompanies = @(Get-DedupeCompanyAliasKeys (Get-RowValue -Row $Right -Name "company_name"))
    $sharedCompanyKeys = @($leftCompanies | Where-Object { $rightCompanies -contains $_ })
    if ($sharedCompanyKeys.Count -eq 0) {
        return $false
    }

    $leftTitle = Get-DedupeTitleKey -Title (Get-RowValue -Row $Left -Name "job_title") -CompanyName (Get-RowValue -Row $Left -Name "company_name")
    $rightTitle = Get-DedupeTitleKey -Title (Get-RowValue -Row $Right -Name "job_title") -CompanyName (Get-RowValue -Row $Right -Name "company_name")
    if ([string]::IsNullOrWhiteSpace($leftTitle) -or [string]::IsNullOrWhiteSpace($rightTitle)) {
        return $false
    }

    $locationsCompatible = Test-DedupeLocationsCompatible -LeftLocation (Get-RowValue -Row $Left -Name "location") -RightLocation (Get-RowValue -Row $Right -Name "location")
    $titleSimilarity = Get-DedupeTitleSimilarity -LeftTitleKey $leftTitle -RightTitleKey $rightTitle
    $strongTitleThreshold = [double](Get-ConfigPathValue -Object $script:JobCrawlerMatchingRules -Path "deduplication.strong_title_similarity_threshold" -DefaultValue 0.72)
    $probableRoleThreshold = [double](Get-ConfigPathValue -Object $script:JobCrawlerMatchingRules -Path "deduplication.probable_role_similarity_threshold" -DefaultValue 0.35)

    if ($titleSimilarity -ge $strongTitleThreshold -and $locationsCompatible) {
        return $true
    }

    if ($AllowProbable) {
        $leftRoleFamily = Get-DedupeRoleFamily $leftTitle
        $rightRoleFamily = Get-DedupeRoleFamily $rightTitle
        if (-not [string]::IsNullOrWhiteSpace($leftRoleFamily) -and
            $leftRoleFamily -eq $rightRoleFamily -and
            $locationsCompatible -and
            $titleSimilarity -ge $probableRoleThreshold) {
            return $true
        }
    }

    return $false
}

function Get-JobDedupeKeyFromValues {
    param(
        [AllowNull()][string]$Title,
        [AllowNull()][string]$CompanyName,
        [AllowNull()][string]$JobLocation,
        [AllowNull()][string]$Url
    )

    $companyKeys = @(Get-DedupeCompanyAliasKeys $CompanyName)
    $companyKey = ""
    if ($companyKeys.Count -gt 0) {
        $companyKey = [string]$companyKeys[0]
    }
    $titleKey = Get-DedupeTitleKey -Title $Title -CompanyName $CompanyName
    $locationKey = Get-DedupeLocationKey $JobLocation

    if (-not [string]::IsNullOrWhiteSpace($companyKey) -and (Test-IsStrongDedupeTitle $titleKey)) {
        if (-not (Test-UseLocationInDedupeKey $titleKey)) {
            return "company-title|{0}|{1}" -f $companyKey, $titleKey
        }

        return "company-title-location|{0}|{1}|{2}" -f $companyKey, $titleKey, $locationKey
    }

    $urlKey = ""
    if (-not [string]::IsNullOrWhiteSpace($Url)) {
        $urlKey = ($Url.Trim().ToLowerInvariant() -replace "\?.*$", "")
    }
    return "url|{0}" -f $urlKey
}

function Get-StableJobId {
    param([string]$IdentityKey)

    if ([string]::IsNullOrWhiteSpace($IdentityKey)) {
        $IdentityKey = [Guid]::NewGuid().ToString("N")
    }

    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($IdentityKey)
        $hash = $sha.ComputeHash($bytes)
        return (($hash | ForEach-Object { $_.ToString("x2") }) -join "").Substring(0, 16)
    }
    finally {
        $sha.Dispose()
    }
}

function Get-JobIdentityKeyFromValues {
    param(
        [AllowNull()][string]$Title,
        [AllowNull()][string]$CompanyName,
        [AllowNull()][string]$JobLocation,
        [AllowNull()][string]$Url
    )

    return Get-JobDedupeKeyFromValues -Title $Title -CompanyName $CompanyName -JobLocation $JobLocation -Url $Url
}

function Get-JobIdentityKeyFromRow {
    param([AllowNull()]$Row)

    return Get-JobIdentityKeyFromValues `
        -Title (Get-RowValue -Row $Row -Name "job_title") `
        -CompanyName (Get-RowValue -Row $Row -Name "company_name") `
        -JobLocation (Get-RowValue -Row $Row -Name "location") `
        -Url (Get-RowValue -Row $Row -Name "job_url_raw")
}

function Get-JobDedupeKeyFromRow {
    param([AllowNull()]$Row)

    return Get-JobDedupeKeyFromValues `
        -Title (Get-RowValue -Row $Row -Name "job_title") `
        -CompanyName (Get-RowValue -Row $Row -Name "company_name") `
        -JobLocation (Get-RowValue -Row $Row -Name "location") `
        -Url (Get-RowValue -Row $Row -Name "job_url_raw")
}

function Get-PreferredValue {
    param(
        [AllowNull()][string]$Primary,
        [AllowNull()][string]$Fallback
    )

    if (-not [string]::IsNullOrWhiteSpace($Primary)) {
        return $Primary
    }

    return $Fallback
}

function Get-SourcePreference {
    param([AllowNull()]$Row)

    $platform = ConvertTo-MatchText (Get-RowValue -Row $Row -Name "platform")
    $url = ConvertTo-MatchText (Get-RowValue -Row $Row -Name "job_url_raw")
    if ($platform -match "welcome|jungle|wttj") {
        return 50
    }
    if ($platform -match "\bapec\b" -or $url -match "apec\.fr") {
        return 45
    }
    if ($platform -match "france\s+travail" -or $url -match "francetravail|pole-emploi") {
        return 40
    }
    if ($platform -match "hellowork" -or $url -match "hellowork") {
        return 35
    }
    if ($platform -match "linkedin") {
        return 30
    }
    if ($platform -match "adzuna") {
        return 15
    }

    return 10
}

function Get-DateSortValue {
    param(
        [AllowNull()]$Row,
        [string]$Name
    )

    $date = ConvertTo-DateOrNull (Get-RowValue -Row $Row -Name $Name)
    if ($null -eq $date) {
        return [DateTime]::MinValue
    }

    return $date
}

function Select-PreferredJobRow {
    param([object[]]$Rows)

    return @($Rows) |
        Sort-Object -Property `
            @{ Expression = { if (Test-IsAppliedStatus (Get-RowValue -Row $_ -Name "status")) { 1 } else { 0 } }; Descending = $true },
            @{ Expression = { if ((ConvertTo-MatchText (Get-RowValue -Row $_ -Name "status")) -eq "interesting") { 1 } else { 0 } }; Descending = $true },
            @{ Expression = { Get-IntegerRowValue -Row $_ -Name "match_score" }; Descending = $true },
            @{ Expression = { Get-DateSortValue -Row $_ -Name "published_date" }; Descending = $true },
            @{ Expression = { Get-SourcePreference $_ }; Descending = $true } |
        Select-Object -First 1
}

function Select-PreferredUrlRow {
    param([object[]]$Rows)

    return @($Rows) |
        Where-Object { -not [string]::IsNullOrWhiteSpace((Get-RowValue -Row $_ -Name "job_url_raw")) } |
        Sort-Object -Property `
            @{ Expression = { Get-SourcePreference $_ }; Descending = $true },
            @{ Expression = { Get-IntegerRowValue -Row $_ -Name "match_score" }; Descending = $true },
            @{ Expression = { Get-DateSortValue -Row $_ -Name "published_date" }; Descending = $true } |
        Select-Object -First 1
}

function Get-UniqueTextValues {
    param(
        [object[]]$Values,
        [string]$SplitPattern = "\s*;\s*"
    )

    $seen = @{}
    $result = New-Object System.Collections.Generic.List[string]
    foreach ($value in @($Values)) {
        if ($null -eq $value) {
            continue
        }

        foreach ($part in ([string]$value -split $SplitPattern)) {
            $text = ([string]$part).Trim()
            if ([string]::IsNullOrWhiteSpace($text)) {
                continue
            }

            $key = $text.ToLowerInvariant()
            if (-not $seen.ContainsKey($key)) {
                $seen[$key] = $true
                $result.Add($text) | Out-Null
            }
        }
    }

    return @($result.ToArray())
}

function Join-UniqueTextValues {
    param(
        [object[]]$Values,
        [string]$Delimiter = "; ",
        [string]$SplitPattern = "\s*;\s*"
    )

    return (Get-UniqueTextValues -Values $Values -SplitPattern $SplitPattern) -join $Delimiter
}

function Get-RowUrlValues {
    param([object[]]$Rows)

    $values = New-Object System.Collections.Generic.List[string]
    foreach ($row in @($Rows)) {
        $primary = Get-RowValue -Row $row -Name "job_url_raw"
        if (-not [string]::IsNullOrWhiteSpace($primary)) {
            $values.Add($primary) | Out-Null
        }

        $alternate = Get-RowValue -Row $row -Name "alternate_urls"
        foreach ($url in (Get-UniqueTextValues -Values @($alternate))) {
            $values.Add($url) | Out-Null
        }
    }

    return Get-UniqueTextValues -Values $values
}

function Get-RowPlatformValues {
    param([object[]]$Rows)

    return Get-UniqueTextValues -Values @($Rows | ForEach-Object { Get-RowValue -Row $_ -Name "platform" })
}

function Get-SourceCountFromRows {
    param([object[]]$Rows)

    $platforms = @(Get-RowPlatformValues $Rows)
    if ($platforms.Count -gt 0) {
        return [Math]::Max(1, $platforms.Count)
    }

    $urls = @(Get-RowUrlValues $Rows)
    return [Math]::Max(1, $urls.Count)
}

function Get-LatestDateText {
    param(
        [object[]]$Rows,
        [string]$Name
    )

    $dates = foreach ($row in @($Rows)) {
        $date = ConvertTo-DateOrNull (Get-RowValue -Row $row -Name $Name)
        if ($null -ne $date) {
            $date
        }
    }

    if ($null -eq $dates) {
        return ""
    }

    $latest = @($dates | Sort-Object -Descending | Select-Object -First 1)
    if ($latest.Count -eq 0) {
        return ""
    }

    return $latest[0].ToString("yyyy-MM-dd")
}

function Get-EarliestDateText {
    param(
        [object[]]$Rows,
        [string]$Name
    )

    $dates = foreach ($row in @($Rows)) {
        $date = ConvertTo-DateOrNull (Get-RowValue -Row $row -Name $Name)
        if ($null -ne $date) {
            $date
        }
    }

    if ($null -eq $dates) {
        return ""
    }

    $earliest = @($dates | Sort-Object | Select-Object -First 1)
    if ($earliest.Count -eq 0) {
        return ""
    }

    return $earliest[0].ToString("yyyy-MM-dd")
}

function Merge-SimilarJobRows {
    param(
        [object[]]$Rows,
        [string]$Reason
    )

    $rowList = @($Rows | Where-Object { $null -ne $_ })
    if ($rowList.Count -eq 0) {
        return $null
    }
    if ($rowList.Count -eq 1 -and [string]::IsNullOrWhiteSpace($Reason)) {
        return $rowList[0]
    }

    $preferred = Select-PreferredJobRow $rowList
    $preferredUrlRow = Select-PreferredUrlRow $rowList
    if ($null -eq $preferredUrlRow) {
        $preferredUrlRow = $preferred
    }

    $values = @{}
    foreach ($column in $MasterColumns) {
        $values[$column] = Get-RowValue -Row $preferred -Name $column
    }

    $urls = @(Get-RowUrlValues $rowList)
    $primaryUrl = Get-RowValue -Row $preferredUrlRow -Name "job_url_raw"
    if ([string]::IsNullOrWhiteSpace($primaryUrl) -and $urls.Count -gt 0) {
        $primaryUrl = $urls[0]
    }
    $alternateUrls = @($urls | Where-Object { $_ -ne $primaryUrl })

    $maxScore = 0
    $maxRoleScore = 0
    $maxEmployerFit = -999
    $maxLocationFit = -999
    $maxSeniorityFit = -999
    $maxContractFit = -999
    foreach ($row in $rowList) {
        $score = Get-IntegerRowValue -Row $row -Name "match_score"
        if ($score -gt $maxScore) {
            $maxScore = $score
        }
        $roleScore = Get-IntegerRowValue -Row $row -Name "role_score"
        if ($roleScore -gt $maxRoleScore) {
            $maxRoleScore = $roleScore
        }
        $employerFit = Get-IntegerRowValue -Row $row -Name "employer_fit"
        if ($employerFit -gt $maxEmployerFit) {
            $maxEmployerFit = $employerFit
        }
        $locationFit = Get-IntegerRowValue -Row $row -Name "location_fit"
        if ($locationFit -gt $maxLocationFit) {
            $maxLocationFit = $locationFit
        }
        $seniorityFit = Get-IntegerRowValue -Row $row -Name "seniority_fit"
        if ($seniorityFit -gt $maxSeniorityFit) {
            $maxSeniorityFit = $seniorityFit
        }
        $contractFit = Get-IntegerRowValue -Row $row -Name "contract_fit"
        if ($contractFit -gt $maxContractFit) {
            $maxContractFit = $contractFit
        }
    }

    $dedupeKey = Get-JobDedupeKeyFromRow $preferred
    $jobId = Get-RowValue -Row $preferred -Name "job_id"
    if ([string]::IsNullOrWhiteSpace($jobId)) {
        $jobId = Get-StableJobId $dedupeKey
    }

    $reasonParts = @($rowList | ForEach-Object { Get-RowValue -Row $_ -Name "duplicate_reason" })
    if (-not [string]::IsNullOrWhiteSpace($Reason)) {
        $reasonParts += $Reason
    }

    $values["job_id"] = $jobId
    $values["job_url_raw"] = $primaryUrl
    $values["job_url"] = ConvertTo-ExcelHyperlinkFormula -Url $primaryUrl -Label "Open"
    $values["alternate_urls"] = ($alternateUrls -join "; ")
    $values["source_count"] = [string](Get-SourceCountFromRows $rowList)
    $values["platform"] = Join-UniqueTextValues -Values @($rowList | ForEach-Object { Get-RowValue -Row $_ -Name "platform" })
    $values["matched_keywords"] = Join-UniqueTextValues -Values @($rowList | ForEach-Object { Get-RowValue -Row $_ -Name "matched_keywords" }) -SplitPattern "\s*;\s*|\s*,\s*"
    $employerTypes = @(Get-UniqueTextValues -Values @($rowList | ForEach-Object { Get-RowValue -Row $_ -Name "employer_type" }))
    foreach ($employerTypeCandidate in @("annonceur", "consulting", "agency", "esn", "unknown")) {
        if ($employerTypes -contains $employerTypeCandidate) {
            $values["employer_type"] = $employerTypeCandidate
            break
        }
    }
    $values["duplicate_reason"] = Join-CleanTextParts $reasonParts
    $values["match_score"] = [string]$maxScore
    $values["match_level"] = Get-MatchLevelFromScore $maxScore
    $values["role_score"] = [string]$maxRoleScore
    $values["employer_fit"] = $(if ($maxEmployerFit -eq -999) { "" } else { [string]$maxEmployerFit })
    $values["location_fit"] = $(if ($maxLocationFit -eq -999) { "" } else { [string]$maxLocationFit })
    $values["seniority_fit"] = $(if ($maxSeniorityFit -eq -999) { "" } else { [string]$maxSeniorityFit })
    $values["contract_fit"] = $(if ($maxContractFit -eq -999) { "" } else { [string]$maxContractFit })
    $values["fit_notes"] = Join-UniqueTextValues -Values @($rowList | ForEach-Object { Get-RowValue -Row $_ -Name "fit_notes" }) -SplitPattern "\s*;\s*"
    $values["published_date"] = Get-LatestDateText -Rows $rowList -Name "published_date"
    $values["first_seen_date"] = Get-EarliestDateText -Rows $rowList -Name "first_seen_date"
    $values["last_seen_date"] = Get-LatestDateText -Rows $rowList -Name "last_seen_date"
    $values["applied_date"] = Get-EarliestDateText -Rows $rowList -Name "applied_date"
    $values["notes"] = Join-UniqueTextValues -Values @($rowList | ForEach-Object { Get-RowValue -Row $_ -Name "notes" }) -Delimiter " | "

    return New-OrderedJobRecord $values
}

function Group-RowsByDedupeKey {
    param([object[]]$Rows)

    $rowList = @($Rows | Where-Object { $null -ne $_ })
    $groups = @{}
    if ($rowList.Count -eq 0) {
        return $groups
    }

    $parent = New-Object object[] $rowList.Count
    for ($index = 0; $index -lt $rowList.Count; $index++) {
        $parent[$index] = $index
    }

    function Find-DedupeGroupRoot {
        param([int]$Index)

        $root = $Index
        while ([int]$parent[$root] -ne $root) {
            $root = [int]$parent[$root]
        }

        $current = $Index
        while ([int]$parent[$current] -ne $current) {
            $next = [int]$parent[$current]
            $parent[$current] = $root
            $current = $next
        }

        return $root
    }

    function Join-DedupeGroupIndexes {
        param(
            [int]$Left,
            [int]$Right
        )

        $leftRoot = Find-DedupeGroupRoot $Left
        $rightRoot = Find-DedupeGroupRoot $Right
        if ($leftRoot -ne $rightRoot) {
            $parent[$rightRoot] = $leftRoot
        }
    }

    $keyToIndexes = @{}
    $companyBucketToIndexes = @{}
    for ($index = 0; $index -lt $rowList.Count; $index++) {
        $row = $rowList[$index]
        $keys = @(Get-JobDuplicateKeysFromRow $row)
        if ($keys.Count -eq 0) {
            $fallbackKey = Get-JobDedupeKeyFromRow $row
            if ([string]::IsNullOrWhiteSpace($fallbackKey)) {
                $fallbackKey = "jobid|{0}" -f (Get-RowValue -Row $row -Name "job_id")
            }
            $keys = @($fallbackKey)
        }

        foreach ($key in @($keys | Select-Object -Unique)) {
            if ([string]::IsNullOrWhiteSpace($key)) {
                continue
            }
            if (-not $keyToIndexes.ContainsKey($key)) {
                $keyToIndexes[$key] = New-Object System.Collections.Generic.List[int]
            }
            $keyToIndexes[$key].Add($index) | Out-Null
        }

        foreach ($companyKey in @(Get-DedupeCompanyAliasKeys (Get-RowValue -Row $row -Name "company_name"))) {
            if ([string]::IsNullOrWhiteSpace($companyKey)) {
                continue
            }
            if (-not $companyBucketToIndexes.ContainsKey($companyKey)) {
                $companyBucketToIndexes[$companyKey] = New-Object System.Collections.Generic.List[int]
            }
            $companyBucketToIndexes[$companyKey].Add($index) | Out-Null
        }
    }

    foreach ($key in $keyToIndexes.Keys) {
        $indexes = @($keyToIndexes[$key].ToArray())
        if ($indexes.Count -lt 2) {
            continue
        }
        $first = [int]$indexes[0]
        foreach ($index in @($indexes | Select-Object -Skip 1)) {
            Join-DedupeGroupIndexes -Left $first -Right ([int]$index)
        }
    }

    $seenPairs = @{}
    foreach ($companyKey in $companyBucketToIndexes.Keys) {
        $bucketIndexes = @($companyBucketToIndexes[$companyKey].ToArray() | Sort-Object -Unique)
        if ($bucketIndexes.Count -lt 2) {
            continue
        }
        for ($leftPosition = 0; $leftPosition -lt $bucketIndexes.Count; $leftPosition++) {
            for ($rightPosition = $leftPosition + 1; $rightPosition -lt $bucketIndexes.Count; $rightPosition++) {
                $leftIndex = [int]$bucketIndexes[$leftPosition]
                $rightIndex = [int]$bucketIndexes[$rightPosition]
                $pairKey = "{0}|{1}" -f ([Math]::Min($leftIndex, $rightIndex)), ([Math]::Max($leftIndex, $rightIndex))
                if ($seenPairs.ContainsKey($pairKey)) {
                    continue
                }
                $seenPairs[$pairKey] = $true
            if ((Find-DedupeGroupRoot $leftIndex) -eq (Find-DedupeGroupRoot $rightIndex)) {
                continue
            }
            if (Test-JobRowsAreDuplicates -Left $rowList[$leftIndex] -Right $rowList[$rightIndex] -AllowProbable) {
                Join-DedupeGroupIndexes -Left $leftIndex -Right $rightIndex
            }
            }
        }
    }

    $componentIndexes = @{}
    for ($index = 0; $index -lt $rowList.Count; $index++) {
        $root = [string](Find-DedupeGroupRoot $index)
        if (-not $componentIndexes.ContainsKey($root)) {
            $componentIndexes[$root] = New-Object System.Collections.Generic.List[int]
        }
        $componentIndexes[$root].Add($index) | Out-Null
    }

    foreach ($root in $componentIndexes.Keys) {
        $indexes = @($componentIndexes[$root].ToArray())
        $groupRows = @($indexes | ForEach-Object { $rowList[[int]$_] })
        $groupKeys = @($groupRows | ForEach-Object { Get-JobDuplicateKeysFromRow $_ } | Select-Object -Unique)
        $preferredKey = @($groupKeys | Where-Object { $_ -match "^hard-url\|" } | Sort-Object | Select-Object -First 1)
        if ($preferredKey.Count -eq 0) {
            $preferredKey = @($groupKeys | Where-Object { $_ -match "^source-id\|" } | Sort-Object | Select-Object -First 1)
        }
        if ($preferredKey.Count -eq 0) {
            $preferredKey = @($groupKeys | Where-Object { $_ -match "^strong-company-title" } | Sort-Object | Select-Object -First 1)
        }
        if ($preferredKey.Count -eq 0) {
            $preferredKey = @($groupKeys | Where-Object { $_ -match "^role-family-location" } | Sort-Object | Select-Object -First 1)
        }
        $key = ""
        if ($preferredKey.Count -gt 0) {
            $key = [string]$preferredKey[0]
        }
        else {
            $key = Get-JobDedupeKeyFromRow $groupRows[0]
        }
        if ([string]::IsNullOrWhiteSpace($key)) {
            $key = "jobid|{0}" -f (Get-RowValue -Row $groupRows[0] -Name "job_id")
        }

        if (-not $groups.ContainsKey($key)) {
            $groups[$key] = New-Object System.Collections.Generic.List[object]
        }
        foreach ($row in $groupRows) {
            $groups[$key].Add($row) | Out-Null
        }
    }

    return $groups
}

function Find-DuplicateGroupKey {
    param(
        [AllowNull()]$Row,
        [hashtable]$GroupsByKey,
        [string[]]$ExcludedKeys = @()
    )

    if ($null -eq $Row -or $null -eq $GroupsByKey) {
        return ""
    }

    foreach ($key in @($GroupsByKey.Keys | Sort-Object)) {
        if ($ExcludedKeys -contains $key) {
            continue
        }
        foreach ($candidate in @($GroupsByKey[$key].ToArray())) {
            if (Test-JobRowsAreDuplicates -Left $Row -Right $candidate -AllowProbable) {
                return [string]$key
            }
        }
    }

    return ""
}

function Test-IsInvalidExistingTrackerRow {
    param([AllowNull()]$Row)

    if ($null -eq $Row) {
        return $false
    }

    $platform = Get-RowValue -Row $Row -Name "platform"
    if ($platform -notmatch "(?i)welcome\s+to\s+the\s+jungle|wttj") {
        return $false
    }

    $locationCheck = Get-Command "Test-IsWttjLocationAllowed" -ErrorAction SilentlyContinue
    if ($null -eq $locationCheck) {
        return $false
    }

    $location = Get-RowValue -Row $Row -Name "location"
    $url = Get-RowValue -Row $Row -Name "job_url_raw"
    if ([string]::IsNullOrWhiteSpace($location)) {
        return $true
    }
    if (Test-IsWttjKnownForeignLocation $location) {
        return $true
    }
    if (Test-IsWttjKnownForeignLocation (Get-WttjLocationFromUrl $url)) {
        return $true
    }

    $text = Join-CleanTextParts @(
        (Get-RowValue -Row $Row -Name "job_title"),
        (Get-RowValue -Row $Row -Name "company_name"),
        $url
    )

    return (-not (Test-IsWttjLocationAllowed -JobLocation $location -Url $url -Text $text))
}

function Merge-JobsWithTracker {
    param(
        [object[]]$CurrentRows,
        [object[]]$ExistingRows,
        [string]$Path,
        [switch]$SkipBackup
    )

    $existingRecords = New-Object System.Collections.Generic.List[object]
    foreach ($existing in @($ExistingRows)) {
        $record = ConvertTo-TrackerRecordFromExisting $existing
        if ($null -ne $record) {
            $existingRecords.Add($record) | Out-Null
        }
    }

    $duplicateCount = 0
    $existingByKey = @{}
    $existingGroupCountsByKey = @{}
    $existingGroups = Group-RowsByDedupeKey -Rows @($existingRecords.ToArray())
    foreach ($key in $existingGroups.Keys) {
        $groupRows = @($existingGroups[$key].ToArray())
        $existingGroupCountsByKey[$key] = $groupRows.Count
        if ($groupRows.Count -gt 1) {
            $duplicateCount += ($groupRows.Count - 1)
        }

        $reason = ""
        if ($groupRows.Count -gt 1) {
            $reason = "merged similar tracker rows"
        }
        $existingByKey[$key] = Merge-SimilarJobRows -Rows $groupRows -Reason $reason
    }

    $currentInputGate = Invoke-JobPipelineEligibilityGate -Rows @($CurrentRows) -Stage "merge_current_input"
    $currentByKey = @{}
    $currentGroups = Group-RowsByDedupeKey -Rows @($currentInputGate.KeptRows)
    foreach ($key in $currentGroups.Keys) {
        $groupRows = @($currentGroups[$key].ToArray())
        if ($groupRows.Count -gt 1) {
            $duplicateCount += ($groupRows.Count - 1)
        }

        $reasonParts = New-Object System.Collections.Generic.List[string]
        if ($groupRows.Count -gt 1) {
            $reasonParts.Add(("merged {0} similar current postings" -f $groupRows.Count)) | Out-Null
        }

        $platforms = @(Get-UniqueTextValues -Values @($groupRows | ForEach-Object { Get-RowValue -Row $_ -Name "platform" }))
        if ($platforms.Count -gt 1) {
            $reasonParts.Add(("same job found on multiple sources: {0}" -f ($platforms -join "; "))) | Out-Null
        }

        $currentByKey[$key] = Merge-SimilarJobRows -Rows $groupRows -Reason (Join-CleanTextParts $reasonParts.ToArray())
    }

    $trackerByKey = @{}
    $currentTrackerRows = New-Object System.Collections.Generic.List[object]
    $removedCount = 0
    foreach ($key in $currentByKey.Keys) {
        $existingKey = $key
        if (-not $existingByKey.ContainsKey($existingKey)) {
            $existingKey = Find-DuplicateGroupKey -Row $currentByKey[$key] -GroupsByKey $existingGroups -ExcludedKeys @($trackerByKey.Keys)
        }

        $existing = $null
        if (-not [string]::IsNullOrWhiteSpace($existingKey) -and $existingByKey.ContainsKey($existingKey)) {
            $existing = $existingByKey[$existingKey]
        }

        $duplicateReason = $(if ($null -ne $existing) { "same hierarchical duplicate identity from previous crawl" } else { "" })
        $currentDuplicateReason = Get-RowValue -Row $currentByKey[$key] -Name "duplicate_reason"
        if (-not [string]::IsNullOrWhiteSpace($currentDuplicateReason)) {
            $duplicateReason = Join-CleanTextParts @($duplicateReason, $currentDuplicateReason)
        }
        $trackerRecord = ConvertTo-TrackerRecord -CurrentRow $currentByKey[$key] -ExistingRow $existing -SeenInCurrentCrawl:$true -DuplicateReason $duplicateReason
        $trackerDecision = Get-JobPipelineEligibility -Row $trackerRecord -CurrentRow $currentByKey[$key] -ExistingRow $existing -Stage "merge_current_final"
        if ($trackerDecision.IsEligible) {
            $trackerKey = $(if (-not [string]::IsNullOrWhiteSpace($existingKey) -and $existingByKey.ContainsKey($existingKey)) { $existingKey } else { $key })
            $trackerByKey[$trackerKey] = $trackerRecord
            $currentTrackerRows.Add($trackerRecord) | Out-Null
        }
    }

    $preservedAppliedCount = 0
    foreach ($key in $existingByKey.Keys) {
        if ($trackerByKey.ContainsKey($key)) {
            continue
        }

        $existing = $existingByKey[$key]
        $existingDecision = Get-JobPipelineEligibility -Row $existing -ExistingRow $existing -Stage "merge_existing_retention"
        if ($existingDecision.KeepForever) {
            $trackerByKey[$key] = ConvertTo-TrackerRecord -CurrentRow $existing -ExistingRow $existing -SeenInCurrentCrawl:$false -DuplicateReason "kept by application status"
            $preservedAppliedCount++
        }
        elseif ($existingDecision.IsEligible) {
            $trackerByKey[$key] = ConvertTo-TrackerRecord -CurrentRow $existing -ExistingRow $existing -SeenInCurrentCrawl:$false -DuplicateReason "not seen this crawl, still inside published-date retention window"
        }
        else {
            $removedCount += [Math]::Max(1, [int]$existingGroupCountsByKey[$key])
        }
    }

    $trackerRows = @($trackerByKey.Values) |
        Sort-Object -Property `
            @{ Expression = { if ((Get-RowValue -Row $_ -Name "seen_in_current_crawl") -eq "yes") { 1 } else { 0 } }; Descending = $true },
            @{ Expression = { try { [int](Get-RowValue -Row $_ -Name "match_score") } catch { 0 } }; Descending = $true },
            @{ Expression = "published_date"; Descending = $true },
            platform,
            job_title

    Test-JobPipelineInvariants -Rows $trackerRows -Stage "merge_output" -ThrowOnIssue | Out-Null

    $backupPath = ""
    if (-not $SkipBackup) {
        $backupPath = Backup-TrackerFile -Path $Path
    }

    return @{
        TrackerRows = @($trackerRows)
        CurrentRows = @($currentTrackerRows.ToArray())
        RemovedCount = $removedCount
        RejectedCurrentCount = $currentInputGate.ExcludedCount
        RejectedCurrentReasons = $currentInputGate.CountByReason
        DuplicateCount = $duplicateCount
        PreservedAppliedCount = $preservedAppliedCount
        BackupPath = $backupPath
    }
}

