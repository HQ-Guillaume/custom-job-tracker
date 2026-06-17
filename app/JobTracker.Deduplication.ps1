# Auto-extracted from Find-AnalyticsJobs.ps1. Keep dot-sourced execution order in the main script.

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

function Get-DedupeCompanyKey {
    param([AllowNull()][string]$CompanyName)

    if (Test-IsGenericJobBoardName $CompanyName) {
        return ""
    }

    $tokens = @(Split-NormalizedTokens $CompanyName)
    if ($tokens.Count -eq 0) {
        return ""
    }

    $noise = @(
        "the", "and", "et", "de", "du", "des", "la", "le", "les",
        "sa", "sas", "sasu", "ltd", "limited", "inc", "plc", "fr", "france",
        "group", "groupe", "company", "companies", "media", "digital",
        "consulting", "consultants", "technology", "technologies", "solutions"
    )
    $weakCompanyTokens = @("confidential", "confidentiel", "jobgether", "licorne", "recrutement", "talent", "emploi", "travail", "adzuna", "linkedin", "indeed", "hellowork", "meteojob", "jobijoba", "monster", "apec")
    $strongTokens = @($tokens | Where-Object { $_.Length -gt 1 -and $noise -notcontains $_ -and $weakCompanyTokens -notcontains $_ })
    if ($strongTokens.Count -eq 0) {
        return ""
    }

    if ($strongTokens.Count -eq 1) {
        return $strongTokens[0]
    }

    return ($strongTokens | Select-Object -First 2) -join " "
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

function Get-JobDedupeKeyFromValues {
    param(
        [AllowNull()][string]$Title,
        [AllowNull()][string]$CompanyName,
        [AllowNull()][string]$JobLocation,
        [AllowNull()][string]$Url
    )

    $companyKey = Get-DedupeCompanyKey $CompanyName
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

    $groups = @{}
    foreach ($row in @($Rows)) {
        $key = Get-JobDedupeKeyFromRow $row
        if ([string]::IsNullOrWhiteSpace($key)) {
            $key = "jobid|{0}" -f (Get-RowValue -Row $row -Name "job_id")
        }

        if (-not $groups.ContainsKey($key)) {
            $groups[$key] = New-Object System.Collections.Generic.List[object]
        }
        $groups[$key].Add($row) | Out-Null
    }

    return $groups
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
    $existingGroups = Group-RowsByDedupeKey -Rows @($existingRecords.ToArray())
    foreach ($key in $existingGroups.Keys) {
        $groupRows = @($existingGroups[$key].ToArray())
        if ($groupRows.Count -gt 1) {
            $duplicateCount += ($groupRows.Count - 1)
        }

        $reason = ""
        if ($groupRows.Count -gt 1) {
            $reason = "merged similar tracker rows"
        }
        $existingByKey[$key] = Merge-SimilarJobRows -Rows $groupRows -Reason $reason
    }

    $currentByKey = @{}
    $currentGroups = Group-RowsByDedupeKey -Rows @($CurrentRows)
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
    foreach ($key in $currentByKey.Keys) {
        $existing = $null
        if ($existingByKey.ContainsKey($key)) {
            $existing = $existingByKey[$key]
        }

        $duplicateReason = $(if ($null -ne $existing) { "same normalized company/title/location from previous crawl" } else { "" })
        $currentDuplicateReason = Get-RowValue -Row $currentByKey[$key] -Name "duplicate_reason"
        if (-not [string]::IsNullOrWhiteSpace($currentDuplicateReason)) {
            $duplicateReason = Join-CleanTextParts @($duplicateReason, $currentDuplicateReason)
        }
        $trackerByKey[$key] = ConvertTo-TrackerRecord -CurrentRow $currentByKey[$key] -ExistingRow $existing -SeenInCurrentCrawl:$true -DuplicateReason $duplicateReason
    }

    $removedCount = 0
    $preservedAppliedCount = 0
    foreach ($key in $existingByKey.Keys) {
        if ($trackerByKey.ContainsKey($key)) {
            continue
        }

        $existing = $existingByKey[$key]
        if (Test-IsKeepForeverStatus (Get-RowValue -Row $existing -Name "status")) {
            $trackerByKey[$key] = ConvertTo-TrackerRecord -CurrentRow $existing -ExistingRow $existing -SeenInCurrentCrawl:$false -DuplicateReason "kept by application status"
            $preservedAppliedCount++
        }
        elseif (Test-IsInvalidExistingTrackerRow $existing) {
            $removedCount++
        }
        elseif (Test-IsExcludedContractType (Get-RowValue -Row $existing -Name "contract_type")) {
            $removedCount++
        }
        elseif (Test-IsRecentTrackerRow $existing) {
            $trackerByKey[$key] = ConvertTo-TrackerRecord -CurrentRow $existing -ExistingRow $existing -SeenInCurrentCrawl:$false -DuplicateReason "not seen this crawl, still inside published-date retention window"
        }
        else {
            $removedCount++
        }
    }

    $trackerRows = @($trackerByKey.Values) |
        Sort-Object -Property `
            @{ Expression = { if ((Get-RowValue -Row $_ -Name "seen_in_current_crawl") -eq "yes") { 1 } else { 0 } }; Descending = $true },
            @{ Expression = { try { [int](Get-RowValue -Row $_ -Name "match_score") } catch { 0 } }; Descending = $true },
            @{ Expression = "published_date"; Descending = $true },
            platform,
            job_title

    $backupPath = ""
    if (-not $SkipBackup) {
        $backupPath = Backup-TrackerFile -Path $Path
    }

    return @{
        TrackerRows = @($trackerRows)
        CurrentRows = @($currentByKey.Values)
        RemovedCount = $removedCount
        DuplicateCount = $duplicateCount
        PreservedAppliedCount = $preservedAppliedCount
        BackupPath = $backupPath
    }
}

