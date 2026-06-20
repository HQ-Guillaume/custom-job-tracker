# Profile builder and expansion helpers for configurable job-search intents.

function ConvertTo-JobCrawlerProfileId {
    param([AllowNull()][string]$Text)

    $raw = ([string]$Text).Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($raw)) {
        return ""
    }

    $normalized = $raw.Normalize([System.Text.NormalizationForm]::FormD)
    $builder = New-Object System.Text.StringBuilder
    foreach ($char in $normalized.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($char) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$builder.Append($char)
        }
    }

    $id = [regex]::Replace($builder.ToString(), "[^a-z0-9]+", "_").Trim("_")
    return $id
}

function ConvertTo-JobCrawlerPlainText {
    param([AllowNull()][string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ""
    }

    $normalized = ([string]$Text).Trim().Normalize([System.Text.NormalizationForm]::FormD)
    $builder = New-Object System.Text.StringBuilder
    foreach ($char in $normalized.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($char) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$builder.Append($char)
        }
    }

    return $builder.ToString().ToLowerInvariant()
}

function ConvertTo-JobCrawlerProfileLineArray {
    param([AllowNull()]$Value)

    $items = New-Object System.Collections.Generic.List[string]
    if ($null -eq $Value) {
        return @()
    }

    $rawItems = @()
    if ($Value -is [string]) {
        $rawItems = @(([string]$Value) -split "(`r`n|`n|,|;)")
    }
    else {
        $rawItems = @($Value)
    }

    foreach ($item in $rawItems) {
        $text = ([string]$item).Trim()
        if ([string]::IsNullOrWhiteSpace($text)) {
            continue
        }
        if (-not $items.Contains($text)) {
            $items.Add($text) | Out-Null
        }
    }

    return @($items.ToArray())
}

function ConvertTo-JobCrawlerRegexAlternation {
    param(
        [AllowNull()][string[]]$Terms,
        [string]$FallbackPattern = "(?!)"
    )

    $patterns = New-Object System.Collections.Generic.List[string]
    foreach ($term in @(ConvertTo-JobCrawlerProfileLineArray $Terms)) {
        $plain = ConvertTo-JobCrawlerPlainText $term
        if ([string]::IsNullOrWhiteSpace($plain)) {
            continue
        }

        $escaped = [regex]::Escape($plain)
        $escaped = $escaped -replace "\\ ", "\s+"
        $escaped = $escaped -replace "-", "[-_\s]*"
        $patterns.Add($escaped) | Out-Null
    }

    if ($patterns.Count -eq 0) {
        return $FallbackPattern
    }

    return (($patterns.ToArray() | Sort-Object -Unique) -join "|")
}

function New-JobCrawlerProfileSignal {
    param(
        [string]$Scope,
        [string]$Pattern,
        [string]$Keyword,
        [int]$Score
    )

    return [ordered]@{
        scope   = $Scope
        pattern = $Pattern
        keyword = $Keyword
        score   = $Score
    }
}

function Add-JobCrawlerProfileUniqueString {
    param(
        $List,
        [AllowNull()][string]$Value
    )

    if ($null -eq $List) {
        return
    }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return
    }

    $key = ConvertTo-JobCrawlerPlainText $text
    foreach ($existing in @($List.ToArray())) {
        if ((ConvertTo-JobCrawlerPlainText $existing) -eq $key) {
            return
        }
    }

    $List.Add($text) | Out-Null
}

function Get-JobCrawlerProfileRoleVariants {
    param([AllowNull()][string]$Title)

    $variants = New-Object System.Collections.Generic.List[string]
    $plain = ConvertTo-JobCrawlerPlainText $Title
    if ([string]::IsNullOrWhiteSpace($plain)) {
        return @()
    }

    $clean = [regex]::Replace($plain, "\([^)]*\)|\[[^\]]*\]", " ")
    $clean = [regex]::Replace($clean, "\b(h|f|x|m|cdi|cdd|stage|alternance|apprentissage)\b", " ")
    $clean = [regex]::Replace($clean, "\b(senior|junior|lead|head|principal)\b", " ")
    $clean = [regex]::Replace($clean, "\s+", " ").Trim()
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return @()
    }

    if ($clean -match "^(?<base>.+?)\s+(analyst|analyste)$") {
        $base = [string]$Matches["base"]
        Add-JobCrawlerProfileUniqueString -List $variants -Value ("{0} analytics" -f $base)
        Add-JobCrawlerProfileUniqueString -List $variants -Value ("{0} analysis" -f $base)
        Add-JobCrawlerProfileUniqueString -List $variants -Value ("{0} insights" -f $base)
    }
    if ($clean -match "^(?<base>.+?)\s+(manager|lead)$") {
        $base = [string]$Matches["base"]
        Add-JobCrawlerProfileUniqueString -List $variants -Value ("{0} management" -f $base)
        Add-JobCrawlerProfileUniqueString -List $variants -Value ("{0} lead" -f $base)
    }
    if ($clean -match "^(?<base>.+?)\s+(developer|developpeur|developpeuse|dev)$") {
        $base = [string]$Matches["base"]
        Add-JobCrawlerProfileUniqueString -List $variants -Value ("{0} development" -f $base)
        Add-JobCrawlerProfileUniqueString -List $variants -Value ("{0} engineer" -f $base)
    }
    if ($clean -match "^(?<base>.+?)\s+(engineer|ingenieur|ingenieure)$") {
        $base = [string]$Matches["base"]
        Add-JobCrawlerProfileUniqueString -List $variants -Value ("{0} developer" -f $base)
        Add-JobCrawlerProfileUniqueString -List $variants -Value ("{0} engineering" -f $base)
    }
    if ($clean -match "^(?<base>.+?)\s+(consultant|consultante)$") {
        $base = [string]$Matches["base"]
        Add-JobCrawlerProfileUniqueString -List $variants -Value ("{0} consulting" -f $base)
        Add-JobCrawlerProfileUniqueString -List $variants -Value ("{0} conseil" -f $base)
    }
    if ($clean -match "business\s+development") {
        Add-JobCrawlerProfileUniqueString -List $variants -Value "business developer"
        Add-JobCrawlerProfileUniqueString -List $variants -Value "sales development"
        Add-JobCrawlerProfileUniqueString -List $variants -Value "bizdev"
    }

    return @($variants.ToArray())
}

function Get-JobCrawlerProfileNgrams {
    param(
        [AllowNull()][string]$Text,
        [int]$MinTokens = 2,
        [int]$MaxTokens = 3
    )

    $plain = ConvertTo-JobCrawlerPlainText $Text
    if ([string]::IsNullOrWhiteSpace($plain)) {
        return @()
    }

    $stop = @(
        "a", "an", "and", "avec", "cdd", "cdi", "de", "des", "du", "en", "et", "f", "for", "h", "la", "le",
        "les", "m", "of", "pour", "sur", "the", "x", "junior", "senior", "stage", "alternance", "apprentissage"
    )
    $tokens = New-Object System.Collections.Generic.List[string]
    foreach ($match in [regex]::Matches($plain, "[a-z0-9]{2,}")) {
        $token = [string]$match.Value
        if ($token -notin $stop) {
            $tokens.Add($token) | Out-Null
        }
    }

    if ($tokens.Count -lt $MinTokens) {
        return @()
    }

    $phrases = New-Object System.Collections.Generic.List[string]
    for ($size = $MinTokens; $size -le $MaxTokens; $size++) {
        if ($tokens.Count -lt $size) {
            continue
        }
        $tokenArray = @($tokens.ToArray())
        for ($i = 0; $i -le ($tokens.Count - $size); $i++) {
            $phrase = (($tokenArray[$i..($i + $size - 1)]) -join " ")
            if ($phrase.Length -ge 6) {
                Add-JobCrawlerProfileUniqueString -List $phrases -Value $phrase
            }
        }
    }

    return @($phrases.ToArray())
}

function Get-JobCrawlerSearchQuerySuggestions {
    param(
        [AllowNull()][string]$Label = "",
        [AllowNull()][string[]]$TargetTitles = @(),
        [AllowNull()][string[]]$ImportantSkills = @(),
        [int]$MaxQueries = 24
    )

    $queries = New-Object System.Collections.Generic.List[string]
    $titles = @(ConvertTo-JobCrawlerProfileLineArray $TargetTitles)
    if ($titles.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace($Label)) {
        $titles = @(([string]$Label).Trim())
    }
    $skills = @(ConvertTo-JobCrawlerProfileLineArray $ImportantSkills)

    foreach ($title in ($titles | Select-Object -First 10)) {
        Add-JobCrawlerProfileUniqueString -List $queries -Value $title
    }
    foreach ($title in ($titles | Select-Object -First 8)) {
        foreach ($variant in @(Get-JobCrawlerProfileRoleVariants -Title $title)) {
            Add-JobCrawlerProfileUniqueString -List $queries -Value $variant
        }
    }
    foreach ($title in ($titles | Select-Object -First 6)) {
        foreach ($phrase in @(Get-JobCrawlerProfileNgrams -Text $title -MinTokens 2 -MaxTokens 3)) {
            Add-JobCrawlerProfileUniqueString -List $queries -Value $phrase
        }
    }
    foreach ($title in ($titles | Select-Object -First 5)) {
        foreach ($skill in ($skills | Select-Object -First 7)) {
            Add-JobCrawlerProfileUniqueString -List $queries -Value ("{0} {1}" -f $title, $skill)
        }
    }
    foreach ($skill in ($skills | Select-Object -First 8)) {
        $plainSkill = ConvertTo-JobCrawlerPlainText $skill
        if ($plainSkill.Length -ge 4 -and ($plainSkill -match "\s" -or $plainSkill -match "^[a-z0-9]{2,8}$")) {
            Add-JobCrawlerProfileUniqueString -List $queries -Value $skill
        }
    }

    return @($queries.ToArray() | Select-Object -First $MaxQueries)
}

function Get-JobCrawlerProfileQuality {
    param(
        [AllowNull()][string]$Label = "",
        [AllowNull()][string[]]$TargetTitles = @(),
        [AllowNull()][string[]]$SearchQueries = @(),
        [AllowNull()][string[]]$ImportantSkills = @(),
        [AllowNull()][string[]]$ExclusionKeywords = @(),
        [AllowNull()][string[]]$TargetLocations = @(),
        [AllowNull()][string[]]$ExcludedContracts = @()
    )

    $titles = @(ConvertTo-JobCrawlerProfileLineArray $TargetTitles)
    $queries = @(ConvertTo-JobCrawlerProfileLineArray $SearchQueries)
    $skills = @(ConvertTo-JobCrawlerProfileLineArray $ImportantSkills)
    $exclusions = @(ConvertTo-JobCrawlerProfileLineArray $ExclusionKeywords)
    $locations = @(ConvertTo-JobCrawlerProfileLineArray $TargetLocations)
    $contracts = @(ConvertTo-JobCrawlerProfileLineArray $ExcludedContracts)
    $suggestions = @(Get-JobCrawlerSearchQuerySuggestions -Label $Label -TargetTitles $titles -ImportantSkills $skills -MaxQueries 24)

    $score = 0
    if ($titles.Count -ge 5) { $score += 20 }
    elseif ($titles.Count -ge 2) { $score += 16 }
    elseif ($titles.Count -eq 1) { $score += 10 }

    if ($queries.Count -ge 8) { $score += 30 }
    elseif ($queries.Count -ge 4) { $score += 24 }
    elseif ($queries.Count -ge 2) { $score += 14 }
    elseif ($queries.Count -eq 1) { $score += 6 }

    if ($skills.Count -ge 6) { $score += 20 }
    elseif ($skills.Count -ge 3) { $score += 14 }
    elseif ($skills.Count -ge 1) { $score += 8 }

    if ($exclusions.Count -ge 3) { $score += 12 }
    elseif ($exclusions.Count -ge 1) { $score += 6 }

    if ($locations.Count -ge 1) { $score += 8 }
    if ($contracts.Count -ge 3) { $score += 5 }
    elseif ($contracts.Count -ge 1) { $score += 3 }

    $titleKeys = @($titles | ForEach-Object { ConvertTo-JobCrawlerPlainText $_ })
    $queryKeys = @($queries | ForEach-Object { ConvertTo-JobCrawlerPlainText $_ })
    $exactTitleOnly = $false
    if ($queries.Count -gt 0 -and $queries.Count -le [Math]::Max(1, $titles.Count)) {
        $nonTitleQueries = @($queryKeys | Where-Object { $_ -notin $titleKeys })
        $exactTitleOnly = ($nonTitleQueries.Count -eq 0)
    }

    $findings = New-Object System.Collections.Generic.List[string]
    if ($titles.Count -lt 2) {
        $findings.Add("Add related target titles so matching is not tied to one exact wording.") | Out-Null
    }
    if ($queries.Count -lt 4) {
        $findings.Add("Use at least 4 search queries; one exact query can easily return 0 results.") | Out-Null
    }
    if ($exactTitleOnly) {
        $findings.Add("Queries are too close to the exact titles; add broader wording and title + skill combinations.") | Out-Null
    }
    if ($skills.Count -lt 3) {
        $findings.Add("Add important skills, tools, industries, or mission keywords to separate good jobs from broad noise.") | Out-Null
    }
    if ($exclusions.Count -eq 0) {
        $findings.Add("Add excluded keywords for jobs you regularly reject.") | Out-Null
    }
    if ($locations.Count -eq 0) {
        $findings.Add("Add target locations or country names to avoid irrelevant geographies.") | Out-Null
    }
    if ($contracts.Count -eq 0) {
        $findings.Add("Select excluded contract types if some contracts are not useful for this search.") | Out-Null
    }
    if ($findings.Count -eq 0) {
        $findings.Add("Profile has enough breadth for crawling and enough criteria for filtering.") | Out-Null
    }

    $level = "Risky"
    if ($score -ge 85) { $level = "Strong" }
    elseif ($score -ge 70) { $level = "Good" }
    elseif ($score -ge 50) { $level = "Needs work" }

    return [PSCustomObject]@{
        Score            = [Math]::Min(100, $score)
        Level            = $level
        Findings         = @($findings.ToArray())
        QuerySuggestions = @($suggestions)
    }
}

function Write-JobCrawlerJsonConfig {
    param(
        [string]$Path,
        [AllowNull()]$Value
    )

    $directory = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    $json = $Value | ConvertTo-Json -Depth 100
    $encoding = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, ($json + [Environment]::NewLine), $encoding)
}

function Get-JobCrawlerContractPatternFromNames {
    param([AllowNull()][string[]]$ContractNames)

    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($contractName in @(ConvertTo-JobCrawlerProfileLineArray $ContractNames)) {
        switch -Regex ((ConvertTo-JobCrawlerPlainText $contractName)) {
            "^cdd$|fixed|temporary" {
                $parts.Add("\bcdd\b|fixed[-\s]*term|temporary|temporaire|contrat\s+a\s+duree\s+determinee|\b\d+\s*(mois|months)\b") | Out-Null
            }
            "interim|int[eé]rim" {
                $parts.Add("interim|int[eé]rim") | Out-Null
            }
            "^mission$|contract\s+mission" {
                $parts.Add("mission\s+(de\s+)?\d+\s*(mois|months)|contrat\s+de\s+mission|mission\s+freelance") | Out-Null
            }
            "apprentice|apprentissage|alternance" {
                $parts.Add("apprenticeship|apprenti|apprentissage|alternance") | Out-Null
            }
            "intern|stage|stagiaire" {
                $parts.Add("internship|intern\b|stagiaire|\bstage\b") | Out-Null
            }
            "freelance|contractor|independent|independant" {
                $parts.Add("freelance|independent|independant|contractor") | Out-Null
            }
        }
    }

    if ($parts.Count -eq 0) {
        return "(?!)"
    }

    return (($parts.ToArray() | Sort-Object -Unique) -join "|")
}

function New-JobCrawlerProfileFromBuilder {
    param(
        [string]$Label,
        [string]$Id = "",
        [AllowNull()][string]$Description = "",
        [AllowNull()][string[]]$TargetTitles = @(),
        [AllowNull()][string[]]$ImportantSkills = @(),
        [AllowNull()][string[]]$ExclusionKeywords = @(),
        [AllowNull()][string[]]$SearchQueries = @(),
        [AllowNull()][string[]]$TargetLocations = @(),
        [AllowNull()][string[]]$ExcludedLocations = @(),
        [AllowNull()][string[]]$ExcludedContracts = @("CDD", "Apprenticeship", "Internship", "Freelance"),
        [string]$EmployerPreference = "neutral",
        [switch]$Compact
    )

    $profileLabel = ([string]$Label).Trim()
    if ([string]::IsNullOrWhiteSpace($profileLabel)) {
        throw "Profile name is required."
    }

    $profileId = ConvertTo-JobCrawlerProfileId $(if ([string]::IsNullOrWhiteSpace($Id)) { $profileLabel } else { $Id })
    if ([string]::IsNullOrWhiteSpace($profileId)) {
        throw "Profile id could not be generated from '$profileLabel'."
    }

    $titles = @(ConvertTo-JobCrawlerProfileLineArray $TargetTitles)
    if ($titles.Count -eq 0) {
        $titles = @($profileLabel)
    }
    $skills = @(ConvertTo-JobCrawlerProfileLineArray $ImportantSkills)
    $exclusions = @(ConvertTo-JobCrawlerProfileLineArray $ExclusionKeywords)
    $targetLocations = @(ConvertTo-JobCrawlerProfileLineArray $TargetLocations)
    if ($targetLocations.Count -eq 0) {
        $targetLocations = @("France")
    }
    $excludedLocations = @(ConvertTo-JobCrawlerProfileLineArray $ExcludedLocations)
    $excludedContractsClean = @(ConvertTo-JobCrawlerProfileLineArray $ExcludedContracts)
    $queries = @(ConvertTo-JobCrawlerProfileLineArray $SearchQueries)
    $suggestedQueries = @(Get-JobCrawlerSearchQuerySuggestions -Label $profileLabel -TargetTitles $titles -ImportantSkills $skills -MaxQueries 24)

    if ($queries.Count -eq 0) {
        $queries = @($suggestedQueries)
    }
    elseif ($queries.Count -lt 4) {
        $expandedQueries = New-Object System.Collections.Generic.List[string]
        foreach ($query in @($queries + $suggestedQueries)) {
            Add-JobCrawlerProfileUniqueString -List $expandedQueries -Value $query
        }
        $queries = @($expandedQueries.ToArray() | Select-Object -First 24)
    }

    $compactProfile = [ordered]@{
        id              = $profileId
        label           = $profileLabel
        description     = $(if ([string]::IsNullOrWhiteSpace($Description)) { "Custom job crawler profile." } else { [string]$Description })
        version         = 1
        profile_builder = [ordered]@{
            target_titles       = @($titles)
            important_skills    = @($skills)
            exclusion_keywords  = @($exclusions)
            search_queries      = @($queries)
            target_locations    = @($targetLocations)
            excluded_locations  = @($excludedLocations)
            excluded_contracts  = @($excludedContractsClean)
            employer_preference = $(if ([string]::IsNullOrWhiteSpace($EmployerPreference)) { "neutral" } else { $EmployerPreference })
        }
    }
    if ($Compact) {
        return $compactProfile
    }

    $apiQueries = @((@($titles) + @($skills | Select-Object -First 12)) | Select-Object -Unique)
    if ($apiQueries.Count -eq 0) {
        $apiQueries = @($profileLabel)
    }

    $titlePattern = ConvertTo-JobCrawlerRegexAlternation -Terms $titles
    $skillPattern = ConvertTo-JobCrawlerRegexAlternation -Terms $skills
    $exclusionPattern = ConvertTo-JobCrawlerRegexAlternation -Terms $exclusions
    $locationPattern = ConvertTo-JobCrawlerRegexAlternation -Terms $targetLocations
    $foreignLocationPatterns = @($excludedLocations | ForEach-Object { ConvertTo-JobCrawlerRegexAlternation -Terms @($_) } | Where-Object { $_ -ne "(?!)" })
    $contractPattern = Get-JobCrawlerContractPatternFromNames -ContractNames $excludedContractsClean

    $signals = New-Object System.Collections.Generic.List[object]
    foreach ($title in $titles) {
        $pattern = ConvertTo-JobCrawlerRegexAlternation -Terms @($title)
        $signals.Add((New-JobCrawlerProfileSignal -Scope "title" -Pattern $pattern -Keyword ("Role: {0}" -f $title) -Score 55)) | Out-Null
    }
    foreach ($skill in $skills) {
        $pattern = ConvertTo-JobCrawlerRegexAlternation -Terms @($skill)
        $signals.Add((New-JobCrawlerProfileSignal -Scope "full" -Pattern $pattern -Keyword ("Tool: {0}" -f $skill) -Score 25)) | Out-Null
    }

    $feedbackSignals = @($skills | ForEach-Object {
        [ordered]@{
            key     = ConvertTo-JobCrawlerProfileId $_
            label   = ("feedback positive: {0}" -f $_)
            pattern = ConvertTo-JobCrawlerRegexAlternation -Terms @($_)
        }
    })
    $feedbackNegativeRules = @()
    if ($exclusions.Count -gt 0) {
        $feedbackNegativeRules = @(
            [ordered]@{
                key     = "profile_exclusions"
                pattern = $exclusionPattern
                label   = "feedback ignored: profile exclusions"
                max     = 14
            }
        )
    }

    $employerWeights = [ordered]@{
        annonceur  = 0
        agency     = 0
        consulting = 0
        esn        = 0
        unknown    = 0
    }
    switch ((ConvertTo-JobCrawlerPlainText $EmployerPreference)) {
        "annonceur" {
            $employerWeights.annonceur = 10
            $employerWeights.agency = -8
            $employerWeights.consulting = -8
            $employerWeights.esn = -10
        }
        "agency_consulting_ok" {
            $employerWeights.agency = 4
            $employerWeights.consulting = 4
        }
    }

    return [ordered]@{
        id              = $compactProfile.id
        label           = $compactProfile.label
        description     = $compactProfile.description
        version         = $compactProfile.version
        profile_builder = $compactProfile.profile_builder
        sources         = [ordered]@{
            queries  = [ordered]@{
                linkedin       = @($queries)
                hellowork      = @($apiQueries)
                apec           = @($apiQueries)
                france_travail = @($apiQueries)
                adzuna         = @($apiQueries)
                api            = @($apiQueries)
            }
            patterns = [ordered]@{
                wttj_url_candidate = ("(?i)({0}|{1})" -f $titlePattern, $skillPattern)
            }
        }
        matching_rules  = [ordered]@{
            contract_rules           = [ordered]@{
                excluded_pattern                = $contractPattern
                early_explicit_excluded_pattern = $contractPattern
                preferred_pattern               = "\bcdi\b|permanent|full\s*time|temps\s+plein"
            }
            contexts                 = [ordered]@{
                core_title              = $titlePattern
                profile_skill_context   = $skillPattern
                profile_title_context   = $titlePattern
                marketing_only          = $exclusionPattern
                data_warehouse          = $exclusionPattern
                go_to_market            = "go\s*[- ]?\s*to\s*[- ]?\s*market"
            }
            positive_signals         = @($signals.ToArray())
            special_positive_signals = [ordered]@{}
            negative_signals         = [ordered]@{
                marketing_only            = [ordered]@{ keyword = "Risk: excluded keyword only"; score = -25 }
                marketing_related         = [ordered]@{ pattern = $exclusionPattern; keyword = "Risk: profile exclusion keyword"; score = -8 }
                data_analyst_engineering  = [ordered]@{ keyword = "Risk: profile exclusion context"; score = -25 }
                warehouse_python          = [ordered]@{ keyword = "Risk: profile exclusion context"; score = -12 }
                engineering               = [ordered]@{ pattern = $exclusionPattern; keyword = "Risk: profile exclusion keyword"; score = -10 }
            }
            feedback_positive_signals = @($feedbackSignals)
            feedback_negative_rules   = @($feedbackNegativeRules)
        }
        preferences     = [ordered]@{
            preferred_employer_type  = $(if ([string]::IsNullOrWhiteSpace($EmployerPreference)) { "neutral" } else { $EmployerPreference })
            employer_type_weights    = $employerWeights
            location_fit_weights     = [ordered]@{ target = 8; france_other = -4; foreign = -20; unknown = 0 }
            seniority_fit_weights    = [ordered]@{ target = 0; senior_ok = 0; too_junior = -12; too_managerial = -12; unknown = 0 }
            contract_fit_weights     = [ordered]@{ preferred = 5; excluded = -100; unknown = 0 }
            target_location_patterns = @($targetLocations | ForEach-Object { ConvertTo-JobCrawlerRegexAlternation -Terms @($_) })
            foreign_location_patterns = @($foreignLocationPatterns)
        }
        workbook        = [ordered]@{
            ignore_reason_options = @(
                "ignore_reason=not_relevant_enough; detail=",
                "ignore_reason=too_senior; detail=",
                "ignore_reason=too_junior; detail=",
                "ignore_reason=wrong_location; detail=",
                "ignore_reason=wrong_remote_policy; detail=",
                "ignore_reason=wrong_contract; detail=",
                "ignore_reason=company_not_interested; detail=",
                "ignore_reason=industry_not_interested; detail=",
                "ignore_reason=agency_consulting_esn; detail=",
                "ignore_reason=duplicate; company_alias=; detail=",
                "ignore_reason=low_quality_posting; detail=",
                "ignore_reason=other; detail="
            )
        }
    }
}

function Expand-JobCrawlerProfile {
    param([AllowNull()]$Profile)

    if ($null -eq $Profile) {
        return $null
    }

    $builder = Get-ConfigProperty -Object $Profile -Name "profile_builder" -DefaultValue $null
    if ($null -eq $builder) {
        return $Profile
    }

    $generated = New-JobCrawlerProfileFromBuilder `
        -Label ([string](Get-ConfigProperty -Object $Profile -Name "label" -DefaultValue "Custom profile")) `
        -Id ([string](Get-ConfigProperty -Object $Profile -Name "id" -DefaultValue "")) `
        -Description ([string](Get-ConfigProperty -Object $Profile -Name "description" -DefaultValue "")) `
        -TargetTitles (Get-ConfigStringArray (Get-ConfigProperty -Object $builder -Name "target_titles" -DefaultValue @())) `
        -ImportantSkills (Get-ConfigStringArray (Get-ConfigProperty -Object $builder -Name "important_skills" -DefaultValue @())) `
        -ExclusionKeywords (Get-ConfigStringArray (Get-ConfigProperty -Object $builder -Name "exclusion_keywords" -DefaultValue @())) `
        -SearchQueries (Get-ConfigStringArray (Get-ConfigProperty -Object $builder -Name "search_queries" -DefaultValue @())) `
        -TargetLocations (Get-ConfigStringArray (Get-ConfigProperty -Object $builder -Name "target_locations" -DefaultValue @())) `
        -ExcludedLocations (Get-ConfigStringArray (Get-ConfigProperty -Object $builder -Name "excluded_locations" -DefaultValue @())) `
        -ExcludedContracts (Get-ConfigStringArray (Get-ConfigProperty -Object $builder -Name "excluded_contracts" -DefaultValue @("CDD", "Apprenticeship", "Internship", "Freelance"))) `
        -EmployerPreference ([string](Get-ConfigProperty -Object $builder -Name "employer_preference" -DefaultValue "neutral"))

    return Merge-JobCrawlerConfigObjects -Base $generated -Override $Profile
}

function Save-JobCrawlerLocalProfile {
    param(
        [string]$ConfigDirectory,
        [AllowNull()]$Profile
    )

    if ($null -eq $Profile) {
        throw "Profile is required."
    }

    $configRoot = Resolve-Path -LiteralPath $ConfigDirectory -ErrorAction SilentlyContinue
    if ($null -eq $configRoot) {
        throw "Config directory not found: $ConfigDirectory"
    }

    $profileId = ConvertTo-JobCrawlerProfileId ([string](Get-ConfigProperty -Object $Profile -Name "id" -DefaultValue ""))
    if ([string]::IsNullOrWhiteSpace($profileId)) {
        throw "Profile id is required."
    }

    $path = Join-Path (Join-Path (Join-Path $configRoot.Path "local") "profiles") ("{0}.json" -f $profileId)
    Write-JobCrawlerJsonConfig -Path $path -Value $Profile
    return $path
}

function Clear-JobCrawlerDefaultProfile {
    param([string]$ConfigDirectory)

    $configRoot = Resolve-Path -LiteralPath $ConfigDirectory -ErrorAction SilentlyContinue
    if ($null -eq $configRoot) {
        throw "Config directory not found: $ConfigDirectory"
    }

    $path = Join-Path $configRoot.Path "local.runtime.json"
    $current = Read-JobCrawlerJsonConfig -Path $path -DefaultValue ([PSCustomObject]@{})
    $hash = ConvertTo-ConfigHashtable $current
    if ($null -eq $hash -or -not ($hash -is [System.Collections.IDictionary])) {
        $hash = [ordered]@{}
    }
    if (-not $hash.Contains("defaults") -or $null -eq $hash["defaults"] -or -not ($hash["defaults"] -is [System.Collections.IDictionary])) {
        $hash["defaults"] = [ordered]@{}
    }
    $hash["defaults"]["profile_id"] = ""
    Write-JobCrawlerJsonConfig -Path $path -Value $hash
    return $path
}

function Remove-JobCrawlerLocalProfile {
    param(
        [string]$ConfigDirectory,
        [string]$ProfileId
    )

    $configRoot = Resolve-Path -LiteralPath $ConfigDirectory -ErrorAction SilentlyContinue
    if ($null -eq $configRoot) {
        throw "Config directory not found: $ConfigDirectory"
    }

    $selectedId = ConvertTo-JobCrawlerProfileId $ProfileId
    if ([string]::IsNullOrWhiteSpace($selectedId)) {
        throw "Profile id is required."
    }

    $path = Join-Path (Join-Path (Join-Path $configRoot.Path "local") "profiles") ("{0}.json" -f $selectedId)
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Only local profiles can be deleted from the GUI. Local profile not found: $selectedId"
    }

    Remove-Item -LiteralPath $path -Force

    $runtimePath = Join-Path $configRoot.Path "local.runtime.json"
    $runtime = Read-JobCrawlerJsonConfig -Path $runtimePath -DefaultValue ([PSCustomObject]@{})
    $currentDefault = ConvertTo-JobCrawlerProfileId ([string](Get-ConfigPathValue -Object $runtime -Path "defaults.profile_id" -DefaultValue ""))
    if ($currentDefault -eq $selectedId) {
        [void](Clear-JobCrawlerDefaultProfile -ConfigDirectory $configRoot.Path)
    }

    return $path
}

function Set-JobCrawlerDefaultProfile {
    param(
        [string]$ConfigDirectory,
        [string]$ProfileId
    )

    $configRoot = Resolve-Path -LiteralPath $ConfigDirectory -ErrorAction SilentlyContinue
    if ($null -eq $configRoot) {
        throw "Config directory not found: $ConfigDirectory"
    }

    $selectedId = ConvertTo-JobCrawlerProfileId $ProfileId
    if ([string]::IsNullOrWhiteSpace($selectedId)) {
        throw "Profile id is required."
    }

    $path = Join-Path $configRoot.Path "local.runtime.json"
    $current = Read-JobCrawlerJsonConfig -Path $path -DefaultValue ([PSCustomObject]@{})
    $hash = ConvertTo-ConfigHashtable $current
    if ($null -eq $hash -or -not ($hash -is [System.Collections.IDictionary])) {
        $hash = [ordered]@{}
    }
    if (-not $hash.Contains("defaults") -or $null -eq $hash["defaults"] -or -not ($hash["defaults"] -is [System.Collections.IDictionary])) {
        $hash["defaults"] = [ordered]@{}
    }
    $hash["defaults"]["profile_id"] = $selectedId
    Write-JobCrawlerJsonConfig -Path $path -Value $hash
    return $path
}
