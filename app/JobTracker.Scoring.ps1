# Auto-extracted from Find-AnalyticsJobs.ps1. Keep dot-sourced execution order in the main script.

function Add-MatchSignal {
    param(
        [hashtable]$State,
        [string]$Text,
        [string]$Pattern,
        [string]$Keyword,
        [int]$Score
    )

    if ($Text -match $Pattern) {
        $State.Score += $Score
        $State.Keywords[$Keyword] = $true
    }
}

function Format-MatchReasonText {
    param([AllowNull()][string[]]$Reasons)

    $groups = [ordered]@{
        Role = New-Object System.Collections.Generic.List[string]
        Tool = New-Object System.Collections.Generic.List[string]
        Mission = New-Object System.Collections.Generic.List[string]
        Risk = New-Object System.Collections.Generic.List[string]
        Feedback = New-Object System.Collections.Generic.List[string]
        Other = New-Object System.Collections.Generic.List[string]
    }

    foreach ($reason in @($Reasons)) {
        if ([string]::IsNullOrWhiteSpace($reason)) {
            continue
        }

        $clean = [string]$reason
        if ($clean -match "^(?<group>Role|Tool|Mission|Risk|Feedback)\s*:\s*(?<value>.+)$") {
            $groups[$matches.Group].Add($matches.Value.Trim()) | Out-Null
        }
        elseif ($clean -match "^feedback\s*:?\s*(?<value>.+)$") {
            $groups.Feedback.Add($matches.Value.Trim()) | Out-Null
        }
        else {
            $groups.Other.Add($clean.Trim()) | Out-Null
        }
    }

    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($groupName in $groups.Keys) {
        $values = @($groups[$groupName].ToArray() | Sort-Object -Unique)
        if ($values.Count -gt 0) {
            $parts.Add(("{0}: {1}" -f $groupName, ($values -join ", "))) | Out-Null
        }
    }

    return ($parts.ToArray() -join " | ")
}

function Get-JobMatch {
    param(
        [string]$Title,
        [string]$Text
    )

    $titleText = ConvertTo-MatchText $Title
    $fullText = ConvertTo-MatchText ("{0} {1}" -f $Title, $Text)
    $rules = $script:JobCrawlerMatchingRules
    $coreTitlePattern = [string](Get-ConfigPathValue -Object $rules -Path "contexts.core_title" -DefaultValue "\bweb\s*analyst\b|\bdigital\s*analyst\b|analyste\s+(digital|web)|digital\s+analytics?\s+consultant|analytics?\s+consultant|web\s+analytics?|\bcro\b")
    $hasCoreTitleSignal = $titleText -match $coreTitlePattern
    $state = @{
        Score = 0
        Keywords = @{}
    }
    $isGoToMarketContext = $fullText -match [string](Get-ConfigPathValue -Object $rules -Path "contexts.go_to_market" -DefaultValue "go\s*[- ]?\s*to\s*[- ]?\s*market")
    $webAnalyticsToolPattern = [string](Get-ConfigPathValue -Object $rules -Path "contexts.web_analytics_tools" -DefaultValue "google\s+tag\s+manager|google\s+analytics|\bga4\b|piano\s+analytics|contentsquare|content\s+square")
    $hasWebAnalyticsToolSignal = ($fullText -match $webAnalyticsToolPattern) -or (-not $isGoToMarketContext -and $fullText -match "\bgtm\b")
    $hasDigitalAnalyticsContext = $hasWebAnalyticsToolSignal -or ($titleText -match [string](Get-ConfigPathValue -Object $rules -Path "contexts.digital_analytics_title" -DefaultValue "\bweb\s*analyst\b|\bdigital\s*analyst\b|web\s+analytics|digital\s+analytics|tracking|tagging|\bcro\b"))
    $isMarketingOnlyContext = ($fullText -match [string](Get-ConfigPathValue -Object $rules -Path "contexts.marketing_only" -DefaultValue "\bseo\b|\bsea\b|paid\s+social|paid\s+search|performance\s+marketing")) -and -not $hasDigitalAnalyticsContext
    $isDataWarehouseContext = $fullText -match [string](Get-ConfigPathValue -Object $rules -Path "contexts.data_warehouse" -DefaultValue "\bdbt\b|snowflake|airflow|\betl\b|data\s+warehouse|\bpython\b")

    foreach ($signal in @(Get-ConfigPathValue -Object $rules -Path "positive_signals" -DefaultValue @())) {
        $scope = [string](Get-ConfigProperty -Object $signal -Name "scope" -DefaultValue "full")
        $pattern = [string](Get-ConfigProperty -Object $signal -Name "pattern" -DefaultValue "")
        $keyword = [string](Get-ConfigProperty -Object $signal -Name "keyword" -DefaultValue "")
        $score = [int](Get-ConfigProperty -Object $signal -Name "score" -DefaultValue 0)
        if ([string]::IsNullOrWhiteSpace($pattern) -or [string]::IsNullOrWhiteSpace($keyword) -or $score -eq 0) {
            continue
        }
        $signalText = $(if ($scope -eq "title") { $titleText } else { $fullText })
        Add-MatchSignal $state $signalText $pattern $keyword $score
    }

    $gtmSignal = Get-ConfigPathValue -Object $rules -Path "special_positive_signals.gtm_when_not_go_to_market" -DefaultValue $null
    if ($null -ne $gtmSignal -and -not $isGoToMarketContext -and $fullText -match [string](Get-ConfigProperty -Object $gtmSignal -Name "pattern" -DefaultValue "\bgtm\b")) {
        $state.Score += [int](Get-ConfigProperty -Object $gtmSignal -Name "score" -DefaultValue 35)
        $state.Keywords[[string](Get-ConfigProperty -Object $gtmSignal -Name "keyword" -DefaultValue "Tool: GTM")] = $true
    }

    $negative = Get-ConfigProperty -Object $rules -Name "negative_signals" -DefaultValue $null
    if ($isMarketingOnlyContext) {
        $rule = Get-ConfigProperty -Object $negative -Name "marketing_only" -DefaultValue $null
        $state.Score += [int](Get-ConfigProperty -Object $rule -Name "score" -DefaultValue -25)
        $state.Keywords[[string](Get-ConfigProperty -Object $rule -Name "keyword" -DefaultValue "Risk: SEO/SEA/marketing-only role")] = $true
    }
    else {
        $rule = Get-ConfigProperty -Object $negative -Name "marketing_related" -DefaultValue $null
        if ($fullText -match [string](Get-ConfigProperty -Object $rule -Name "pattern" -DefaultValue "\bseo\b|\bsea\b|paid\s+social|performance\s+marketing|growth\s+marketing|digital\s+marketing")) {
            $state.Score += [int](Get-ConfigProperty -Object $rule -Name "score" -DefaultValue -8)
            $state.Keywords[[string](Get-ConfigProperty -Object $rule -Name "keyword" -DefaultValue "Risk: possible marketing role")] = $true
        }
    }

    foreach ($negativeName in @("time_tracking_payroll", "seo_only_title")) {
        $rule = Get-ConfigProperty -Object $negative -Name $negativeName -DefaultValue $null
        $pattern = [string](Get-ConfigProperty -Object $rule -Name "pattern" -DefaultValue "")
        $exceptPattern = [string](Get-ConfigProperty -Object $rule -Name "except_title_pattern" -DefaultValue "")
        if (-not [string]::IsNullOrWhiteSpace($pattern) -and $titleText -match $pattern -and ([string]::IsNullOrWhiteSpace($exceptPattern) -or $titleText -notmatch $exceptPattern)) {
            $state.Score += [int](Get-ConfigProperty -Object $rule -Name "score" -DefaultValue -60)
            $state.Keywords[[string](Get-ConfigProperty -Object $rule -Name "keyword" -DefaultValue ("Risk: {0}" -f $negativeName))] = $true
        }
    }

    $rule = Get-ConfigProperty -Object $negative -Name "hr_recruiting" -DefaultValue $null
    if ($titleText -match [string](Get-ConfigProperty -Object $rule -Name "pattern" -DefaultValue "human\s+resources|talent\s+acquisition|recruit") -and -not $hasCoreTitleSignal) {
        $state.Score += [int](Get-ConfigProperty -Object $rule -Name "score" -DefaultValue -60)
        $state.Keywords[[string](Get-ConfigProperty -Object $rule -Name "keyword" -DefaultValue "Risk: HR/recruiting role")] = $true
    }

    if ($isGoToMarketContext) {
        $rule = Get-ConfigProperty -Object $negative -Name "go_to_market" -DefaultValue $null
        $state.Score += [int](Get-ConfigProperty -Object $rule -Name "score" -DefaultValue -35)
        $state.Keywords[[string](Get-ConfigProperty -Object $rule -Name "keyword" -DefaultValue "Risk: go-to-market role")] = $true
    }

    $rule = Get-ConfigProperty -Object $negative -Name "broad_analyst" -DefaultValue $null
    if ($fullText -match [string](Get-ConfigProperty -Object $rule -Name "pattern" -DefaultValue "business\s+analyst|risk|finance|banking")) {
        $state.Score += [int](Get-ConfigProperty -Object $rule -Name "score" -DefaultValue -10)
        $state.Keywords[[string](Get-ConfigProperty -Object $rule -Name "keyword" -DefaultValue "Risk: broad analyst role")] = $true
    }

    if ((($titleText -match "\bdata\s*analyst\b|analyste\s+de\s+donnees|analytics?\s+engineer|data\s+engineer") -or $isDataWarehouseContext) -and -not $hasDigitalAnalyticsContext) {
        $rule = Get-ConfigProperty -Object $negative -Name "data_analyst_engineering" -DefaultValue $null
        $state.Score += [int](Get-ConfigProperty -Object $rule -Name "score" -DefaultValue -25)
        $state.Keywords[[string](Get-ConfigProperty -Object $rule -Name "keyword" -DefaultValue "Risk: data analyst/engineering role")] = $true
    }
    elseif ($isDataWarehouseContext) {
        $rule = Get-ConfigProperty -Object $negative -Name "warehouse_python" -DefaultValue $null
        $state.Score += [int](Get-ConfigProperty -Object $rule -Name "score" -DefaultValue -12)
        $state.Keywords[[string](Get-ConfigProperty -Object $rule -Name "keyword" -DefaultValue "Risk: warehouse/python role")] = $true
    }

    $rule = Get-ConfigProperty -Object $negative -Name "engineering" -DefaultValue $null
    if ($fullText -match [string](Get-ConfigProperty -Object $rule -Name "pattern" -DefaultValue "software\s+engineer|data\s+engineer|backend|frontend|devops")) {
        $state.Score += [int](Get-ConfigProperty -Object $rule -Name "score" -DefaultValue -15)
        $state.Keywords[[string](Get-ConfigProperty -Object $rule -Name "keyword" -DefaultValue "Risk: engineering role")] = $true
    }

    $learning = Get-FeedbackLearningAdjustment `
        -FullText $fullText `
        -HasCoreTitleSignal:$hasCoreTitleSignal `
        -HasWebAnalyticsToolSignal:$hasWebAnalyticsToolSignal `
        -HasDigitalAnalyticsContext:$hasDigitalAnalyticsContext
    if ($null -ne $learning -and [int]$learning.Adjustment -ne 0) {
        $state.Score += [int]$learning.Adjustment
        foreach ($reason in @($learning.Reasons)) {
            if (-not [string]::IsNullOrWhiteSpace($reason)) {
                $state.Keywords[$reason] = $true
            }
        }
    }

    if ($state.Score -lt 0) {
        $state.Score = 0
    }
    $noCoreTitleCap = [int](Get-ConfigPathValue -Object $rules -Path "thresholds.no_core_title_cap" -DefaultValue 49)
    if (-not $hasCoreTitleSignal -and $state.Score -gt $noCoreTitleCap) {
        $state.Score = $noCoreTitleCap
        $state.Keywords["Risk: no core title signal"] = $true
    }

    $level = "Review"
    if ($state.Score -ge [int](Get-ConfigPathValue -Object $rules -Path "thresholds.high_score" -DefaultValue 80)) {
        $level = "High"
    }
    elseif ($state.Score -ge [int](Get-ConfigPathValue -Object $rules -Path "thresholds.medium_score" -DefaultValue 50)) {
        $level = "Medium"
    }

    [PSCustomObject]@{
        IsMatch = $state.Score -ge $MinimumMatchScore
        Score = [int]$state.Score
        Level = $level
        Keywords = Format-MatchReasonText -Reasons @($state.Keywords.Keys)
    }
}

function Get-ContractTypeFromText {
    param(
        [AllowNull()][string]$Text,
        [AllowNull()][string]$RawContractType = $null
    )

    if (-not [string]::IsNullOrWhiteSpace($RawContractType)) {
        switch -Regex ($RawContractType) {
            "^FULL_TIME$" { return "CDI" }
            "^INTERNSHIP$" { return "Internship" }
            "^APPRENTICESHIP$" { return "Apprenticeship" }
            "^TEMPORARY$" { return "CDD" }
            "^FREELANCE$" { return "Freelance" }
            default { return $RawContractType }
        }
    }

    $matchText = ConvertTo-MatchText $Text
    if ($matchText -match "alternance|alternant|apprentissage|apprenticeship") {
        return "Apprenticeship"
    }
    if ($matchText -match "\bstage\b|stagiaire|internship|intern\b") {
        return "Internship"
    }
    if ($matchText -match "\bcdi\b|contrat\s+a\s+duree\s+indeterminee") {
        return "CDI"
    }
    if ($matchText -match "\bcdd\b|contrat\s+a\s+duree\s+determinee|fixed[-\s]*term|temporary|temporaire") {
        return "CDD"
    }
    if ($matchText -match "freelance|contractor|independant") {
        return "Freelance"
    }
    if ($matchText -match "employment\s+type\s+full-time|type\s+d.?emploi\s+temps\s+plein|\bfull-time\b|\btemps\s+plein\b") {
        return "Full-time"
    }
    if ($matchText -match "permanent\s+(contract|position|role)|\bpermanent\b") {
        return "Permanent"
    }

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ""
    }

    if ($Text -match "(?i)(\bCDI\b|contrat\s+a\s+duree\s+indeterminee)") {
        return "CDI"
    }
    if ($Text -match "(?i)(alternance|alternant|apprentissage|apprenticeship)") {
        return "Apprenticeship"
    }
    if ($Text -match "(?i)(stage|stagiaire|internship|intern\b)") {
        return "Internship"
    }
    if ($Text -match "(?i)(\bCDD\b|contrat\s+a\s+duree\s+determinee|fixed[-\s]*term|temporary|temporaire)") {
        return "CDD"
    }
    if ($Text -match "(?i)(freelance|contractor|independant)") {
        return "Freelance"
    }
    if ($Text -match "(?i)(Employment\s+type\s+Full-time|Type\s+d.?emploi\s+Temps\s+plein|\bFull-time\b|\bTemps\s+plein\b)") {
        return "Full-time"
    }
    if ($Text -match "(?i)(permanent\s+(contract|position|role)|\bpermanent\b)") {
        return "Permanent"
    }

    return ""
}

function Test-IsExcludedContractType {
    param([AllowNull()][string]$ContractType)

    $contractText = ConvertTo-MatchText $ContractType
    if ([string]::IsNullOrWhiteSpace($contractText)) {
        return $false
    }

    $pattern = [string](Get-ConfigPathValue -Object $script:JobCrawlerMatchingRules -Path "contract_rules.excluded_pattern" -DefaultValue "\bcdd\b|apprenticeship|apprentissage|alternance|internship|\bstage\b|stagiaire|temporary|fixed\s+term|freelance|contractor|independant|independent")
    return $contractText -match $pattern
}

function Get-EarlyContractType {
    param(
        [AllowNull()][string]$ContractType,
        [AllowNull()][string]$Text = ""
    )

    if (-not [string]::IsNullOrWhiteSpace($ContractType)) {
        return Get-ContractTypeFromText -Text $Text -RawContractType $ContractType
    }

    return Get-ContractTypeFromText -Text $Text
}

function Test-ShouldSkipEarlyByContract {
    param(
        [AllowNull()][string]$ContractType,
        [AllowNull()][string]$Text = "",
        [switch]$Reliable
    )

    $effectiveContract = Get-EarlyContractType -ContractType $ContractType -Text $Text
    if ([string]::IsNullOrWhiteSpace($effectiveContract)) {
        return $false
    }

    if ($Reliable) {
        return Test-IsExcludedContractType $effectiveContract
    }

    $matchText = ConvertTo-MatchText (Join-CleanTextParts @($ContractType, $Text))
    $earlyPattern = [string](Get-ConfigPathValue -Object $script:JobCrawlerMatchingRules -Path "contract_rules.early_explicit_excluded_pattern" -DefaultValue "\bcdd\b|apprentissage|alternance|apprenticeship|internship|\bstage\b|stagiaire|freelance|contractor|independant|independent")
    if ($matchText -match $earlyPattern) {
        return Test-IsExcludedContractType $effectiveContract
    }

    return $false
}

function Get-PreferenceObjectValue {
    param(
        [AllowNull()]$Object,
        [string]$Name,
        [AllowNull()]$DefaultValue = $null
    )

    if ($null -eq $Object) {
        return $DefaultValue
    }

    if ($Object -is [Collections.IDictionary] -and $Object.Contains($Name)) {
        return $Object[$Name]
    }

    $property = $Object.PSObject.Properties[$Name]
    if ($null -ne $property -and $null -ne $property.Value) {
        return $property.Value
    }

    return $DefaultValue
}

function New-DefaultJobCrawlerPreferences {
    return [PSCustomObject]@{
        preferred_employer_type = "annonceur"
        employer_type_weights = [PSCustomObject]@{
            annonceur = 10
            agency = -8
            consulting = -8
            esn = -10
            unknown = 0
        }
        location_fit_weights = [PSCustomObject]@{
            target = 8
            france_other = -4
            foreign = -20
            unknown = 0
        }
        seniority_fit_weights = [PSCustomObject]@{
            target = 0
            senior_ok = 0
            too_junior = -12
            too_managerial = -12
            unknown = 0
        }
        contract_fit_weights = [PSCustomObject]@{
            preferred = 5
            excluded = -100
            unknown = 0
        }
        target_location_patterns = @(
            "paris",
            "ile\s*de\s*france",
            "la\s+defense",
            "puteaux",
            "boulogne",
            "courbevoie",
            "nanterre",
            "clichy",
            "remote",
            "teletravail",
            "france"
        )
        foreign_location_patterns = @(
            "london",
            "\buk\b",
            "madrid",
            "barcelona",
            "casablanca",
            "montreal",
            "brussels",
            "belgium",
            "luxembourg",
            "switzerland",
            "geneva",
            "lausanne",
            "zurich",
            "cyprus",
            "canada",
            "morocco",
            "spain",
            "united\s+kingdom",
            "new\s*york",
            "united\s*states",
            "\bus\b",
            "u\.?s\.?a\.?",
            "berlin",
            "germany",
            "amsterdam",
            "netherlands",
            "dublin",
            "ireland",
            "milan",
            "italy",
            "singapore",
            "hong\s*kong"
        )
    }
}

function Get-JobCrawlerPreferences {
    $default = New-DefaultJobCrawlerPreferences
    if (Get-Variable -Name JobCrawlerConfig -Scope Script -ErrorAction SilentlyContinue) {
        $configuredPreferences = Get-ConfigProperty -Object $script:JobCrawlerConfig -Name "Preferences" -DefaultValue $null
        if ($null -ne $configuredPreferences -and @($configuredPreferences.PSObject.Properties).Count -gt 0) {
            return (Merge-JobCrawlerConfigObjects -Base $default -Override $configuredPreferences)
        }
    }

    $configRoot = ""
    if (Get-Variable -Name JobCrawlerConfig -Scope Script -ErrorAction SilentlyContinue) {
        $configRoot = [string]$script:JobCrawlerConfig.Root
    }
    if ([string]::IsNullOrWhiteSpace($configRoot)) {
        $configRoot = Join-Path (Split-Path -Parent $PSScriptRoot) "config"
    }
    $path = Join-Path $configRoot "preferences.json"
    if (-not (Test-Path -LiteralPath $path)) {
        return $default
    }

    try {
        $configuredPreferences = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
        return (Merge-JobCrawlerConfigObjects -Base $default -Override $configuredPreferences)
    }
    catch {
        Write-Warning ("Could not read config\preferences.json, using built-in defaults: {0}" -f $_.Exception.Message)
        return $default
    }
}

function Get-PreferenceWeight {
    param(
        [AllowNull()]$Preferences,
        [string]$GroupName,
        [string]$Key,
        [int]$DefaultValue = 0
    )

    $group = Get-PreferenceObjectValue -Object $Preferences -Name $GroupName -DefaultValue $null
    $rawValue = Get-PreferenceObjectValue -Object $group -Name $Key -DefaultValue $DefaultValue
    $number = 0
    if ([int]::TryParse([string]$rawValue, [ref]$number)) {
        return $number
    }

    return $DefaultValue
}

function Get-PreferenceArray {
    param(
        [AllowNull()]$Preferences,
        [string]$Name,
        [string[]]$DefaultValue = @()
    )

    $rawValue = Get-PreferenceObjectValue -Object $Preferences -Name $Name -DefaultValue $DefaultValue
    if ($null -eq $rawValue) {
        return @()
    }
    if ($rawValue -is [string]) {
        return @([string]$rawValue)
    }

    return @($rawValue)
}

function Get-JobCrawlerPreferenceArray {
    param(
        [AllowNull()]$Preferences,
        [string]$Name
    )

    if ($null -eq $Preferences) {
        $Preferences = New-DefaultJobCrawlerPreferences
    }
    $defaults = New-DefaultJobCrawlerPreferences
    $defaultValue = Get-PreferenceArray -Preferences $defaults -Name $Name -DefaultValue @()
    return Get-PreferenceArray -Preferences $Preferences -Name $Name -DefaultValue $defaultValue
}

function Test-AnyPatternMatch {
    param(
        [string]$Text,
        [object[]]$Patterns
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $false
    }

    foreach ($pattern in @($Patterns)) {
        if ([string]::IsNullOrWhiteSpace([string]$pattern)) {
            continue
        }

        try {
            if ($Text -match [string]$pattern) {
                return $true
            }
        }
        catch {
        }
    }

    return $false
}

function Get-EmployerType {
    param(
        [AllowNull()][string]$Title,
        [AllowNull()][string]$CompanyName,
        [AllowNull()][string]$Text = ""
    )

    $titleText = ConvertTo-MatchText $Title
    $companyText = ConvertTo-MatchText $CompanyName
    $combinedText = ConvertTo-MatchText (Join-CleanTextParts @($Title, $CompanyName, $Text))

    $patterns = Get-PreferenceObjectValue -Object $script:JobCrawlerPreferences -Name "employer_type_patterns" -DefaultValue $null
    $genericJobBoardPattern = [string](Get-PreferenceObjectValue -Object $patterns -Name "generic_job_board" -DefaultValue "confidential|jobgether|linkedin|indeed|talent\s*com|jobs?\s+via")
    if ([string]::IsNullOrWhiteSpace($companyText) -or $companyText -match $genericJobBoardPattern) {
        return "unknown"
    }

    $knownAgencyPattern = [string](Get-PreferenceObjectValue -Object $patterns -Name "known_agency" -DefaultValue "publicis|dentsu|havas|labelium|pixalione|eskimoz|jellyfish|performics|iprospect|allmatik|leonar|ekinox")
    $knownConsultingPattern = [string](Get-PreferenceObjectValue -Object $patterns -Name "known_consulting" -DefaultValue "fifty\s*[- ]?\s*five|\b55\b|converteo|artefact|optimal\s+ways|innoha|ekimetrics|wavestone|mc2i|deloitte|pwc|ey|kpmg|accenture")
    $knownEsnPattern = [string](Get-PreferenceObjectValue -Object $patterns -Name "known_esn" -DefaultValue "\bcgi\b|infotel|oventi|keyrus|micropole|business\s+&?\s+decision|devoteam|onepoint|talan|sopra\s+steria|capgemini|\bsqli\b|\bsqly\b|niji|consort|nexton|scalian|amaris|\bsii\b|atos|worldline|inetum|alten|ausy|neosoft")
    $agencyContextPattern = [string](Get-PreferenceObjectValue -Object $patterns -Name "agency_context" -DefaultValue "\bagence\b|agency|paid\s+media\s+agency|marketing\s+agency|media\s+agency")
    $consultingContextPattern = [string](Get-PreferenceObjectValue -Object $patterns -Name "consulting_context" -DefaultValue "cabinet\s+(de\s+)?conseil|societe\s+de\s+conseil|\bconseil\b|consulting\s+(firm|agency|company|cabinet)|missions?\s+chez\s+les?\s+clients|chez\s+nos\s+clients")
    $esnContextPattern = [string](Get-PreferenceObjectValue -Object $patterns -Name "esn_context" -DefaultValue "\besn\b|\bssii\b|services\s+numeriques|entreprise\s+de\s+services\s+du\s+numerique")

    if ($companyText -match $knownEsnPattern -or $combinedText -match $esnContextPattern) {
        return "esn"
    }
    if ($companyText -match $knownAgencyPattern -or $combinedText -match $agencyContextPattern) {
        return "agency"
    }
    if ($companyText -match $knownConsultingPattern -or $combinedText -match $consultingContextPattern) {
        return "consulting"
    }
    $consultantTitleContext = [string](Get-PreferenceObjectValue -Object $patterns -Name "consultant_title_context" -DefaultValue "\bconsultant(e)?\b")
    $consultantBodyContext = [string](Get-PreferenceObjectValue -Object $patterns -Name "consultant_body_context" -DefaultValue "client|mission|conseil|consulting|cabinet")
    if ($titleText -match $consultantTitleContext -and $combinedText -match $consultantBodyContext) {
        return "consulting"
    }

    return "annonceur"
}

function Get-LocationFitCategory {
    param(
        [AllowNull()][string]$Location,
        [AllowNull()]$Preferences
    )

    $locationText = ConvertTo-MatchText $Location
    if ([string]::IsNullOrWhiteSpace($locationText)) {
        return "unknown"
    }

    $targetPatterns = Get-JobCrawlerPreferenceArray -Preferences $Preferences -Name "target_location_patterns"
    $foreignPatterns = Get-JobCrawlerPreferenceArray -Preferences $Preferences -Name "foreign_location_patterns"

    if (Test-AnyPatternMatch -Text $locationText -Patterns $targetPatterns) {
        return "target"
    }
    if (Test-AnyPatternMatch -Text $locationText -Patterns $foreignPatterns) {
        return "foreign"
    }
    if ($locationText -match "\bfrance\b|paris|ile\s*de\s*france") {
        return "target"
    }

    return "france_other"
}

function Get-SeniorityFitCategory {
    param(
        [AllowNull()][string]$Title,
        [AllowNull()][string]$Text
    )

    $titleText = ConvertTo-MatchText $Title
    $fullText = ConvertTo-MatchText (Join-CleanTextParts @($Title, $Text))

    if ($titleText -match "\b(stage|stagiaire|intern|internship|apprentice|apprentissage|alternance|assistant|graduate|junior)\b") {
        return "too_junior"
    }
    if ($titleText -match "\b(head|director|directeur|directrice|lead|manager|responsable|principal)\b") {
        return "too_managerial"
    }
    if ($fullText -match "\b(stage|stagiaire|internship|apprentissage|alternance)\b") {
        return "too_junior"
    }
    if ($titleText -match "\b(senior|sr)\b") {
        return "senior_ok"
    }

    return "target"
}

function Get-ContractFitCategory {
    param([AllowNull()][string]$ContractType)

    $contractText = ConvertTo-MatchText $ContractType
    if ([string]::IsNullOrWhiteSpace($contractText)) {
        return "unknown"
    }
    if (Test-IsExcludedContractType $ContractType) {
        return "excluded"
    }
    $preferredPattern = [string](Get-ConfigPathValue -Object $script:JobCrawlerMatchingRules -Path "contract_rules.preferred_pattern" -DefaultValue "\bcdi\b|permanent|full\s*time|temps\s+plein")
    if ($contractText -match $preferredPattern) {
        return "preferred"
    }

    return "unknown"
}

function Get-JobFitDimensions {
    param(
        [int]$RoleScore,
        [AllowNull()][string]$Title,
        [AllowNull()][string]$CompanyName,
        [AllowNull()][string]$JobLocation,
        [AllowNull()][string]$ContractType,
        [AllowNull()][string]$Text = "",
        [AllowNull()]$Preferences = $JobCrawlerPreferences
    )

    if ($null -eq $Preferences) {
        $Preferences = New-DefaultJobCrawlerPreferences
    }

    $employerType = Get-EmployerType -Title $Title -CompanyName $CompanyName -Text $Text
    $locationCategory = Get-LocationFitCategory -Location $JobLocation -Preferences $Preferences
    $seniorityCategory = Get-SeniorityFitCategory -Title $Title -Text $Text
    $contractCategory = Get-ContractFitCategory $ContractType

    $employerFit = Get-PreferenceWeight -Preferences $Preferences -GroupName "employer_type_weights" -Key $employerType -DefaultValue 0
    $locationFit = Get-PreferenceWeight -Preferences $Preferences -GroupName "location_fit_weights" -Key $locationCategory -DefaultValue 0
    $seniorityFit = Get-PreferenceWeight -Preferences $Preferences -GroupName "seniority_fit_weights" -Key $seniorityCategory -DefaultValue 0
    $contractFit = Get-PreferenceWeight -Preferences $Preferences -GroupName "contract_fit_weights" -Key $contractCategory -DefaultValue 0

    $notes = New-Object System.Collections.Generic.List[string]
    $notes.Add(("role score {0}" -f $RoleScore)) | Out-Null
    if ($employerFit -ne 0) { $notes.Add(("employer {0}: {1}" -f $employerType, $employerFit)) | Out-Null }
    if ($locationFit -ne 0) { $notes.Add(("location {0}: {1}" -f $locationCategory, $locationFit)) | Out-Null }
    if ($seniorityFit -ne 0) { $notes.Add(("seniority {0}: {1}" -f $seniorityCategory, $seniorityFit)) | Out-Null }
    if ($contractFit -ne 0) { $notes.Add(("contract {0}: {1}" -f $contractCategory, $contractFit)) | Out-Null }

    $finalScore = [Math]::Max(0, $RoleScore + $employerFit + $locationFit + $seniorityFit + $contractFit)
    return [PSCustomObject]@{
        EmployerType      = $employerType
        RoleScore         = $RoleScore
        EmployerFit       = $employerFit
        LocationFit       = $locationFit
        SeniorityFit      = $seniorityFit
        ContractFit       = $contractFit
        LocationCategory  = $locationCategory
        SeniorityCategory = $seniorityCategory
        ContractCategory  = $contractCategory
        FinalScore        = [int]$finalScore
        MatchLevel        = Get-MatchLevelFromScore $finalScore
        Notes             = (($notes.ToArray() | Select-Object -Unique) -join "; ")
    }
}

function Test-IsAgencyConsultingEsnSignal {
    param(
        [AllowNull()][string]$Title,
        [AllowNull()][string]$CompanyName,
        [AllowNull()][string]$Text = ""
    )

    $employerType = Get-EmployerType -Title $Title -CompanyName $CompanyName -Text $Text
    return $employerType -in @("agency", "consulting", "esn")
}

function Test-IsAppliedStatus {
    param([AllowNull()][string]$Status)

    $statusText = ConvertTo-MatchText $Status
    return $statusText -match "^(applied|interview|offer|rejected|withdrawn)$"
}

function New-JobResult {
    param(
        [string]$Title,
        [string]$CompanyName,
        [string]$JobLocation,
        [string]$ContractType,
        [int]$MatchScore,
        [string]$MatchLevel,
        [string]$MatchedKeywords,
        [string]$Url,
        [string]$Platform,
        [AllowNull()]$PublishedAt,
        [AllowNull()][string]$SourceText = ""
    )

    if ([string]::IsNullOrWhiteSpace($Title) -or [string]::IsNullOrWhiteSpace($Url) -or $null -eq $PublishedAt) {
        return $null
    }

    $publishedDateValue = $null
    if ($PublishedAt -is [DateTimeOffset]) {
        $publishedDateValue = $PublishedAt
    }
    elseif ($PublishedAt -is [DateTime]) {
        $publishedDateValue = [DateTimeOffset]$PublishedAt
    }
    else {
        $publishedDateValue = ConvertTo-DateTimeOffsetOrNull ([string]$PublishedAt)
    }
    if ($null -eq $publishedDateValue) {
        return $null
    }

    $textContractType = Get-ContractTypeFromText -Text (Join-CleanTextParts @($Title, $SourceText))
    $effectiveContractType = $ContractType
    if ((Test-IsExcludedContractType $textContractType) -or [string]::IsNullOrWhiteSpace($effectiveContractType)) {
        $effectiveContractType = $textContractType
    }

    $key = "{0}|{1}" -f $Platform, $Url.ToLowerInvariant()
    if ($SeenResultKeys.ContainsKey($key)) {
        return $null
    }

    $identityKey = Get-JobIdentityKeyFromValues -Title $Title -CompanyName $CompanyName -JobLocation $JobLocation -Url $Url
    $jobId = Get-StableJobId $identityKey
    $fit = Get-JobFitDimensions -RoleScore $MatchScore -Title $Title -CompanyName $CompanyName -JobLocation $JobLocation -ContractType $effectiveContractType -Text $SourceText
    $adjustedScore = [int]$fit.FinalScore
    $adjustedKeywords = $MatchedKeywords.Trim()
    $fitKeywordNotes = New-Object System.Collections.Generic.List[string]
    if ([int]$fit.EmployerFit -lt 0) {
        $fitKeywordNotes.Add(("employer preference: {0}" -f $fit.EmployerType)) | Out-Null
    }
    if ([int]$fit.LocationFit -lt 0) {
        $fitKeywordNotes.Add(("location fit: {0}" -f $fit.LocationCategory)) | Out-Null
    }
    if ([int]$fit.SeniorityFit -lt 0) {
        $fitKeywordNotes.Add(("seniority fit: {0}" -f $fit.SeniorityCategory)) | Out-Null
    }
    if ([int]$fit.ContractFit -lt 0) {
        $fitKeywordNotes.Add(("contract fit: {0}" -f $fit.ContractCategory)) | Out-Null
    }
    if ($fitKeywordNotes.Count -gt 0) {
        $adjustedKeywords = (Join-CleanTextParts @($adjustedKeywords, (($fitKeywordNotes.ToArray()) -join "; "))) -replace ", ", "; "
    }

    $SeenResultKeys[$key] = $true
    [PSCustomObject]@{
        job_id         = $jobId
        job_title      = $Title.Trim()
        company_name   = $CompanyName.Trim()
        employer_type  = [string]$fit.EmployerType
        location       = $JobLocation.Trim()
        contract_type  = $effectiveContractType.Trim()
        match_score    = $adjustedScore
        match_level    = ([string]$fit.MatchLevel).Trim()
        matched_keywords = $adjustedKeywords
        role_score     = [string]$fit.RoleScore
        employer_fit   = [string]$fit.EmployerFit
        location_fit   = [string]$fit.LocationFit
        seniority_fit  = [string]$fit.SeniorityFit
        contract_fit   = [string]$fit.ContractFit
        fit_notes      = [string]$fit.Notes
        feedback_adjustment = ""
        job_url        = ConvertTo-ExcelHyperlinkFormula -Url $Url -Label "Open"
        job_url_raw    = $Url.Trim()
        platform       = $Platform
        source_count   = "1"
        alternate_urls = ""
        published_date = $publishedDateValue.ToString("yyyy-MM-dd")
    }
}

function Get-MatchLevelFromScore {
    param([int]$Score)

    if ($Score -ge 80) {
        return "High"
    }
    if ($Score -ge 50) {
        return "Medium"
    }

    return "Review"
}

function Get-FeedbackProfileText {
    param([AllowNull()]$Row)

    return ConvertTo-MatchText (Join-CleanTextParts @(
        (Get-RowValue -Row $Row -Name "job_title"),
        (Get-RowValue -Row $Row -Name "company_name"),
        (Get-RowValue -Row $Row -Name "location"),
        (Get-RowValue -Row $Row -Name "contract_type"),
        (Get-RowValue -Row $Row -Name "matched_keywords"),
        (Get-RowValue -Row $Row -Name "notes")
    ))
}

function Test-FeedbackRowHasAgencyConsultingEsnSignal {
    param([AllowNull()]$Row)

    return Test-IsAgencyConsultingEsnSignal `
        -Title (Get-RowValue -Row $Row -Name "job_title") `
        -CompanyName (Get-RowValue -Row $Row -Name "company_name") `
        -Text (Join-CleanTextParts @(
            (Get-RowValue -Row $Row -Name "matched_keywords"),
            (Get-RowValue -Row $Row -Name "notes")
        ))
}

function Get-FeedbackSeniorityBucket {
    param([string]$Text)

    if ($Text -match "\b(stage|intern|internship|apprentice|apprentissage|alternance|junior|graduate)\b") {
        return "junior"
    }
    if ($Text -match "\b(head|director|directeur|directrice|lead|manager|responsable|principal)\b") {
        return "management"
    }
    if ($Text -match "\b(senior|sr)\b") {
        return "senior"
    }

    return ""
}

function Test-FeedbackTextHasWebAnalyticsSignal {
    param([string]$Text)

    return $Text -match "web\s+analytics|digital\s+analytics|web\s*analyst|digital\s*analyst|tracking|tagging|taggage|webtracking|google\s+tag\s+manager|\bgtm\b|google\s+analytics|\bga4\b|piano|contentsquare|content\s+square|tag\s+commander|commanders?\s+act|\btealium\b|data\s*layer|datalayer|tagging\s+plan|tracking\s+plan|plan\s+de\s+(taggage|marquage)|server\s*[- ]?\s*side|consent\s+mode|\brgpd\b|\bgdpr\b|matomo|adobe\s+analytics"
}

function Get-FeedbackSignalDefinitions {
    $configured = @(Get-ConfigPathValue -Object $script:JobCrawlerMatchingRules -Path "feedback_positive_signals" -DefaultValue @())
    if ($configured.Count -gt 0) {
        return @($configured | ForEach-Object {
            [PSCustomObject]@{
                Key     = [string](Get-ConfigProperty -Object $_ -Name "key" -DefaultValue "")
                Label   = [string](Get-ConfigProperty -Object $_ -Name "label" -DefaultValue "")
                Pattern = [string](Get-ConfigProperty -Object $_ -Name "pattern" -DefaultValue "")
            }
        } | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Key) -and -not [string]::IsNullOrWhiteSpace($_.Pattern) })
    }

    return @(
        [PSCustomObject]@{ Key = "google_tag_manager"; Label = "feedback positive: Google Tag Manager"; Pattern = "google\s+tag\s+manager|\bgtm\b" },
        [PSCustomObject]@{ Key = "google_analytics"; Label = "feedback positive: Google Analytics/GA4"; Pattern = "google\s+analytics|\bga4\b" },
        [PSCustomObject]@{ Key = "piano"; Label = "feedback positive: Piano"; Pattern = "piano" },
        [PSCustomObject]@{ Key = "contentsquare"; Label = "feedback positive: ContentSquare"; Pattern = "contentsquare|content\s+square" },
        [PSCustomObject]@{ Key = "tag_commander"; Label = "feedback positive: Tag Commander/Commanders Act"; Pattern = "tag\s+commander|commanders?\s+act" },
        [PSCustomObject]@{ Key = "tealium"; Label = "feedback positive: Tealium"; Pattern = "\btealium\b|tealium\s+iq" },
        [PSCustomObject]@{ Key = "server_side"; Label = "feedback positive: server-side tracking"; Pattern = "server\s*[- ]?\s*side|server\s+container|\bsgtm\b" },
        [PSCustomObject]@{ Key = "rgpd"; Label = "feedback positive: RGPD/GDPR"; Pattern = "\brgpd\b|\bgdpr\b|protection\s+des\s+donn[eé]es|privacy|conformit[eé]" },
        [PSCustomObject]@{ Key = "datalayer"; Label = "feedback positive: dataLayer"; Pattern = "data\s*layer|datalayer" },
        [PSCustomObject]@{ Key = "tagging_plan"; Label = "feedback positive: tagging plan"; Pattern = "tagging\s+plan|tracking\s+plan|plan\s+de\s+(taggage|marquage)" },
        [PSCustomObject]@{ Key = "consent"; Label = "feedback positive: consent tracking"; Pattern = "consent\s+mode|cookie\s+consent|\bcmp\b" },
        [PSCustomObject]@{ Key = "cro"; Label = "feedback positive: CRO"; Pattern = "\bcro\b|conversion\s+rate|conversion\s+optimization|optimisation\s+conversion" }
    )
}

function Get-HashtableIntValue {
    param(
        [AllowNull()]$Table,
        [string]$Key
    )

    if ($null -eq $Table -or [string]::IsNullOrWhiteSpace($Key)) {
        return 0
    }

    if ($Table -is [Collections.IDictionary] -and $Table.Contains($Key)) {
        return [int]$Table[$Key]
    }

    $property = $Table.PSObject.Properties[$Key]
    if ($null -ne $property -and $null -ne $property.Value) {
        return [int]$property.Value
    }

    return 0
}

function Add-FeedbackCount {
    param(
        [hashtable]$Table,
        [string]$Key
    )

    if ([string]::IsNullOrWhiteSpace($Key)) {
        return
    }

    if (-not $Table.ContainsKey($Key)) {
        $Table[$Key] = 0
    }
    $Table[$Key] = [int]$Table[$Key] + 1
}

function New-FeedbackLearningProfile {
    param([object[]]$Rows)

    $positiveCounts = @{}
    $ignoreReasonCounts = @{}
    $positiveRows = 0
    $ignoredRows = 0
    $signals = @(Get-FeedbackSignalDefinitions)

    foreach ($row in @($Rows)) {
        $status = ConvertTo-MatchText (Get-RowValue -Row $row -Name "status")
        if ([string]::IsNullOrWhiteSpace($status)) {
            continue
        }

        $rowText = Get-FeedbackProfileText $row
        if ($status -match "^(applied|interview|offer|interesting)$") {
            $positiveRows++
            foreach ($signal in $signals) {
                if ($rowText -match [string]$signal.Pattern) {
                    Add-FeedbackCount -Table $positiveCounts -Key ([string]$signal.Key)
                }
            }
        }
        elseif ($status -eq "ignored") {
            $ignoredRows++
            $ignoreReason = Get-IgnoreReasonFromNotes (Get-RowValue -Row $row -Name "notes")
            if (-not [string]::IsNullOrWhiteSpace($ignoreReason)) {
                Add-FeedbackCount -Table $ignoreReasonCounts -Key (ConvertTo-IgnoreReasonKey $ignoreReason)
            }
        }
    }

    return [PSCustomObject]@{
        PositiveSignalCounts = $positiveCounts
        IgnoreReasonCounts   = $ignoreReasonCounts
        PositiveRows         = $positiveRows
        IgnoredRows          = $ignoredRows
    }
}

function Get-FeedbackLearningAdjustment {
    param(
        [string]$FullText,
        [bool]$HasCoreTitleSignal,
        [bool]$HasWebAnalyticsToolSignal,
        [bool]$HasDigitalAnalyticsContext
    )

    $profile = $script:FeedbackLearningProfile
    if ($null -eq $profile) {
        return [PSCustomObject]@{ Adjustment = 0; Reasons = @() }
    }

    $adjustment = 0
    $reasons = New-Object System.Collections.Generic.List[string]
    $positiveSignals = $profile.PositiveSignalCounts
    foreach ($signal in @(Get-FeedbackSignalDefinitions)) {
        $count = Get-HashtableIntValue -Table $positiveSignals -Key ([string]$signal.Key)
        if ($count -le 0 -or $FullText -notmatch [string]$signal.Pattern) {
            continue
        }

        $delta = [Math]::Min(8, 2 + (2 * $count))
        $adjustment += $delta
        $reasons.Add([string]$signal.Label) | Out-Null
    }

    if ($adjustment -gt 18) {
        $adjustment = 18
    }

    $ignoreCounts = $profile.IgnoreReasonCounts
    $negativeAdjustment = 0
    $configuredNegativeRules = @(Get-ConfigPathValue -Object $script:JobCrawlerMatchingRules -Path "feedback_negative_rules" -DefaultValue @())
    if ($configuredNegativeRules.Count -gt 0) {
        $negativeRules = @($configuredNegativeRules | ForEach-Object {
            [PSCustomObject]@{
                Key     = [string](Get-ConfigProperty -Object $_ -Name "key" -DefaultValue "")
                Pattern = [string](Get-ConfigProperty -Object $_ -Name "pattern" -DefaultValue "")
                Label   = [string](Get-ConfigProperty -Object $_ -Name "label" -DefaultValue "")
                Max     = [int](Get-ConfigProperty -Object $_ -Name "max" -DefaultValue 8)
            }
        } | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Key) -and -not [string]::IsNullOrWhiteSpace($_.Pattern) })
    }
    else {
        $negativeRules = @(
            [PSCustomObject]@{ Key = "too_seo_sea_marketing"; Pattern = "\bseo\b|\bsea\b|paid\s+social|paid\s+search|paid\s+media|performance\s+marketing|growth\s+marketing|acquisition|campaign|media\s+buyer"; Label = "feedback ignored: SEO/SEA/marketing"; Max = 14 },
            [PSCustomObject]@{ Key = "too_data_analyst"; Pattern = "\bdata\s*analyst\b|analyste\s+de\s+donnees|\bpython\b|\bsql\b|notebook|data\s+warehouse|business\s+analyst"; Label = "feedback ignored: data analyst"; Max = 12 },
            [PSCustomObject]@{ Key = "too_data_engineering"; Pattern = "data\s+engineer|analytics?\s+engineer|\bdbt\b|snowflake|airflow|\betl\b|\belt\b|data\s+warehouse|datawarehouse|data\s+platform|databricks|pyspark|spark|pipeline|backend|devops"; Label = "feedback ignored: data engineering"; Max = 16 },
            [PSCustomObject]@{ Key = "too_bi_reporting"; Pattern = "\bbi\b|business\s+intelligence|power\s*bi|tableau|dashboard|reporting|looker|data\s+studio|tableau\s+de\s+bord"; Label = "feedback ignored: BI/reporting"; Max = 10 },
            [PSCustomObject]@{ Key = "too_crm_emailing"; Pattern = "\bcrm\b|emailing|email\s+marketing|marketing\s+automation|salesforce|hubspot|braze|batch|campaign"; Label = "feedback ignored: CRM/emailing"; Max = 12 },
            [PSCustomObject]@{ Key = "too_content_social"; Pattern = "content\s+marketing|social\s+media|community\s+manager|editorial|copywriting|seo\s+content"; Label = "feedback ignored: content/social"; Max = 12 },
            [PSCustomObject]@{ Key = "too_product_analytics"; Pattern = "product\s+analyst|product\s+analytics|amplitude|mixpanel|heap"; Label = "feedback ignored: product analytics"; Max = 8 },
            [PSCustomObject]@{ Key = "too_managerial"; Pattern = "\bhead\b|director|directeur|directrice|lead|manager|responsable|principal"; Label = "feedback ignored: managerial"; Max = 8 },
            [PSCustomObject]@{ Key = "agency_consulting_esn"; Pattern = "consultant|consulting|cabinet|agence|agency|\besn\b|ssii"; Label = "feedback ignored: agency/consulting/ESN"; Max = 8 }
        )
    }

    foreach ($rule in $negativeRules) {
        $count = Get-HashtableIntValue -Table $ignoreCounts -Key ([string]$rule.Key)
        if ($count -le 0 -or $FullText -notmatch [string]$rule.Pattern) {
            continue
        }

        if ($rule.Key -match "too_data|too_bi|too_product" -and $HasWebAnalyticsToolSignal) {
            continue
        }
        if ($rule.Key -eq "too_seo_sea_marketing" -and $HasDigitalAnalyticsContext) {
            continue
        }

        $delta = [Math]::Min([int]$rule.Max, 4 + (3 * $count))
        $negativeAdjustment -= $delta
        $reasons.Add([string]$rule.Label) | Out-Null
    }

    $notAnalyticsCount = Get-HashtableIntValue -Table $ignoreCounts -Key "not_analytics_enough"
    if ($notAnalyticsCount -gt 0 -and -not $HasCoreTitleSignal -and -not $HasWebAnalyticsToolSignal) {
        $negativeAdjustment -= [Math]::Min(12, 4 + (3 * $notAnalyticsCount))
        $reasons.Add("feedback ignored: not analytics enough") | Out-Null
    }

    if ($negativeAdjustment -lt -25) {
        $negativeAdjustment = -25
    }

    return [PSCustomObject]@{
        Adjustment = [int]($adjustment + $negativeAdjustment)
        Reasons    = @($reasons.ToArray() | Select-Object -Unique)
    }
}

function Get-IgnoredFeedbackPenalty {
    param(
        $Row,
        $ExistingRow,
        [AllowNull()][string]$IgnoreReason,
        [bool]$SameCompany,
        [bool]$SameTitle,
        [bool]$KeywordOverlap
    )

    $reason = ConvertTo-IgnoreReasonKey $IgnoreReason
    $rowText = Get-FeedbackProfileText $Row
    $existingText = Get-FeedbackProfileText $ExistingRow
    $hasWebAnalyticsSignal = Test-FeedbackTextHasWebAnalyticsSignal $rowText

    if ($reason -eq "duplicate") {
        return [PSCustomObject]@{ Penalty = 0; Reason = "" }
    }

    switch ($reason) {
        "not_analytics_enough" {
            if (-not $hasWebAnalyticsSignal -or $rowText -match "possible marketing|possible broad analyst|possible data analyst") {
                return [PSCustomObject]@{ Penalty = 22; Reason = "ignored reason: not analytics enough" }
            }
        }
        "too_seo_sea_marketing" {
            if ($rowText -match "\bseo\b|\bsea\b|paid\s+social|paid\s+search|paid\s+media|performance\s+marketing|growth\s+marketing|acquisition|digital\s+marketing|campaign|media\s+buyer") {
                return [PSCustomObject]@{ Penalty = 26; Reason = "ignored reason: SEO/SEA/marketing" }
            }
        }
        "too_data_analyst" {
            if (($rowText -match "\bdata\s*analyst\b|analyste\s+de\s+donnees|\bpython\b|\bsql\b|notebook|data\s+warehouse|possible broad analyst") -and -not $hasWebAnalyticsSignal) {
                return [PSCustomObject]@{ Penalty = 24; Reason = "ignored reason: data analyst" }
            }
        }
        "too_data_engineering" {
            if ($rowText -match "data\s+engineer|analytics?\s+engineer|\bdbt\b|snowflake|airflow|\betl\b|\belt\b|data\s+warehouse|datawarehouse|data\s+platform|databricks|pyspark|spark|pipeline|backend|devops|possible engineering") {
                return [PSCustomObject]@{ Penalty = 28; Reason = "ignored reason: data engineering" }
            }
        }
        "too_bi_reporting" {
            if (($rowText -match "\bbi\b|business\s+intelligence|power\s*bi|tableau|dashboard|reporting|looker|data\s+studio|tableau\s+de\s+bord") -and -not $hasWebAnalyticsSignal) {
                return [PSCustomObject]@{ Penalty = 20; Reason = "ignored reason: BI/reporting" }
            }
        }
        "too_crm_emailing" {
            if ($rowText -match "\bcrm\b|emailing|email\s+marketing|marketing\s+automation|salesforce|hubspot|braze|batch|campaign") {
                return [PSCustomObject]@{ Penalty = 22; Reason = "ignored reason: CRM/emailing" }
            }
        }
        "too_content_social" {
            if ($rowText -match "content\s+marketing|social\s+media|community\s+manager|editorial|copywriting|seo\s+content") {
                return [PSCustomObject]@{ Penalty = 22; Reason = "ignored reason: content/social" }
            }
        }
        "too_product_analytics" {
            if (($rowText -match "product\s+analyst|product\s+analytics|amplitude|mixpanel|heap") -and -not $hasWebAnalyticsSignal) {
                return [PSCustomObject]@{ Penalty = 16; Reason = "ignored reason: product analytics" }
            }
        }
        "too_managerial" {
            $rowBucket = Get-FeedbackSeniorityBucket $rowText
            if ($rowBucket -eq "management") {
                return [PSCustomObject]@{ Penalty = 16; Reason = "ignored reason: too managerial" }
            }
        }
        "agency_consulting_esn" {
            if ((Test-FeedbackRowHasAgencyConsultingEsnSignal $Row) -or $SameCompany -or $SameTitle) {
                return [PSCustomObject]@{ Penalty = 18; Reason = "ignored reason: agency/consulting/ESN preference" }
            }
        }
        "wrong_seniority" {
            $rowBucket = Get-FeedbackSeniorityBucket $rowText
            $existingBucket = Get-FeedbackSeniorityBucket $existingText
            if (-not [string]::IsNullOrWhiteSpace($rowBucket) -and $rowBucket -eq $existingBucket) {
                return [PSCustomObject]@{ Penalty = 14; Reason = "ignored reason: seniority" }
            }
        }
        "wrong_location" {
            $rowLocation = ConvertTo-IdentityText -Text (Get-RowValue -Row $Row -Name "location")
            $existingLocation = ConvertTo-IdentityText -Text (Get-RowValue -Row $ExistingRow -Name "location")
            if (-not [string]::IsNullOrWhiteSpace($rowLocation) -and $rowLocation -eq $existingLocation) {
                return [PSCustomObject]@{ Penalty = 12; Reason = "ignored reason: location" }
            }
        }
        "wrong_remote_policy" {
            if ($rowText -match "on\s*site|onsite|hybrid|remote|teletravail") {
                return [PSCustomObject]@{ Penalty = 8; Reason = "ignored reason: remote policy" }
            }
        }
        "wrong_contract" {
            $rowContract = ConvertTo-IdentityText -Text (Get-RowValue -Row $Row -Name "contract_type")
            $existingContract = ConvertTo-IdentityText -Text (Get-RowValue -Row $ExistingRow -Name "contract_type")
            if (-not [string]::IsNullOrWhiteSpace($rowContract) -and $rowContract -eq $existingContract) {
                return [PSCustomObject]@{ Penalty = 10; Reason = "ignored reason: contract" }
            }
        }
        "language_issue" {
            if ($rowText -match "english|french|francais|bilingual|native|fluent") {
                return [PSCustomObject]@{ Penalty = 8; Reason = "ignored reason: language" }
            }
        }
        "salary_issue" {
            return [PSCustomObject]@{ Penalty = 0; Reason = "" }
        }
        "company_not_interested" {
            if ($SameCompany) {
                return [PSCustomObject]@{ Penalty = 28; Reason = "ignored reason: company" }
            }
        }
        "industry_not_interested" {
            if ($SameCompany) {
                return [PSCustomObject]@{ Penalty = 12; Reason = "ignored reason: industry/company proxy" }
            }
        }
        "low_quality_posting" {
            if ($SameCompany -or $SameTitle) {
                return [PSCustomObject]@{ Penalty = 10; Reason = "ignored reason: low-quality posting" }
            }
        }
        "other" {
            if ($SameCompany -or $SameTitle) {
                return [PSCustomObject]@{ Penalty = 8; Reason = "ignored reason: other" }
            }
        }
    }

    if ($SameCompany -or $SameTitle -or $KeywordOverlap) {
        return [PSCustomObject]@{ Penalty = 12; Reason = $(if ([string]::IsNullOrWhiteSpace($reason)) { "similar ignored job without reason" } else { "similar ignored job" }) }
    }

    return [PSCustomObject]@{ Penalty = 0; Reason = "" }
}

function Get-FeedbackAdjustment {
    param(
        $Row,
        [object[]]$ExistingRows
    )

    $titleText = ConvertTo-IdentityText -Text (Get-RowValue -Row $Row -Name "job_title") -Title
    $companyText = ConvertTo-IdentityText -Text (Get-RowValue -Row $Row -Name "company_name")
    $keywordText = ConvertTo-MatchText (Get-RowValue -Row $Row -Name "matched_keywords")
    $adjustment = 0
    $reasons = New-Object System.Collections.Generic.List[string]

    foreach ($existing in @($ExistingRows)) {
        $status = ConvertTo-MatchText (Get-RowValue -Row $existing -Name "status")
        if ([string]::IsNullOrWhiteSpace($status)) {
            continue
        }

        $existingTitle = ConvertTo-IdentityText -Text (Get-RowValue -Row $existing -Name "job_title") -Title
        $existingCompany = ConvertTo-IdentityText -Text (Get-RowValue -Row $existing -Name "company_name")
        $existingKeywords = ConvertTo-MatchText (Get-RowValue -Row $existing -Name "matched_keywords")
        $sameCompany = -not [string]::IsNullOrWhiteSpace($companyText) -and $companyText -eq $existingCompany
        $sameTitle = -not [string]::IsNullOrWhiteSpace($titleText) -and $titleText -eq $existingTitle
        $keywordOverlap = -not [string]::IsNullOrWhiteSpace($keywordText) -and -not [string]::IsNullOrWhiteSpace($existingKeywords) -and ($keywordText -match "google|gtm|ga4|piano|contentsquare|tag\s+commander|commanders?\s+act|tealium|server-side|server\s+side|rgpd|gdpr|tracking|tagging|cro") -and ($existingKeywords -match "google|gtm|ga4|piano|contentsquare|tag\s+commander|commanders?\s+act|tealium|server-side|server\s+side|rgpd|gdpr|tracking|tagging|cro")

        if ($status -match "^(applied|interview|offer|interesting)$" -and ($sameCompany -or $sameTitle -or $keywordOverlap)) {
            $adjustment += 10
            $reasons.Add("positive history") | Out-Null
        }
        elseif ($status -eq "ignored") {
            $ignoreReason = Get-IgnoreReasonFromNotes (Get-RowValue -Row $existing -Name "notes")
            $ignoredFeedback = Get-IgnoredFeedbackPenalty -Row $Row -ExistingRow $existing -IgnoreReason $ignoreReason -SameCompany:$sameCompany -SameTitle:$sameTitle -KeywordOverlap:$keywordOverlap
            if ([int]$ignoredFeedback.Penalty -gt 0) {
                $adjustment -= [int]$ignoredFeedback.Penalty
                $reasons.Add([string]$ignoredFeedback.Reason) | Out-Null
            }
        }
    }

    if ($adjustment -gt 30) { $adjustment = 30 }
    if ($adjustment -lt -40) { $adjustment = -40 }

    return [PSCustomObject]@{
        Adjustment = $adjustment
        Reason = (($reasons.ToArray() | Select-Object -Unique) -join "; ")
    }
}

function Apply-FeedbackScoring {
    param(
        [object[]]$Rows,
        [object[]]$ExistingRows
    )

    foreach ($row in @($Rows)) {
        $feedback = Get-FeedbackAdjustment -Row $row -ExistingRows $ExistingRows
        $oldScore = 0
        try { $oldScore = [int](Get-RowValue -Row $row -Name "match_score") } catch { $oldScore = 0 }
        $newScore = [Math]::Max(0, $oldScore + [int]$feedback.Adjustment)
        $row.match_score = [string]$newScore
        $row.feedback_adjustment = [string]$feedback.Adjustment
        if (-not [string]::IsNullOrWhiteSpace($feedback.Reason)) {
            $row.matched_keywords = (Join-CleanTextParts @((Get-RowValue -Row $row -Name "matched_keywords"), ("feedback: " + $feedback.Reason))) -replace ", ", "; "
        }

        if ($newScore -ge 80) {
            $row.match_level = "High"
        }
        elseif ($newScore -ge 50) {
            $row.match_level = "Medium"
        }
        else {
            $row.match_level = "Review"
        }
    }

    return @($Rows)
}

