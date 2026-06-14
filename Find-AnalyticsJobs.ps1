[CmdletBinding()]
param(
    [int]$DaysBack = 7,
    [string]$Location = "France",
    [string]$TrackerPath = "",
    [string]$WelcomeKitApiKey = $env:WK_API_KEY,
    [int]$MaxLinkedInSearchPages = 3,
    [int]$MaxWelcomeKitPages = 10,
    [int]$MaxWttjCandidatePages = 120,
    [int]$MaxBackups = 5,
    [switch]$SkipLinkedIn,
    [switch]$SkipWttj,
    [switch]$DisableWttjPublicFallback
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "JobTracker.Common.ps1")

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$BrowserUserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
$Cutoff = [DateTimeOffset]::Now.AddDays(-[Math]::Abs($DaysBack))
$CutoffDate = $Cutoff.ToString("yyyy-MM-dd")
$RunDate = Get-Date -Format "yyyy-MM-dd"
$RunStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$DefaultTrackerPath = Join-Path $PSScriptRoot "output\jobs_tracker.xlsx"

if ([string]::IsNullOrWhiteSpace($TrackerPath)) {
    $TrackerPath = $DefaultTrackerPath
}
if ([IO.Path]::GetExtension($TrackerPath).ToLowerInvariant() -ne ".xlsx") {
    throw "This crawler uses only the XLSX tracker file. Use output\jobs_tracker.xlsx for -TrackerPath."
}

$SeenResultKeys = @{}
$LinkedInDelayMilliseconds = 1200
$MinimumMatchScore = 35

$WttjUrlCandidatePattern = "(?i)(web[-_\s]*analyst|digital[-_\s]*analyst|web[-_\s]*analytics|digital[-_\s]*analytics|analytics[-_\s]*(consultant|specialist|manager|engineer)|tracking|taggage|tagging|data[-_\s]*analyst|(^|[-_\s])ga4($|[-_\s])|(^|[-_\s])gtm($|[-_\s])|google[-_\s]*(analytics|tag[-_\s]*manager)|piano|contentsquare|content[-_\s]*square|(^|[-_\s])cro($|[-_\s]))"

$LinkedInQueries = @(
    "web analyst google tag manager",
    "web analyst google analytics",
    "web analyst cro google analytics",
    "digital analyst google tag manager",
    "digital analyst google analytics",
    "digital analytics consultant",
    "senior digital analytics consultant",
    "tracking analyst ga4 gtm",
    "tracking analytics specialist",
    "analytics consultant google tag manager",
    "performance digital google analytics",
    "chargé performance digital google analytics",
    "piano analytics",
    "contentsquare analytics",
    "tagging plan analytics"
)

function Set-RunWindowTitle {
    param([string]$Title)

    try {
        $Host.UI.RawUI.WindowTitle = $Title
    }
    catch {
    }
}

function Write-RunStatus {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "HH:mm:ss"
    Write-Host ("[{0}] [{1}] {2}" -f $timestamp, $Level, $Message)
}

function Write-CountProgress {
    param(
        [string]$Activity,
        [int]$Current,
        [int]$Total,
        [int]$Found = -1,
        [int]$Every = 10
    )

    if ($Total -le 0) {
        return
    }

    if ($Current -ne 1 -and $Current -ne $Total -and ($Current % $Every) -ne 0) {
        return
    }

    $percent = [int](($Current / [Math]::Max(1, $Total)) * 100)
    $foundText = ""
    if ($Found -ge 0) {
        $foundText = "; {0} matches so far" -f $Found
    }

    Write-RunStatus ("{0}: {1}/{2} ({3}%){4}" -f $Activity, $Current, $Total, $percent, $foundText)
}

function ConvertTo-QueryString {
    param([hashtable]$Params)

    ($Params.GetEnumerator() | ForEach-Object {
        "{0}={1}" -f [Uri]::EscapeDataString([string]$_.Key), [Uri]::EscapeDataString([string]$_.Value)
    }) -join "&"
}

function ConvertFrom-HtmlText {
    param([AllowNull()][string]$Html)

    if ([string]::IsNullOrWhiteSpace($Html)) {
        return ""
    }

    $text = [regex]::Replace($Html, "(?is)<script\b.*?</script>|<style\b.*?</style>", " ")
    $text = [regex]::Replace($text, "(?is)<br\s*/?>|</p>|</li>|</div>|</h\d>", " ")
    $text = [regex]::Replace($text, "(?is)<[^>]+>", " ")
    $text = [System.Net.WebUtility]::HtmlDecode($text)
    return ([regex]::Replace($text, "\s+", " ")).Trim()
}

function ConvertFrom-HtmlAttribute {
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ""
    }

    return ([System.Net.WebUtility]::HtmlDecode($Value)).Trim()
}

function ConvertTo-MatchText {
    param([AllowNull()][string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ""
    }

    $decoded = [System.Net.WebUtility]::HtmlDecode($Text)
    $normalized = $decoded.Normalize([Text.NormalizationForm]::FormD)
    $builder = New-Object Text.StringBuilder
    foreach ($char in $normalized.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($char) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$builder.Append($char)
        }
    }

    return $builder.ToString().ToLowerInvariant()
}

function ConvertTo-CleanUrl {
    param([string]$Url)

    $clean = ConvertFrom-HtmlAttribute $Url
    $clean = $clean -replace "\?.*$", ""
    return $clean
}

function ConvertTo-DateTimeOffsetOrNull {
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    $parsed = [DateTimeOffset]::MinValue
    $styles = [Globalization.DateTimeStyles]::AllowWhiteSpaces
    if ([DateTimeOffset]::TryParse($Value, [Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$parsed)) {
        return $parsed
    }

    return $null
}

function Test-IsRecent {
    param([AllowNull()][DateTimeOffset]$PublishedAt)

    if ($null -eq $PublishedAt) {
        return $false
    }

    return $PublishedAt -ge $Cutoff
}

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

function Get-JobMatch {
    param(
        [string]$Title,
        [string]$Text
    )

    $titleText = ConvertTo-MatchText $Title
    $fullText = ConvertTo-MatchText ("{0} {1}" -f $Title, $Text)
    $coreTitlePattern = "\bweb\s*analyst\b|\bdigital\s*analyst\b|analyste\s+(digital|web)|digital\s+analytics?\s+consultant|analytics?\s+consultant|web\s+analytics?\s+consultant|web\s+analytics|digital\s+analytics|tracking|webtracking|taggage|tagging|\bdata\s*analyst\b|analyste\s+de\s+donnees|performance\s+digital|performance\s+digitale|digital\s+performance|\bcro\b|conversion\s+rate|conversion\s+optimization|optimisation\s+conversion"
    $hasCoreTitleSignal = $titleText -match $coreTitlePattern
    $state = @{
        Score = 0
        Keywords = @{}
    }
    $isGoToMarketContext = $fullText -match "go\s*[- ]?\s*to\s*[- ]?\s*market|\banalytics\s+engineer\s+gtm\b|\bgtm\s+(strategy|strategies|motion|motions|operations|ops|sales|revenue|revops|demand|pipeline|finance|engineer|generation|outbound|inbound)\b|\b(growth|demand|pipeline|revops|revenue|sales|finance|outbound|inbound)\s+gtm\b|\b(head|director|manager|lead|chief)\s+of\s+gtm\b|\b(sdr|bdr)\b.*\bgtm\b"
    $webAnalyticsToolPattern = "google\s+tag\s+manager|google\s+analytics|\bga4\b|piano\s+analytics|contentsquare|content\s+square|matomo|adobe\s+analytics|omniture|data\s*layer|datalayer|tagging\s+plan|tracking\s+plan|plan\s+de\s+(taggage|marquage)|consent\s+mode|cookie\s+consent"
    $hasWebAnalyticsToolSignal = ($fullText -match $webAnalyticsToolPattern) -or (-not $isGoToMarketContext -and $fullText -match "\bgtm\b")
    $hasDigitalAnalyticsContext = $hasWebAnalyticsToolSignal -or ($titleText -match "\bweb\s*analyst\b|\bdigital\s*analyst\b|analyste\s+(digital|web)|web\s+analytics|digital\s+analytics|tracking|webtracking|taggage|tagging|analytics?\s+consultant|performance\s+digital|performance\s+digitale|digital\s+performance|\bcro\b|conversion")
    $isMarketingOnlyContext = ($fullText -match "\bseo\b|\bsea\b|paid\s+social|paid\s+search|paid\s+media|performance\s+marketing|growth\s+marketing|acquisition\s+marketing|digital\s+marketing|social\s+media|content\s+marketing|campaign\s+manager|media\s+buyer") -and -not $hasDigitalAnalyticsContext
    $isDataWarehouseContext = $fullText -match "\bdbt\b|snowflake|airflow|\betl\b|\belt\b|data\s+warehouse|datawarehouse|data\s+platform|databricks|pyspark|spark|data\s+pipeline|\bpython\b"

    Add-MatchSignal $state $titleText "\bweb\s*analyst\b" "title:web analyst" 60
    Add-MatchSignal $state $titleText "\bdigital\s*analyst\b|analyste\s+digital|analyste\s+web" "title:digital/web analyst" 55
    Add-MatchSignal $state $titleText "digital\s+analytics?\s+consultant|analytics?\s+consultant|web\s+analytics?\s+consultant" "title:analytics consultant" 55
    Add-MatchSignal $state $titleText "tracking|webtracking|taggage|tagging" "title:tracking/tagging" 50
    Add-MatchSignal $state $titleText "web\s+analytics|digital\s+analytics" "title:web/digital analytics" 45
    Add-MatchSignal $state $titleText "\bdata\s*analyst\b|analyste\s+de\s+donnees" "title:data analyst" 25
    Add-MatchSignal $state $titleText "performance\s+digital|performance\s+digitale|digital\s+performance" "title:digital performance" 30
    Add-MatchSignal $state $titleText "\bcro\b|conversion\s+rate|conversion\s+optimization|optimisation\s+conversion" "title:CRO" 25

    Add-MatchSignal $state $fullText "google\s+tag\s+manager" "Google Tag Manager" 35
    if (-not $isGoToMarketContext -and $fullText -match "\bgtm\b") {
        $state.Score += 35
        $state.Keywords["GTM"] = $true
    }
    Add-MatchSignal $state $fullText "google\s+analytics|\bga4\b" "Google Analytics/GA4" 35
    Add-MatchSignal $state $fullText "piano\s+analytics" "Piano Analytics" 35
    Add-MatchSignal $state $fullText "contentsquare|content\s+square" "ContentSquare" 35
    Add-MatchSignal $state $fullText "data\s*layer|datalayer" "dataLayer" 25
    Add-MatchSignal $state $fullText "plan\s+de\s+(taggage|marquage)|tagging\s+plan|tracking\s+plan" "tagging plan" 30
    Add-MatchSignal $state $fullText "tracking|taggage|tagging|tag\s+management" "tracking/tagging" 20
    Add-MatchSignal $state $fullText "consent\s+mode|cmp|cookie\s+consent" "consent tracking" 20
    Add-MatchSignal $state $fullText "a/b\s*test|ab\s*test|experimentation" "A/B testing" 15
    Add-MatchSignal $state $fullText "matomo|adobe\s+analytics|omniture" "other analytics tools" 20
    Add-MatchSignal $state $fullText "looker\s+studio|data\s+studio|dashboard|reporting|tableau\s+de\s+bord" "reporting/dashboard" 10
    Add-MatchSignal $state $fullText "\bkpi\b|conversion|funnel|parcours\s+utilisateur|user\s+journey" "analytics responsibilities" 10

    if ($isMarketingOnlyContext) {
        $state.Score -= 25
        $state.Keywords["possible SEO/SEA/marketing-only role"] = $true
    }
    elseif ($fullText -match "\bseo\b|\bsea\b|paid\s+social|performance\s+marketing|growth\s+marketing|digital\s+marketing") {
        $state.Score -= 8
        $state.Keywords["possible marketing role"] = $true
    }
    if ($titleText -match "time\s+tracking|absence|payroll" -and $titleText -notmatch "web|analytics|tagging|taggage|webtracking|\bcro\b") {
        $state.Score -= 60
        $state.Keywords["possible time-tracking/payroll role"] = $true
    }
    if ($titleText -match "\bseo\s+specialist\b" -and $titleText -notmatch "analytics|web\s+analytics|tracking|tagging|taggage") {
        $state.Score -= 60
        $state.Keywords["possible SEO-only role"] = $true
    }
    if ($isGoToMarketContext) {
        $state.Score -= 35
        $state.Keywords["possible go-to-market role"] = $true
    }
    if ($fullText -match "business\s+analyst|risk|risque|finance|banking|bancaire") {
        $state.Score -= 10
        $state.Keywords["possible broad analyst role"] = $true
    }
    if ((($titleText -match "\bdata\s*analyst\b|analyste\s+de\s+donnees|analytics?\s+engineer|data\s+engineer") -or $isDataWarehouseContext) -and -not $hasDigitalAnalyticsContext) {
        $state.Score -= 25
        $state.Keywords["possible data analyst/engineering role"] = $true
    }
    elseif ($isDataWarehouseContext) {
        $state.Score -= 12
        $state.Keywords["possible warehouse/python role"] = $true
    }
    if ($fullText -match "software\s+engineer|data\s+engineer|sales\s+engineer|backend|frontend|devops") {
        $state.Score -= 15
        $state.Keywords["possible engineering role"] = $true
    }

    if ($state.Score -lt 0) {
        $state.Score = 0
    }
    if (-not $hasCoreTitleSignal -and $state.Score -gt 49) {
        $state.Score = 49
        $state.Keywords["no core title signal"] = $true
    }

    $level = "Review"
    if ($state.Score -ge 80) {
        $level = "High"
    }
    elseif ($state.Score -ge 50) {
        $level = "Medium"
    }

    [PSCustomObject]@{
        IsMatch = $state.Score -ge $MinimumMatchScore
        Score = [int]$state.Score
        Level = $level
        Keywords = (($state.Keywords.Keys | Sort-Object) -join "; ")
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

    if ($Text -match "(?i)(\bCDI\b|contrat\s+(a|à)\s+dur[eé]e\s+ind[eé]termin[eé]e)") {
        return "CDI"
    }
    if ($Text -match "(?i)(alternance|alternant|apprentissage|apprenticeship)") {
        return "Apprenticeship"
    }
    if ($Text -match "(?i)(stage|stagiaire|internship|intern\b)") {
        return "Internship"
    }
    if ($Text -match "(?i)(\bCDD\b|contrat\s+(a|à)\s+dur[eé]e\s+d[eé]termin[eé]e|fixed[-\s]*term|temporary|temporaire)") {
        return "CDD"
    }
    if ($Text -match "(?i)(freelance|contractor|ind[eé]pendant)") {
        return "Freelance"
    }
    if ($Text -match "(?i)(Employment\s+type\s+Full-time|Type\s+d[’']emploi\s+Temps\s+plein|\bFull-time\b|\bTemps\s+plein\b)") {
        return "Full-time"
    }
    if ($Text -match "(?i)(permanent\s+(contract|position|role)|\bpermanent\b)") {
        return "Permanent"
    }

    return ""
}

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

function ConvertTo-ExcelHyperlinkFormula {
    param(
        [AllowNull()][string]$Url,
        [string]$Label = "Open"
    )

    if ([string]::IsNullOrWhiteSpace($Url)) {
        return ""
    }

    $cleanUrl = $Url.Trim()
    if ($cleanUrl -notmatch "^https?://") {
        return $cleanUrl
    }

    $escapedUrl = $cleanUrl.Replace('"', '""')
    $escapedLabel = $Label.Replace('"', '""')
    return '=HYPERLINK("{0}","{1}")' -f $escapedUrl, $escapedLabel
}

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

function Get-DedupeCompanyKey {
    param([AllowNull()][string]$CompanyName)

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
    $strongTokens = @($tokens | Where-Object { $_.Length -gt 1 -and $noise -notcontains $_ })
    if ($strongTokens.Count -eq 0) {
        $strongTokens = $tokens
    }

    $weakCompanyNames = @("confidential", "jobgether", "licorne", "recrutement", "talent")
    if ($weakCompanyNames -contains $strongTokens[0] -and $strongTokens.Count -gt 1) {
        return ($strongTokens | Select-Object -First 2) -join " "
    }

    if ($strongTokens.Count -eq 1) {
        return $strongTokens[0]
    }

    return ($strongTokens | Select-Object -First 2) -join " "
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
    $tokens = @($text -split "\s+" | Where-Object { $_.Length -gt 1 -and $noise -notcontains $_ })
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

    if ($TitleKey -match "\b(web analyst|digital analyst|tracking|tagging|analytics|web analytics|digital analytics|gtm|cro|data analyst)\b") {
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

function Get-RowValue {
    param(
        [AllowNull()]$Row,
        [string]$Name
    )

    if ($null -eq $Row) {
        return ""
    }

    if (@($Row.PSObject.Properties.Name) -contains $Name) {
        $value = $Row.PSObject.Properties[$Name].Value
        if ($null -ne $value) {
            return [string]$value
        }
    }

    return ""
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

function Test-IsExcludedContractType {
    param([AllowNull()][string]$ContractType)

    $contractText = ConvertTo-MatchText $ContractType
    if ([string]::IsNullOrWhiteSpace($contractText)) {
        return $false
    }

    return $contractText -match "\bcdd\b|apprenticeship|apprentissage|alternance|internship|\bstage\b|stagiaire|temporary|fixed\s+term"
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
        [DateTimeOffset]$PublishedAt
    )

    if ([string]::IsNullOrWhiteSpace($Title) -or [string]::IsNullOrWhiteSpace($Url)) {
        return $null
    }

    $key = "{0}|{1}" -f $Platform, $Url.ToLowerInvariant()
    if ($SeenResultKeys.ContainsKey($key)) {
        return $null
    }

    $identityKey = Get-JobIdentityKeyFromValues -Title $Title -CompanyName $CompanyName -JobLocation $JobLocation -Url $Url
    $jobId = Get-StableJobId $identityKey

    $SeenResultKeys[$key] = $true
    [PSCustomObject]@{
        job_id         = $jobId
        job_title      = $Title.Trim()
        company_name   = $CompanyName.Trim()
        location       = $JobLocation.Trim()
        contract_type  = $ContractType.Trim()
        match_score    = $MatchScore
        match_level    = $MatchLevel.Trim()
        matched_keywords = $MatchedKeywords.Trim()
        feedback_adjustment = ""
        job_url        = ConvertTo-ExcelHyperlinkFormula -Url $Url -Label "Open"
        job_url_raw    = $Url.Trim()
        platform       = $Platform
        source_count   = "1"
        alternate_urls = ""
        published_date = $PublishedAt.ToString("yyyy-MM-dd")
    }
}

$MasterColumns = Get-JobTrackerMasterColumns
$ColumnLabels = Get-JobTrackerColumnLabels

function New-OrderedJobRecord {
    param([hashtable]$Values)

    $ordered = [ordered]@{}
    foreach ($column in $MasterColumns) {
        if ($Values.ContainsKey($column) -and $null -ne $Values[$column]) {
            $ordered[$column] = [string]$Values[$column]
        }
        else {
            $ordered[$column] = ""
        }
    }

    return [PSCustomObject]$ordered
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

function Import-TrackerRowsFromXlsx {
    param([string]$Path)

    $excel = $null
    $workbook = $null
    $sheet = $null
    $usedRange = $null

    try {
        $fullPath = (Resolve-Path $Path).Path
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($fullPath, $null, $true)

        try {
            $sheet = $workbook.Worksheets.Item("Jobs")
        }
        catch {
            $sheet = $workbook.Worksheets.Item(1)
        }

        $usedRange = $sheet.UsedRange
        $rowCount = [int]$usedRange.Rows.Count
        $columnCount = [int]$usedRange.Columns.Count
        if ($rowCount -lt 2 -or $columnCount -lt 1) {
            return @()
        }

        $headers = New-Object System.Collections.Generic.List[string]
        for ($column = 1; $column -le $columnCount; $column++) {
            $header = [string]$sheet.Cells.Item(1, $column).Text
            if ([string]::IsNullOrWhiteSpace($header)) {
                $header = "Column$column"
            }
            $headers.Add((ConvertTo-CanonicalColumnName $header.Trim())) | Out-Null
        }

        $rows = New-Object System.Collections.Generic.List[object]
        for ($row = 2; $row -le $rowCount; $row++) {
            $values = @{}
            $hasValue = $false
            for ($column = 1; $column -le $columnCount; $column++) {
                $name = $headers[$column - 1]
                $value = [string]$sheet.Cells.Item($row, $column).Text
                if (-not [string]::IsNullOrWhiteSpace($value)) {
                    $hasValue = $true
                }
                $values[$name] = $value
            }

            if ($hasValue) {
                $rows.Add((New-OrderedJobRecord $values)) | Out-Null
            }
        }

        return @($rows.ToArray())
    }
    finally {
        if ($null -ne $workbook) {
            $workbook.Close($false) | Out-Null
        }
        if ($null -ne $excel) {
            $excel.Quit() | Out-Null
        }

        Release-ComObject $usedRange
        Release-ComObject $sheet
        Release-ComObject $workbook
        Release-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Get-SummaryValue {
    param(
        [AllowNull()]$Summary,
        [string]$Name
    )

    if ($null -ne $Summary -and $Summary.ContainsKey($Name)) {
        return [string]$Summary[$Name]
    }

    return ""
}

function Export-TrackerWorkbook {
    param(
        [object[]]$Rows,
        [string]$Path,
        [AllowNull()]$Summary = $null
    )

    $directory = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path $directory)) {
        New-Item -ItemType Directory -Force -Path $directory | Out-Null
    }

    $fullPath = [IO.Path]::GetFullPath($Path)
    $excel = $null
    $workbook = $null
    $jobsSheet = $null
    $summarySheet = $null
    $tableRange = $null
    $table = $null
    $dataRange = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Add()

        while ([int]$workbook.Worksheets.Count -gt 1) {
            $workbook.Worksheets.Item([int]$workbook.Worksheets.Count).Delete()
        }

        $jobsSheet = $workbook.Worksheets.Item(1)
        $jobsSheet.Name = "Jobs"
        $summarySheet = $workbook.Worksheets.Add([System.Type]::Missing, $jobsSheet)
        $summarySheet.Name = "Summary"

        $columnIndex = @{}
        for ($index = 0; $index -lt $MasterColumns.Count; $index++) {
            $columnNumber = $index + 1
            $columnName = $MasterColumns[$index]
            $columnIndex[$columnName] = $columnNumber
            $jobsSheet.Cells.Item(1, $columnNumber).Value2 = Get-ColumnLabel $columnName
        }

        $rowCount = @($Rows).Count
        for ($rowIndex = 0; $rowIndex -lt $rowCount; $rowIndex++) {
            $row = $Rows[$rowIndex]
            $excelRow = $rowIndex + 2
            foreach ($columnName in $MasterColumns) {
                $excelColumn = [int]$columnIndex[$columnName]
                $cell = $jobsSheet.Cells.Item($excelRow, $excelColumn)
                $value = Get-RowValue -Row $row -Name $columnName

                if ($columnName -eq "job_url") {
                    $url = Get-RowValue -Row $row -Name "job_url_raw"
                    if ([string]::IsNullOrWhiteSpace($url) -and $value -match "^https?://") {
                        $url = $value
                    }

                    if ($url -match "^https?://") {
                        $escapedUrl = $url.Replace('"', '""')
                        $cell.Formula = '=HYPERLINK("{0}","Open")' -f $escapedUrl
                    }
                    else {
                        $cell.Value2 = $value
                    }
                }
                elseif ($columnName -in @("match_score", "days_since_published", "days_since_first_seen", "days_since_last_seen", "feedback_adjustment")) {
                    $number = 0
                    if ([int]::TryParse($value, [ref]$number)) {
                        $cell.Value2 = [string]$number
                    }
                    else {
                        $cell.Value2 = $value
                    }
                }
                else {
                    $cell.Value2 = $value
                }
            }
        }

        $lastDataRow = [Math]::Max(2, $rowCount + 1)
        $lastColumn = $MasterColumns.Count
        $tableRange = $jobsSheet.Range($jobsSheet.Cells.Item(1, 1), $jobsSheet.Cells.Item($lastDataRow, $lastColumn))
        try {
            $table = $jobsSheet.ListObjects.Add(1, $tableRange, $null, 1)
            $table.Name = "JobsTracker"
            $table.TableStyle = "TableStyleLight9"
            $table.ShowTableStyleRowStripes = $false
            $table.ShowTableStyleColumnStripes = $false
        }
        catch {
            $jobsSheet.Range($jobsSheet.Cells.Item(1, 1), $jobsSheet.Cells.Item(1, $lastColumn)).AutoFilter() | Out-Null
        }

        $headerColor = Get-ExcelColor 38 50 56
        $darkTextColor = Get-ExcelColor 40 47 52
        $mutedTextColor = Get-ExcelColor 100 116 139
        $jobsSheet.Rows.Item(1).Font.Bold = $true
        $jobsSheet.Rows.Item(1).Font.Color = Get-ExcelColor 255 255 255
        $jobsSheet.Rows.Item(1).Interior.Color = $headerColor
        $jobsSheet.Rows.Item(1).VerticalAlignment = -4108
        $jobsSheet.Rows.Item(1).RowHeight = 24
        $jobsSheet.Cells.Font.Name = "Segoe UI"
        $jobsSheet.Cells.Font.Size = 10
        $jobsSheet.Cells.Font.Color = $darkTextColor
        $jobsSheet.Rows.Item(1).Font.Color = Get-ExcelColor 255 255 255

        foreach ($columnName in @("duplicate_reason", "job_title", "matched_keywords", "job_url_raw", "notes")) {
            if ($columnIndex.ContainsKey($columnName)) {
                $jobsSheet.Columns.Item([int]$columnIndex[$columnName]).WrapText = $true
            }
        }
        $jobsSheet.Columns.AutoFit() | Out-Null
        $columnSizing = Get-JobTrackerColumnSizing
        foreach ($entry in $columnSizing.GetEnumerator()) {
            if ($columnIndex.ContainsKey($entry.Key)) {
                $column = $jobsSheet.Columns.Item([int]$columnIndex[$entry.Key])
                $minWidth = [double]$entry.Value.Min
                $maxWidth = [double]$entry.Value.Max
                if ([double]$column.ColumnWidth -lt $minWidth) {
                    $column.ColumnWidth = $minWidth
                }
                elseif ([double]$column.ColumnWidth -gt $maxWidth) {
                    $column.ColumnWidth = $maxWidth
                }
            }
        }
        Set-JobTrackerColumnVisibility -Sheet $jobsSheet -ColumnIndex $columnIndex
        foreach ($columnName in @("review_priority", "status", "contract_type", "platform", "source_count", "published_date", "days_since_published", "job_url", "applied_date", "match_level", "match_score", "seen_in_current_crawl", "first_seen_date", "last_seen_date", "is_new")) {
            if ($columnIndex.ContainsKey($columnName)) {
                $jobsSheet.Columns.Item([int]$columnIndex[$columnName]).HorizontalAlignment = -4108
            }
        }
        foreach ($columnName in @("published_date", "applied_date", "first_seen_date", "last_seen_date")) {
            if ($columnIndex.ContainsKey($columnName)) {
                $jobsSheet.Columns.Item([int]$columnIndex[$columnName]).NumberFormat = "yyyy-mm-dd"
            }
        }
        foreach ($columnName in @("source_count", "days_since_published", "match_score", "days_since_first_seen", "days_since_last_seen", "feedback_adjustment")) {
            if ($columnIndex.ContainsKey($columnName)) {
                $jobsSheet.Columns.Item([int]$columnIndex[$columnName]).NumberFormat = "0"
            }
        }

        Set-JobTrackerDataValidation -Workbook $workbook -Excel $excel -Sheet $jobsSheet -ColumnIndex $columnIndex -LastDataRow $lastDataRow

        if ($rowCount -gt 0) {
            $dataRange = $jobsSheet.Range($jobsSheet.Cells.Item(2, 1), $jobsSheet.Cells.Item($lastDataRow, $lastColumn))
            $dataRange.VerticalAlignment = -4160
            $dataRange.RowHeight = 30
            $dataRange.Interior.Color = Get-ExcelColor 255 255 255
            Set-StatusRowConditionalFormatting -Range $dataRange -ColumnIndex $columnIndex
            Set-IgnoredNotesReminderFormatting -Sheet $jobsSheet -ColumnIndex $columnIndex -LastDataRow $lastDataRow
        }

        $colors = Get-JobTrackerWorkbookColors -DarkTextColor $darkTextColor
        for ($excelRow = 2; $excelRow -le ($rowCount + 1); $excelRow++) {
            $status = (Get-RowValue -Row $Rows[$excelRow - 2] -Name "status").ToLowerInvariant()
            $matchLevel = Get-RowValue -Row $Rows[$excelRow - 2] -Name "match_level"
            $priority = Get-RowValue -Row $Rows[$excelRow - 2] -Name "review_priority"
            $seen = Get-RowValue -Row $Rows[$excelRow - 2] -Name "seen_in_current_crawl"

            if ($columnIndex.ContainsKey("status")) {
                $statusCell = $jobsSheet.Cells.Item($excelRow, [int]$columnIndex["status"])
                $statusCell.Font.Bold = $false
                switch ($status) {
                    "interesting" { $statusCell.Font.Color = $colors.AmberText; $statusCell.Font.Bold = $true }
                    "applied" { $statusCell.Font.Color = $colors.GreenText; $statusCell.Font.Bold = $true }
                    "interview" { $statusCell.Font.Color = $colors.BlueText; $statusCell.Font.Bold = $true }
                    "offer" { $statusCell.Font.Color = $colors.GreenText; $statusCell.Font.Bold = $true }
                    "ignored" { $statusCell.Font.Color = $colors.GrayText }
                    "rejected" { $statusCell.Font.Color = $colors.RedText }
                    "withdrawn" { $statusCell.Font.Color = $colors.GrayText }
                    default { $statusCell.Font.Color = $colors.DarkText }
                }
            }

            if ($columnIndex.ContainsKey("match_level")) {
                $matchCell = $jobsSheet.Cells.Item($excelRow, [int]$columnIndex["match_level"])
                $matchCell.Font.Bold = $true
                switch ($matchLevel) {
                    "High" { $matchCell.Font.Color = $colors.GreenText }
                    "Medium" { $matchCell.Font.Color = $colors.AmberText }
                    "Review" { $matchCell.Font.Color = $colors.GrayText; $matchCell.Font.Bold = $false }
                    default { $matchCell.Font.Color = $colors.DarkText }
                }
            }

            if ($columnIndex.ContainsKey("review_priority")) {
                $priorityCell = $jobsSheet.Cells.Item($excelRow, [int]$columnIndex["review_priority"])
                $priorityCell.Font.Bold = $true
                switch ($priority) {
                    "Application" { $priorityCell.Font.Color = $colors.GreenText }
                    "New High" { $priorityCell.Font.Color = $colors.AmberText }
                    "Ignored" { $priorityCell.Font.Color = $colors.GrayText; $priorityCell.Font.Bold = $false }
                    default { $priorityCell.Font.Color = $colors.DarkText; $priorityCell.Font.Bold = $false }
                }
            }

            if ($seen -eq "no" -and $columnIndex.ContainsKey("seen_in_current_crawl")) {
                $jobsSheet.Cells.Item($excelRow, [int]$columnIndex["seen_in_current_crawl"]).Font.Color = $mutedTextColor
            }
        }

        $jobsSheet.Activate() | Out-Null
        $excel.ActiveWindow.SplitRow = 1
        $excel.ActiveWindow.SplitColumn = 2
        $excel.ActiveWindow.FreezePanes = $true
        $excel.ActiveWindow.DisplayGridlines = $false

        $summarySheet.Cells.Item(1, 1).Value2 = "Analytics Job Tracker"
        $summarySheet.Cells.Item(1, 1).Font.Bold = $true
        $summarySheet.Cells.Item(1, 1).Font.Size = 16
        $summarySheet.Cells.Font.Name = "Segoe UI"
        $summarySheet.Cells.Font.Size = 10
        $summarySheet.Cells.Font.Color = Get-ExcelColor 40 47 52
        $summarySheet.Cells.Item(1, 1).Font.Bold = $true
        $summarySheet.Cells.Item(1, 1).Font.Size = 16
        $currentVisibleCount = @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "seen_in_current_crawl") -eq "yes" }).Count
        $newVisibleCount = @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "is_new") -eq "yes" }).Count
        $applicationVisibleCount = @($Rows | Where-Object { Test-IsAppliedStatus (Get-RowValue -Row $_ -Name "status") }).Count
        $highVisibleCount = @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "match_level") -eq "High" }).Count
        $mediumVisibleCount = @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "match_level") -eq "Medium" }).Count
        $reviewVisibleCount = @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "match_level") -eq "Review" }).Count
        $summaryPairs = @(
            @("Generated", $RunStamp),
            @("Retention rule", "Keep non-application jobs only when Published is on or after $CutoffDate."),
            @("Rows in workbook", [string](@($Rows).Count)),
            @("Seen in this crawl", [string]$currentVisibleCount),
            @("New this run", [string]$newVisibleCount),
            @("Application rows kept", [string]$applicationVisibleCount),
            @("Match levels", ("High {0} | Medium {1} | Review {2}" -f $highVisibleCount, $mediumVisibleCount, $reviewVisibleCount)),
            @("Total matched before contract filter", (Get-SummaryValue -Summary $Summary -Name "TotalMatched")),
            @("Excluded CDD/apprenticeship/internship", (Get-SummaryValue -Summary $Summary -Name "ExcludedContractCount")),
            @("Duplicates merged this run", (Get-SummaryValue -Summary $Summary -Name "DuplicateCount")),
            @("Rows removed by retention", (Get-SummaryValue -Summary $Summary -Name "RemovedCount")),
            @("Backup", (Get-SummaryValue -Summary $Summary -Name "BackupPath")),
            @("Tracker", $fullPath),
            @("Manual fields", "Status, Applied date, Apply notes with ignore_reason templates"),
            @("Reminder", "Close this workbook before launching the crawler.")
        )
        $summaryRow = 3
        foreach ($pair in $summaryPairs) {
            $summarySheet.Cells.Item($summaryRow, 1).Value2 = [string]$pair[0]
            $summarySheet.Cells.Item($summaryRow, 2).Value2 = [string]$pair[1]
            $summaryRow++
        }
        $summarySheet.Rows.Item(1).Font.Color = Get-ExcelColor 38 50 56
        $summarySheet.Columns.Item(1).ColumnWidth = 36
        $summarySheet.Columns.Item(2).ColumnWidth = 90
        $summarySheet.Columns.Item(2).WrapText = $true
        $summarySheet.Range("A3:A$($summaryRow - 1)").Font.Bold = $true
        $summarySheet.Range("A3:B$($summaryRow - 1)").Borders.LineStyle = 1
        $summarySheet.Range("A3:B$($summaryRow - 1)").Borders.Color = Get-ExcelColor 226 232 240

        $workbook.Worksheets.Item("Jobs").Activate() | Out-Null
        $workbook.SaveAs($fullPath, 51)
    }
    finally {
        if ($null -ne $workbook) {
            $workbook.Close($false) | Out-Null
        }
        if ($null -ne $excel) {
            $excel.Quit() | Out-Null
        }

        Release-ComObject $table
        Release-ComObject $tableRange
        Release-ComObject $dataRange
        Release-ComObject $summarySheet
        Release-ComObject $jobsSheet
        Release-ComObject $workbook
        Release-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function ConvertTo-DateOrNull {
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    $parsed = [DateTime]::MinValue
    if ([DateTime]::TryParse($Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AllowWhiteSpaces, [ref]$parsed)) {
        return $parsed.Date
    }

    return $null
}

function Get-DaysSince {
    param([AllowNull()][string]$DateText)

    $date = ConvertTo-DateOrNull $DateText
    if ($null -eq $date) {
        return ""
    }

    return [string]([Math]::Max(0, ([DateTime]::Today - $date).Days))
}

function Test-IsKeepForeverStatus {
    param([AllowNull()][string]$Status)

    return Test-IsAppliedStatus $Status
}

function Test-IsRecentTrackerRow {
    param([AllowNull()]$Row)

    $publishedDate = ConvertTo-DateOrNull (Get-RowValue -Row $Row -Name "published_date")
    return ($null -ne $publishedDate -and $publishedDate -ge $Cutoff.Date)
}

function Get-IntegerRowValue {
    param(
        [AllowNull()]$Row,
        [string]$Name
    )

    $value = 0
    if ([int]::TryParse((Get-RowValue -Row $Row -Name $Name), [ref]$value)) {
        return $value
    }

    return 0
}

function Get-SourcePreference {
    param([AllowNull()]$Row)

    $platform = ConvertTo-MatchText (Get-RowValue -Row $Row -Name "platform")
    if ($platform -match "welcome|jungle|wttj") {
        return 30
    }
    if ($platform -match "linkedin") {
        return 20
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
    foreach ($row in $rowList) {
        $score = Get-IntegerRowValue -Row $row -Name "match_score"
        if ($score -gt $maxScore) {
            $maxScore = $score
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
    $values["source_count"] = [string]([Math]::Max(1, $urls.Count))
    $values["platform"] = Join-UniqueTextValues -Values @($rowList | ForEach-Object { Get-RowValue -Row $_ -Name "platform" })
    $values["matched_keywords"] = Join-UniqueTextValues -Values @($rowList | ForEach-Object { Get-RowValue -Row $_ -Name "matched_keywords" }) -SplitPattern "\s*;\s*|\s*,\s*"
    $values["duplicate_reason"] = Join-CleanTextParts $reasonParts
    $values["match_score"] = [string]$maxScore
    $values["match_level"] = Get-MatchLevelFromScore $maxScore
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

function Backup-TrackerFile {
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        return ""
    }

    $backupDirectory = Join-Path (Split-Path -Parent $Path) "backups"
    if (-not (Test-Path $backupDirectory)) {
        New-Item -ItemType Directory -Force -Path $backupDirectory | Out-Null
    }

    $baseName = [IO.Path]::GetFileNameWithoutExtension($Path)
    $extension = [IO.Path]::GetExtension($Path)
    $backupPath = Join-Path $backupDirectory ("{0}_{1}{2}" -f $baseName, $RunStamp, $extension)
    Copy-Item -LiteralPath $Path -Destination $backupPath -Force

    if ($MaxBackups -gt 0) {
        $oldBackups = @(Get-ChildItem -LiteralPath $backupDirectory -File -Filter ("{0}_*{1}" -f $baseName, $extension) | Sort-Object LastWriteTime -Descending)
        if ($oldBackups.Count -gt $MaxBackups) {
            $oldBackups | Select-Object -Skip $MaxBackups | ForEach-Object {
                Remove-Item -LiteralPath $_.FullName -Force
            }
        }
    }

    return $backupPath
}

function Import-TrackerRows {
    param([string]$Path)

    if (Test-Path $Path) {
        if ([IO.Path]::GetExtension($Path).ToLowerInvariant() -ne ".xlsx") {
            throw "Unsupported tracker file type. This project uses only output\jobs_tracker.xlsx."
        }

        return @(Import-TrackerRowsFromXlsx -Path $Path)
    }

    return @()
}

function Get-ReviewPriority {
    param(
        [AllowNull()][string]$Status,
        [AllowNull()][string]$MatchLevel,
        [AllowNull()][string]$IsNew
    )

    if (Test-IsKeepForeverStatus $Status) {
        return "Application"
    }
    if ((ConvertTo-MatchText $Status) -eq "ignored") {
        return "Ignored"
    }
    if ($IsNew -eq "yes" -and $MatchLevel -eq "High") {
        return "New High"
    }
    if ($IsNew -eq "yes") {
        return "New"
    }
    if ($MatchLevel -eq "High") {
        return "High"
    }
    return $MatchLevel
}

function ConvertTo-TrackerRecord {
    param(
        $CurrentRow,
        [AllowNull()]$ExistingRow = $null,
        [bool]$SeenInCurrentCrawl = $true,
        [AllowNull()][string]$DuplicateReason = ""
    )

    $existingJobId = Get-RowValue -Row $ExistingRow -Name "job_id"
    $currentJobId = Get-RowValue -Row $CurrentRow -Name "job_id"
    $jobId = Get-PreferredValue -Primary $existingJobId -Fallback (Get-PreferredValue -Primary $currentJobId -Fallback (Get-StableJobId (Get-JobIdentityKeyFromRow $CurrentRow)))

    $status = Get-RowValue -Row $ExistingRow -Name "status"
    if ([string]::IsNullOrWhiteSpace($status)) {
        $status = "new"
    }

    $firstSeen = Get-RowValue -Row $ExistingRow -Name "first_seen_date"
    $isNew = "no"
    $seenBefore = "yes"
    if ([string]::IsNullOrWhiteSpace($firstSeen)) {
        $firstSeen = $RunDate
        $isNew = "yes"
        $seenBefore = "no"
    }

    $lastSeen = Get-RowValue -Row $ExistingRow -Name "last_seen_date"
    if ($SeenInCurrentCrawl) {
        $lastSeen = $RunDate
    }
    elseif ([string]::IsNullOrWhiteSpace($lastSeen)) {
        $lastSeen = $firstSeen
    }

    $matchLevel = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "match_level") -Fallback (Get-RowValue -Row $ExistingRow -Name "match_level")
    $duplicateValue = Get-PreferredValue -Primary $DuplicateReason -Fallback (Get-RowValue -Row $ExistingRow -Name "duplicate_reason")
    $publishedDate = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "published_date") -Fallback (Get-RowValue -Row $ExistingRow -Name "published_date")
    $primaryUrl = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "job_url_raw") -Fallback (Get-RowValue -Row $ExistingRow -Name "job_url_raw")
    $allUrls = @(Get-RowUrlValues @($CurrentRow, $ExistingRow))
    $alternateUrls = @($allUrls | Where-Object { $_ -ne $primaryUrl })

    return New-OrderedJobRecord @{
        job_id                = $jobId
        status                = $status
        applied_date          = Get-RowValue -Row $ExistingRow -Name "applied_date"
        first_seen_date       = $firstSeen
        last_seen_date        = $lastSeen
        is_new                = $isNew
        seen_before           = $seenBefore
        seen_in_current_crawl = $(if ($SeenInCurrentCrawl) { "yes" } else { "no" })
        days_since_first_seen = Get-DaysSince $firstSeen
        days_since_last_seen  = Get-DaysSince $lastSeen
        duplicate_reason      = $duplicateValue
        feedback_adjustment   = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "feedback_adjustment") -Fallback (Get-RowValue -Row $ExistingRow -Name "feedback_adjustment")
        review_priority       = Get-ReviewPriority -Status $status -MatchLevel $matchLevel -IsNew $isNew
        job_title             = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "job_title") -Fallback (Get-RowValue -Row $ExistingRow -Name "job_title")
        company_name          = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "company_name") -Fallback (Get-RowValue -Row $ExistingRow -Name "company_name")
        location              = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "location") -Fallback (Get-RowValue -Row $ExistingRow -Name "location")
        contract_type         = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "contract_type") -Fallback (Get-RowValue -Row $ExistingRow -Name "contract_type")
        match_score           = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "match_score") -Fallback (Get-RowValue -Row $ExistingRow -Name "match_score")
        match_level           = $matchLevel
        matched_keywords      = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "matched_keywords") -Fallback (Get-RowValue -Row $ExistingRow -Name "matched_keywords")
        job_url               = ConvertTo-ExcelHyperlinkFormula -Url $primaryUrl -Label "Open"
        job_url_raw           = $primaryUrl
        alternate_urls        = ($alternateUrls -join "; ")
        platform              = Join-UniqueTextValues -Values @((Get-RowValue -Row $CurrentRow -Name "platform"), (Get-RowValue -Row $ExistingRow -Name "platform"))
        source_count          = [string]([Math]::Max(1, $allUrls.Count))
        published_date        = $publishedDate
        days_since_published  = Get-DaysSince $publishedDate
        notes                 = Get-RowValue -Row $ExistingRow -Name "notes"
    }
}

function ConvertTo-TrackerRecordFromExisting {
    param($ExistingRow)

    return ConvertTo-TrackerRecord -CurrentRow $ExistingRow -ExistingRow $ExistingRow -SeenInCurrentCrawl:$false
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

    return $Text -match "web\s+analytics|digital\s+analytics|web\s*analyst|digital\s*analyst|tracking|tagging|taggage|webtracking|google\s+tag\s+manager|\bgtm\b|google\s+analytics|\bga4\b|piano|contentsquare|content\s+square|data\s*layer|datalayer|tagging\s+plan|tracking\s+plan|plan\s+de\s+(taggage|marquage)|consent\s+mode|matomo|adobe\s+analytics"
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
        $keywordOverlap = -not [string]::IsNullOrWhiteSpace($keywordText) -and -not [string]::IsNullOrWhiteSpace($existingKeywords) -and ($keywordText -match "google|gtm|ga4|piano|contentsquare|tracking|tagging|cro") -and ($existingKeywords -match "google|gtm|ga4|piano|contentsquare|tracking|tagging|cro")

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

function Merge-JobsWithTracker {
    param(
        [object[]]$CurrentRows,
        [object[]]$ExistingRows,
        [string]$Path
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

    $backupPath = Backup-TrackerFile -Path $Path

    return @{
        TrackerRows = @($trackerRows)
        CurrentRows = @($currentByKey.Values)
        RemovedCount = $removedCount
        DuplicateCount = $duplicateCount
        PreservedAppliedCount = $preservedAppliedCount
        BackupPath = $backupPath
    }
}

function Invoke-TextRequest {
    param(
        [string]$Url,
        [hashtable]$Headers = @{},
        [int]$TimeoutSec = 30
    )

    $mergedHeaders = @{
        "User-Agent"      = $BrowserUserAgent
        "Accept"          = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
        "Accept-Language" = "fr-FR,fr;q=0.9,en;q=0.8"
    }

    foreach ($key in $Headers.Keys) {
        $mergedHeaders[$key] = $Headers[$key]
    }

    $response = Invoke-WebRequest -Uri $Url -UseBasicParsing -Headers $mergedHeaders -TimeoutSec $TimeoutSec
    return [string]$response.Content
}

function Invoke-CurlTextRequest {
    param([string]$Url)

    $curl = Get-Command "curl.exe" -ErrorAction Stop
    $lines = & $curl.Source -L -s --compressed $Url `
        -H "User-Agent: $BrowserUserAgent" `
        -H "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8" `
        -H "Accept-Language: fr-FR,fr;q=0.9,en;q=0.8" `
        -H "Connection: keep-alive"

    if ($LASTEXITCODE -ne 0) {
        throw "curl.exe failed for $Url with exit code $LASTEXITCODE"
    }

    return ($lines -join "`n")
}

function Get-MetaContent {
    param(
        [string]$Html,
        [string]$Name
    )

    $escapedName = [regex]::Escape($Name)
    $pattern = "<meta[^>]+(?:property|name)=[""']$escapedName[""'][^>]+content=[""'](?<content>[^""']+)[""']"
    $match = [regex]::Match($Html, $pattern, [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $match.Success) {
        $pattern = "<meta[^>]+content=[""'](?<content>[^""']+)[""'][^>]+(?:property|name)=[""']$escapedName[""']"
        $match = [regex]::Match($Html, $pattern, [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    }

    if ($match.Success) {
        return ConvertFrom-HtmlAttribute $match.Groups["content"].Value
    }

    return ""
}

function Get-TitleFromHtml {
    param([string]$Html)

    $ogTitle = Get-MetaContent -Html $Html -Name "og:title"
    if (-not [string]::IsNullOrWhiteSpace($ogTitle)) {
        return ($ogTitle -replace "\s+-\s+Welcome to the Jungle.*$", "").Trim()
    }

    $titleMatch = [regex]::Match($Html, "(?is)<title[^>]*>(?<title>.*?)</title>")
    if ($titleMatch.Success) {
        return (ConvertFrom-HtmlText $titleMatch.Groups["title"].Value)
    }

    return ""
}

function Get-TitleFromWttjUrl {
    param([string]$Url)

    $slugMatch = [regex]::Match($Url, "/jobs/(?<slug>[^/?#]+)")
    if (-not $slugMatch.Success) {
        return $Url
    }

    $slug = $slugMatch.Groups["slug"].Value
    $slug = ($slug -split "_")[0]
    $slug = $slug -replace "-", " "
    return [Globalization.CultureInfo]::CurrentCulture.TextInfo.ToTitleCase($slug)
}

function ConvertFrom-SlugToTitle {
    param([AllowNull()][string]$Slug)

    if ([string]::IsNullOrWhiteSpace($Slug)) {
        return ""
    }

    $clean = $Slug -replace "[-_]+", " "
    $clean = ([regex]::Replace($clean, "\s+", " ")).Trim()
    return [Globalization.CultureInfo]::CurrentCulture.TextInfo.ToTitleCase($clean)
}

function Join-CleanTextParts {
    param([object[]]$Parts)

    $clean = New-Object System.Collections.Generic.List[string]
    foreach ($part in @($Parts)) {
        if ($null -eq $part) {
            continue
        }

        $text = ConvertFrom-HtmlText ([string]$part)
        if ([string]::IsNullOrWhiteSpace($text)) {
            continue
        }

        if (-not $clean.Contains($text)) {
            $clean.Add($text) | Out-Null
        }
    }

    return ($clean.ToArray() -join ", ")
}

function Get-ObjectPropertyValue {
    param(
        [AllowNull()]$Object,
        [string[]]$Names
    )

    if ($null -eq $Object) {
        return $null
    }

    $propertyNames = @($Object.PSObject.Properties.Name)
    foreach ($name in $Names) {
        if ($propertyNames -contains $name) {
            return $Object.PSObject.Properties[$name].Value
        }
    }

    return $null
}

function ConvertTo-LocationText {
    param(
        [AllowNull()]$Value,
        [int]$Depth = 0
    )

    if ($null -eq $Value -or $Depth -gt 3) {
        return ""
    }

    if ($Value -is [string]) {
        return ConvertFrom-HtmlText $Value
    }

    if ($Value -is [System.Collections.IEnumerable]) {
        $locations = foreach ($item in @($Value)) {
            ConvertTo-LocationText -Value $item -Depth ($Depth + 1)
        }
        return Join-CleanTextParts $locations
    }

    $city = Get-ObjectPropertyValue -Object $Value -Names @("city", "locality", "addressLocality", "town")
    $region = Get-ObjectPropertyValue -Object $Value -Names @("region", "state", "addressRegion", "department")
    $country = Get-ObjectPropertyValue -Object $Value -Names @("country", "countryCode", "addressCountry")
    $directLocation = Join-CleanTextParts @($city, $region, $country)
    if (-not [string]::IsNullOrWhiteSpace($directLocation)) {
        return $directLocation
    }

    foreach ($nestedName in @("location", "locations", "address", "addresses", "office", "offices", "place", "places")) {
        $nestedValue = Get-ObjectPropertyValue -Object $Value -Names @($nestedName)
        $nestedLocation = ConvertTo-LocationText -Value $nestedValue -Depth ($Depth + 1)
        if (-not [string]::IsNullOrWhiteSpace($nestedLocation)) {
            return $nestedLocation
        }
    }

    $name = Get-ObjectPropertyValue -Object $Value -Names @("name", "label", "formatted", "full_address", "fullAddress")
    return ConvertTo-LocationText -Value $name -Depth ($Depth + 1)
}

function Get-LocationFromStructuredHtml {
    param([AllowNull()][string]$Html)

    if ([string]::IsNullOrWhiteSpace($Html)) {
        return ""
    }

    $cityMatch = [regex]::Match($Html, '(?i)"addressLocality"\s*:\s*"(?<value>[^"]+)"')
    $regionMatch = [regex]::Match($Html, '(?i)"addressRegion"\s*:\s*"(?<value>[^"]+)"')
    $countryMatch = [regex]::Match($Html, '(?i)"addressCountry"\s*:\s*"(?<value>[^"]+)"')

    return Join-CleanTextParts @(
        $(if ($cityMatch.Success) { $cityMatch.Groups["value"].Value }),
        $(if ($regionMatch.Success) { $regionMatch.Groups["value"].Value }),
        $(if ($countryMatch.Success) { $countryMatch.Groups["value"].Value })
    )
}

function Get-LocationFromText {
    param([AllowNull()][string]$Text)

    $clean = ConvertFrom-HtmlText $Text
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return ""
    }

    $match = [regex]::Match($clean, "(?i)(?:\ba\s+|\b\xE0\s+|\bin\s+)(?<location>\p{L}[\p{L}\p{M}' -]{2,})(?:$|[,.])")
    if ($match.Success) {
        return $match.Groups["location"].Value.Trim()
    }

    return ""
}

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

    $locationMatch = [regex]::Match($path, "/jobs/[^/?#]*_(?<location>[^/_?#]+)$", [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $locationMatch.Success) {
        return ""
    }

    $locationSlug = $locationMatch.Groups["location"].Value
    if ($locationSlug -match "^(h|f|m|x|nb|stage|internship|cdi|cdd)$") {
        return ""
    }

    return ConvertFrom-SlugToTitle $locationSlug
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
        return "https://www.welcometothejungle.com/fr/jobs/{0}" -f $Job.reference
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
    $url = "https://www.welcomekit.co/api/v1/external/jobs/{0}?{1}" -f [Uri]::EscapeDataString($Reference), (ConvertTo-QueryString $params)

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
        $url = "https://www.welcomekit.co/api/v1/external/jobs/all?{0}" -f (ConvertTo-QueryString $params)

        try {
            $jobs = Invoke-RestMethod -Uri $url -Headers $headers -Method Get -TimeoutSec 45
        }
        catch {
            Write-Warning ("WelcomeKit API call failed on page {0}: {1}" -f $page, $_.Exception.Message)
            break
        }

        $jobArray = @($jobs)
        if ($jobArray.Count -eq 0) {
            break
        }

        foreach ($job in $jobArray) {
            $publishedAt = ConvertTo-DateTimeOffsetOrNull $job.published_at
            if (-not (Test-IsRecent $publishedAt)) {
                continue
            }

            $combined = ConvertFrom-HtmlText ("{0} {1} {2} {3}" -f $job.name, $job.profile, $job.description, $job.company_description)
            $match = Get-JobMatch -Title ([string]$job.name) -Text $combined
            if (-not $match.IsMatch) {
                continue
            }

            $jobUrl = Get-WelcomeKitJobUrl $job
            $jobDetails = $null
            if ($job.PSObject.Properties.Name -contains "reference") {
                $jobDetails = Get-WelcomeKitJobDetails -Reference ([string]$job.reference) -Headers $headers
            }
            if ($null -ne $jobDetails) {
                $jobUrl = Get-WelcomeKitJobUrl $jobDetails
                $job = $jobDetails
            }

            $companyName = Get-WelcomeKitCompanyName -Job $job -JobUrl $jobUrl
            $jobLocation = Get-WelcomeKitLocation -Job $job -JobUrl $jobUrl
            $contractType = Get-ContractTypeFromText -Text $combined -RawContractType ([string]$job.contract_type)
            $result = New-JobResult -Title ([string]$job.name) -CompanyName $companyName -JobLocation $jobLocation -ContractType $contractType -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url $jobUrl -Platform "Welcome to the Jungle" -PublishedAt $publishedAt
            if ($null -ne $result) {
                $results.Add($result) | Out-Null
            }
        }

        if ($jobArray.Count -lt 100) {
            break
        }
    }

    Write-RunStatus ("WelcomeKit API complete: {0} matching jobs." -f $results.Count)
    return $results.ToArray()
}

function Get-WttjPublicFallbackJobs {
    if ($DisableWttjPublicFallback) {
        return @()
    }

    Set-RunWindowTitle "Analytics Job Crawler - WTTJ"
    Write-RunStatus "Collecting Welcome to the Jungle jobs from public sitemaps..."
    $results = New-Object System.Collections.Generic.List[object]
    $candidateSeen = @{}
    $candidates = New-Object System.Collections.Generic.List[object]

    try {
        $indexXml = Invoke-CurlTextRequest "https://www.welcometothejungle.com/sitemaps/index.xml.gz"
    }
    catch {
        Write-Warning ("Could not read WTTJ sitemap index: {0}" -f $_.Exception.Message)
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
            $xml = Invoke-CurlTextRequest $sitemapUrl
        }
        catch {
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

            $lastmod = ConvertTo-DateTimeOffsetOrNull $lastmodMatch.Groups["lastmod"].Value
            if (-not (Test-IsRecent $lastmod)) {
                continue
            }

            if ($loc -notmatch $WttjUrlCandidatePattern) {
                continue
            }

            $candidateSeen[$loc] = $true
            $candidates.Add([PSCustomObject]@{
                Url       = $loc
                LastMod   = $lastmod
                SlugTitle = Get-TitleFromWttjUrl $loc
                Score     = Get-WttjCandidateScore $loc
            }) | Out-Null
        }
    }

    $selectedCandidates = $candidates |
        Sort-Object -Property @{ Expression = "Score"; Descending = $true }, @{ Expression = "LastMod"; Descending = $true } |
        Select-Object -First $MaxWttjCandidatePages

    $selectedCandidateCount = @($selectedCandidates).Count
    Write-RunStatus ("WTTJ candidates selected: {0} page(s) to inspect." -f $selectedCandidateCount)
    $count = 0
    foreach ($candidate in $selectedCandidates) {
        $count++
        Write-CountProgress -Activity "WTTJ candidate pages" -Current $count -Total $selectedCandidateCount -Found $results.Count -Every 10

        $urlText = "{0} {1}" -f $candidate.Url, (($candidate.SlugTitle -replace "[-_]", " "))
        $urlMatchResult = Get-JobMatch -Title $candidate.SlugTitle -Text $urlText
        $urlOnlyMatch = $urlMatchResult.IsMatch

        try {
            $html = Invoke-CurlTextRequest $candidate.Url
        }
        catch {
            if ($urlOnlyMatch) {
                $jobLocation = Get-WttjLocation -Html "" -Url $candidate.Url -Title $candidate.SlugTitle
                $fallbackResult = New-JobResult -Title $candidate.SlugTitle -CompanyName (Get-WttjCompanyNameFromUrl $candidate.Url) -JobLocation $jobLocation -ContractType (Get-ContractTypeFromText -Text $urlText) -MatchScore $urlMatchResult.Score -MatchLevel $urlMatchResult.Level -MatchedKeywords $urlMatchResult.Keywords -Url $candidate.Url -Platform "Welcome to the Jungle" -PublishedAt $candidate.LastMod
                if ($null -ne $fallbackResult) {
                    $results.Add($fallbackResult) | Out-Null
                }
            }
            continue
        }

        if ($html -match "(?i)<title>ERROR: The request could not be satisfied</title>|AwsWafIntegration|challenge-container") {
            if ($urlOnlyMatch) {
                $jobLocation = Get-WttjLocation -Html $html -Url $candidate.Url -Title $candidate.SlugTitle
                $fallbackResult = New-JobResult -Title $candidate.SlugTitle -CompanyName (Get-WttjCompanyNameFromUrl $candidate.Url) -JobLocation $jobLocation -ContractType (Get-ContractTypeFromText -Text $urlText) -MatchScore $urlMatchResult.Score -MatchLevel $urlMatchResult.Level -MatchedKeywords $urlMatchResult.Keywords -Url $candidate.Url -Platform "Welcome to the Jungle" -PublishedAt $candidate.LastMod
                if ($null -ne $fallbackResult) {
                    $results.Add($fallbackResult) | Out-Null
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
            continue
        }

        $pageText = ConvertFrom-HtmlText $html
        $combined = "{0} {1} {2}" -f $title, $candidate.Url, $pageText
        $match = Get-JobMatch -Title $title -Text $combined
        if (-not $match.IsMatch -and -not $urlOnlyMatch) {
            continue
        }
        if (-not $match.IsMatch) {
            $match = $urlMatchResult
        }

        $jobLocation = Get-WttjLocation -Html $html -Url $candidate.Url -Title $title
        $result = New-JobResult -Title $title -CompanyName (Get-WttjCompanyNameFromUrl $candidate.Url) -JobLocation $jobLocation -ContractType (Get-ContractTypeFromText -Text $combined) -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url $candidate.Url -Platform "Welcome to the Jungle" -PublishedAt $publishedAt
        if ($null -ne $result) {
            $results.Add($result) | Out-Null
        }

        Start-Sleep -Milliseconds 250
    }

    Write-RunStatus ("WTTJ public fallback complete: {0} matching jobs." -f $results.Count)
    return $results.ToArray()
}

function Get-LinkedInJobs {
    Set-RunWindowTitle "Analytics Job Crawler - LinkedIn"
    Write-RunStatus "Collecting LinkedIn jobs from public guest endpoints..."
    Write-RunStatus ("LinkedIn plan: {0} search query/queries, up to {1} page(s) each. Detail pages are fetched slowly to reduce rate-limit errors." -f $LinkedInQueries.Count, $MaxLinkedInSearchPages)
    $results = New-Object System.Collections.Generic.List[object]
    $detailCache = @{}
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
            $url = "https://www.linkedin.com/jobs-guest/jobs/api/seeMoreJobPostings/search?{0}" -f (ConvertTo-QueryString $params)

            try {
                $html = Invoke-TextRequest $url -Headers @{ "Accept" = "text/html,*/*" } -TimeoutSec 30
            }
            catch {
                Start-Sleep -Seconds 8
                try {
                    $html = Invoke-TextRequest $url -Headers @{ "Accept" = "text/html,*/*" } -TimeoutSec 30
                }
                catch {
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
            Write-RunStatus ("LinkedIn query {0}/{1}, page {2}/{3}: {4} card(s) found; {5} matches so far." -f $queryIndex, $LinkedInQueries.Count, ($page + 1), $MaxLinkedInSearchPages, $cards.Count, $results.Count)

            $cardIndex = 0
            foreach ($card in $cards) {
                $cardIndex++
                Write-CountProgress -Activity ("LinkedIn details q{0}/{1} p{2}/{3}" -f $queryIndex, $LinkedInQueries.Count, ($page + 1), $MaxLinkedInSearchPages) -Current $cardIndex -Total $cards.Count -Found $results.Count -Every 10
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
                    continue
                }

                if (-not $detailCache.ContainsKey($id)) {
                    try {
                        $detailCache[$id] = Invoke-TextRequest "https://www.linkedin.com/jobs-guest/jobs/api/jobPosting/$id" -Headers @{ "Accept" = "text/html,*/*" } -TimeoutSec 30
                    }
                    catch {
                        Start-Sleep -Seconds 5
                        try {
                            $detailCache[$id] = Invoke-TextRequest "https://www.linkedin.com/jobs-guest/jobs/api/jobPosting/$id" -Headers @{ "Accept" = "text/html,*/*" } -TimeoutSec 30
                        }
                        catch {
                            $detailCache[$id] = ""
                        }
                    }
                    Start-Sleep -Milliseconds $LinkedInDelayMilliseconds
                }

                $detailText = ConvertFrom-HtmlText $detailCache[$id]
                $combined = "{0} {1} {2}" -f $title, $jobUrl, $detailText
                $match = Get-JobMatch -Title $title -Text $combined
                if (-not $match.IsMatch) {
                    continue
                }

                if ([string]::IsNullOrWhiteSpace($companyName)) {
                    $detailCompanyMatch = [regex]::Match($detailCache[$id], '(?is)<a[^>]*class="[^"]*topcard__org-name-link[^"]*"[^>]*>(?<company>.*?)</a>')
                    if ($detailCompanyMatch.Success) {
                        $companyName = ConvertFrom-HtmlText $detailCompanyMatch.Groups["company"].Value
                    }
                }
                if ([string]::IsNullOrWhiteSpace($jobLocation)) {
                    $jobLocation = Get-LinkedInLocationFromHtml $detailCache[$id]
                }

                $contractType = Get-LinkedInContractType -Title $title -DetailText $detailText
                $result = New-JobResult -Title $title -CompanyName $companyName -JobLocation $jobLocation -ContractType $contractType -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url $jobUrl -Platform "LinkedIn" -PublishedAt $publishedAt
                if ($null -ne $result) {
                    $results.Add($result) | Out-Null
                }
            }

            Start-Sleep -Milliseconds $LinkedInDelayMilliseconds
        }
    }

    Write-RunStatus ("LinkedIn complete: {0} matching jobs." -f $results.Count)
    return $results.ToArray()
}

Set-RunWindowTitle "Analytics Job Crawler - Starting"
Write-RunStatus ("Starting crawl for jobs published since {0}." -f $CutoffDate)
Write-RunStatus ("Tracker file: {0}" -f $TrackerPath)

$existingTrackerRows = @(Import-TrackerRows -Path $TrackerPath)
Write-RunStatus ("Loaded {0} existing tracker row(s)." -f $existingTrackerRows.Count)

$allResults = New-Object System.Collections.Generic.List[object]

if (-not $SkipWttj) {
    $welcomeKitResults = @(Get-WelcomeKitJobs)
    foreach ($result in $welcomeKitResults) {
        $allResults.Add($result) | Out-Null
    }

    if ($welcomeKitResults.Count -eq 0 -and [string]::IsNullOrWhiteSpace($WelcomeKitApiKey)) {
        foreach ($result in @(Get-WttjPublicFallbackJobs)) {
            $allResults.Add($result) | Out-Null
        }
    }
}

if (-not $SkipLinkedIn) {
    foreach ($result in @(Get-LinkedInJobs)) {
        $allResults.Add($result) | Out-Null
    }
}

$sortedCrawlResults = $allResults |
    Sort-Object -Property @{ Expression = "match_score"; Descending = $true }, @{ Expression = "published_date"; Descending = $true }, platform, job_title
$filteredCrawlResults = @($sortedCrawlResults | Where-Object { -not (Test-IsExcludedContractType (Get-RowValue -Row $_ -Name "contract_type")) })
$excludedContractCount = @($sortedCrawlResults).Count - @($filteredCrawlResults).Count

if ($excludedContractCount -gt 0) {
    Write-RunStatus ("Excluded {0} CDD/apprenticeship/internship job(s) from this crawl." -f $excludedContractCount)
}

$feedbackAdjustedResults = @(Apply-FeedbackScoring -Rows $filteredCrawlResults -ExistingRows $existingTrackerRows)
$mergeResult = Merge-JobsWithTracker -CurrentRows $feedbackAdjustedResults -ExistingRows $existingTrackerRows -Path $TrackerPath
$finalResults = @($mergeResult.TrackerRows)

$crawlSummary = @{
    TotalMatched = @($sortedCrawlResults).Count
    ExcludedContractCount = $excludedContractCount
    CurrentCount = @($mergeResult.CurrentRows).Count
    TrackerCount = @($finalResults).Count
    DuplicateCount = $mergeResult.DuplicateCount
    RemovedCount = $mergeResult.RemovedCount
    PreservedAppliedCount = $mergeResult.PreservedAppliedCount
    BackupPath = $mergeResult.BackupPath
}
Export-TrackerWorkbook -Rows $finalResults -Path $TrackerPath -Summary $crawlSummary

Set-RunWindowTitle "Analytics Job Crawler - Finished"
Write-Host ""
Write-RunStatus ("Wrote tracker with {0} row(s): {1} current job(s), {2} preserved application row(s), {3} removed by retention." -f @($finalResults).Count, @($mergeResult.CurrentRows).Count, $mergeResult.PreservedAppliedCount, $mergeResult.RemovedCount)
if (-not [string]::IsNullOrWhiteSpace($mergeResult.BackupPath)) {
    Write-RunStatus ("Backup: {0}" -f $mergeResult.BackupPath)
}
Write-RunStatus ("Tracker: {0}" -f (Resolve-Path $TrackerPath).Path)

if (@($finalResults).Count -gt 0) {
    $finalResults | Format-Table -AutoSize
}
