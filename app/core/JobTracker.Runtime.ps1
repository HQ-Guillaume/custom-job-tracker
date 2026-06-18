function Repair-DisplayText {
    param([AllowNull()][string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ""
    }

    $clean = [string]$Text
    $clean = $clean.Replace(([string][char]0x00A0), " ")

    $mojibakeScore = {
        param([string]$Value)

        if ([string]::IsNullOrWhiteSpace($Value)) {
            return 0
        }

        $score = 0
        foreach ($marker in @(
            [string][char]0x00C2,
            [string][char]0x00C3,
            [string][char]0x00E2,
            [string][char]0x251C,
            [string][char]0xFFFD
        )) {
            $score += [regex]::Matches($Value, [regex]::Escape($marker)).Count
        }

        return $score
    }

    if (& $mojibakeScore $clean) {
        foreach ($codePage in @(1252, 850, 437)) {
            try {
                $sourceEncoding = [Text.Encoding]::GetEncoding($codePage)
                $decoded = [Text.Encoding]::UTF8.GetString($sourceEncoding.GetBytes($clean))
                if (-not [string]::IsNullOrWhiteSpace($decoded) -and (& $mojibakeScore $decoded) -lt (& $mojibakeScore $clean)) {
                    $clean = $decoded
                    break
                }
            }
            catch {
            }
        }
    }

    $mojibakeReplacements = @(
        @{ From = ([string][char]0x00C3 + [string][char]0x0080); To = ([string][char]0x00C0) },
        @{ From = ([string][char]0x00C3 + [string][char]0x0087); To = ([string][char]0x00C7) },
        @{ From = ([string][char]0x00C3 + [string][char]0x0088); To = ([string][char]0x00C8) },
        @{ From = ([string][char]0x00C3 + [string][char]0x0089); To = ([string][char]0x00C9) },
        @{ From = ([string][char]0x00C3 + [string][char]0x008A); To = ([string][char]0x00CA) },
        @{ From = ([string][char]0x00C3 + [string][char]0x00A0); To = ([string][char]0x00E0) },
        @{ From = ([string][char]0x00C3 + [string][char]0x00A2); To = ([string][char]0x00E2) },
        @{ From = ([string][char]0x00C3 + [string][char]0x00A7); To = ([string][char]0x00E7) },
        @{ From = ([string][char]0x00C3 + [string][char]0x00A8); To = ([string][char]0x00E8) },
        @{ From = ([string][char]0x00C3 + [string][char]0x00A9); To = ([string][char]0x00E9) },
        @{ From = ([string][char]0x00C3 + [string][char]0x00AA); To = ([string][char]0x00EA) },
        @{ From = ([string][char]0x00C3 + [string][char]0x00AB); To = ([string][char]0x00EB) },
        @{ From = ([string][char]0x00C3 + [string][char]0x00AE); To = ([string][char]0x00EE) },
        @{ From = ([string][char]0x00C3 + [string][char]0x00AF); To = ([string][char]0x00EF) },
        @{ From = ([string][char]0x00C3 + [string][char]0x00B4); To = ([string][char]0x00F4) },
        @{ From = ([string][char]0x00C3 + [string][char]0x00B9); To = ([string][char]0x00F9) },
        @{ From = ([string][char]0x00C3 + [string][char]0x00BB); To = ([string][char]0x00FB) },
        @{ From = ([string][char]0x251C + [string][char]0x00A1); To = ([string][char]0x00E0) },
        @{ From = ([string][char]0x251C + [string][char]0x00E1); To = ([string][char]0x00E0) },
        @{ From = ([string][char]0x251C + [string][char]0x00A7); To = ([string][char]0x00E7) },
        @{ From = ([string][char]0x251C + [string][char]0x00A8); To = ([string][char]0x00E8) },
        @{ From = ([string][char]0x251C + [string][char]0x00A9); To = ([string][char]0x00E9) },
        @{ From = ([string][char]0x251C + [string][char]0x00AE); To = ([string][char]0x00E9) },
        @{ From = ([string][char]0x251C + [string][char]0x00AA); To = ([string][char]0x00EA) },
        @{ From = ([string][char]0x00E2 + [string][char]0x0080 + [string][char]0x0099); To = "'" },
        @{ From = ([string][char]0x00E2 + [string][char]0x0080 + [string][char]0x0093); To = "-" }
    )

    foreach ($replacement in $mojibakeReplacements) {
        $clean = $clean.Replace([string]$replacement.From, [string]$replacement.To)
    }

    $clean = $clean.Replace(([string][char]0x2018), "'")
    $clean = $clean.Replace(([string][char]0x2019), "'")
    $clean = $clean.Replace(([string][char]0x201C), '"')
    $clean = $clean.Replace(([string][char]0x201D), '"')
    $clean = $clean.Replace(([string][char]0x2013), "-")
    $clean = $clean.Replace(([string][char]0x2014), "-")

    return ([regex]::Replace($clean, "\s+", " ")).Trim()
}

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

function Start-SourceStats {
    param([string]$Source)

    return [ordered]@{
        Source          = $Source
        StartedAt       = [DateTimeOffset]::Now
        FinishedAt      = $null
        DurationSeconds = 0
        SearchRequests  = 0
        DetailRequests  = 0
        CacheHits       = 0
        Candidates      = 0
        SelectedDetails = 0
        SkippedOld      = 0
        SkippedContract = 0
        SkippedNoMatch  = 0
        SkippedByCap    = 0
        Errors          = 0
        Matches         = 0
        Notes           = ""
    }
}

function Add-SourceMetric {
    param(
        [AllowNull()]$Stats,
        [string]$Name,
        [int]$Amount = 1
    )

    if ($null -eq $Stats -or -not $Stats.Contains($Name)) {
        return
    }

    $Stats[$Name] = [int]$Stats[$Name] + $Amount
}

function Complete-SourceStats {
    param([AllowNull()]$Stats)

    if ($null -eq $Stats) {
        return
    }

    $Stats["FinishedAt"] = [DateTimeOffset]::Now
    $Stats["DurationSeconds"] = [int][Math]::Round((([DateTimeOffset]$Stats["FinishedAt"]) - ([DateTimeOffset]$Stats["StartedAt"])).TotalSeconds, 0)
    $script:SourceRunStats.Add([PSCustomObject]$Stats) | Out-Null
    Write-RunStatus ("{0} diagnostics: {1}s, search {2}, details {3}, cache hits {4}, candidates {5}, selected {6}, old {7}, contract {8}, no-match {9}, cap {10}, errors {11}, matches {12}." -f `
            $Stats["Source"],
            $Stats["DurationSeconds"],
            $Stats["SearchRequests"],
            $Stats["DetailRequests"],
            $Stats["CacheHits"],
            $Stats["Candidates"],
            $Stats["SelectedDetails"],
            $Stats["SkippedOld"],
            $Stats["SkippedContract"],
            $Stats["SkippedNoMatch"],
            $Stats["SkippedByCap"],
            $Stats["Errors"],
            $Stats["Matches"])
}

function Get-SourceStatsSummaryText {
    if ($script:SourceRunStats.Count -eq 0) {
        return ""
    }

    $parts = foreach ($stat in @($script:SourceRunStats.ToArray())) {
        "{0}: {1}s, {2} match(es), {3} candidate(s), {4} detail(s), {5} cap-skip, {6} cache-hit(s)" -f `
            $stat.Source,
            $stat.DurationSeconds,
            $stat.Matches,
            $stat.Candidates,
            $stat.DetailRequests,
            $stat.SkippedByCap,
            $stat.CacheHits
    }

    return ($parts -join " | ")
}

function Get-JobCrawlerAcceptLanguage {
    if (Get-Variable -Name JobCrawlerRuntimeConfig -Scope Script -ErrorAction SilentlyContinue) {
        return [string](Get-ConfigPathValue -Object $script:JobCrawlerRuntimeConfig -Path "http.accept_language" -DefaultValue "fr-FR,fr;q=0.9,en;q=0.8")
    }

    return "fr-FR,fr;q=0.9,en;q=0.8"
}

function Get-JobCrawlerHttpRetryPolicy {
    $maxRetries = 2
    $retryDelayMs = 1200
    $backoff = 2.0
    if (Get-Variable -Name JobCrawlerRuntimeConfig -Scope Script -ErrorAction SilentlyContinue) {
        $maxRetries = [int](Get-ConfigPathValue -Object $script:JobCrawlerRuntimeConfig -Path "http.max_retries" -DefaultValue $maxRetries)
        $retryDelayMs = [int](Get-ConfigPathValue -Object $script:JobCrawlerRuntimeConfig -Path "http.retry_delay_ms" -DefaultValue $retryDelayMs)
        $backoff = [double](Get-ConfigPathValue -Object $script:JobCrawlerRuntimeConfig -Path "http.retry_backoff_multiplier" -DefaultValue $backoff)
    }

    return [PSCustomObject]@{
        MaxAttempts = [Math]::Max(1, $maxRetries + 1)
        DelayMs = [Math]::Max(0, $retryDelayMs)
        BackoffMultiplier = [Math]::Max(1.0, $backoff)
    }
}

function Invoke-JobCrawlerRetry {
    param(
        [scriptblock]$Operation,
        [string]$Description = "HTTP request"
    )

    $policy = Get-JobCrawlerHttpRetryPolicy
    $delay = [int]$policy.DelayMs
    for ($attempt = 1; $attempt -le [int]$policy.MaxAttempts; $attempt++) {
        try {
            return & $Operation
        }
        catch {
            if ($attempt -ge [int]$policy.MaxAttempts) {
                throw
            }

            Write-RunStatus ("{0} failed on attempt {1}/{2}: {3}. Retrying..." -f $Description, $attempt, $policy.MaxAttempts, $_.Exception.Message) "WARN"
            if ($delay -gt 0) {
                Start-Sleep -Milliseconds $delay
                $delay = [int]([double]$delay * [double]$policy.BackoffMultiplier)
            }
        }
    }
}

function Invoke-JobCrawlerCachePrune {
    param(
        [string]$Path = $CacheDirectory,
        [bool]$Enabled = $true
    )

    if (-not $Enabled -or [string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        return [PSCustomObject]@{ RemovedFiles = 0; RemovedBytes = 0; RemainingBytes = 0 }
    }

    $maxAgeDays = [int](Get-ConfigPathValue -Object $script:JobCrawlerRuntimeConfig -Path "defaults.cache_max_age_days" -DefaultValue 30)
    $maxMb = [double](Get-ConfigPathValue -Object $script:JobCrawlerRuntimeConfig -Path "defaults.cache_max_mb" -DefaultValue 250)
    $removedFiles = 0
    $removedBytes = [int64]0
    $cutoffDate = (Get-Date).AddDays(-[Math]::Abs($maxAgeDays))

    $files = @(Get-ChildItem -LiteralPath $Path -Recurse -File -ErrorAction SilentlyContinue)
    foreach ($file in @($files | Where-Object { $_.LastWriteTime -lt $cutoffDate })) {
        $length = [int64]$file.Length
        Remove-Item -LiteralPath $file.FullName -Force -ErrorAction SilentlyContinue
        $removedFiles++
        $removedBytes += $length
    }

    $files = @(Get-ChildItem -LiteralPath $Path -Recurse -File -ErrorAction SilentlyContinue)
    $measurement = $files | Measure-Object Length -Sum
    $totalBytes = [int64]0
    if ($null -ne $measurement -and $null -ne $measurement.Sum) {
        $totalBytes = [int64]$measurement.Sum
    }
    $maxBytes = [int64]($maxMb * 1MB)
    if ($maxBytes -gt 0 -and $totalBytes -gt $maxBytes) {
        foreach ($file in @($files | Sort-Object LastWriteTime)) {
            if ($totalBytes -le $maxBytes) {
                break
            }
            $length = [int64]$file.Length
            Remove-Item -LiteralPath $file.FullName -Force -ErrorAction SilentlyContinue
            $removedFiles++
            $removedBytes += $length
            $totalBytes -= $length
        }
    }

    return [PSCustomObject]@{
        RemovedFiles = $removedFiles
        RemovedBytes = $removedBytes
        RemainingBytes = [int64]([Math]::Max(0, $totalBytes))
    }
}

function Write-RunHistoryEntry {
    param(
        [hashtable]$Summary,
        [string]$Path,
        [int]$MaxEntries = 250
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return
    }

    $directory = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    $sourceStats = @()
    if (Get-Variable -Name SourceRunStats -Scope Script -ErrorAction SilentlyContinue) {
        $sourceStats = @($script:SourceRunStats.ToArray())
    }

    $entry = [ordered]@{
        run_stamp = $RunStamp
        run_date = $RunDate
        created_at = ([DateTimeOffset]::Now.ToString("o"))
        profile = $(if ($Summary.ContainsKey("Profile")) { $Summary.Profile } else { "" })
        crawl_mode = $CrawlMode
        dry_run = $Summary.DryRun
        diagnostic_mode = $Summary.DiagnosticMode
        total_matched = $Summary.TotalMatched
        current_count = $Summary.CurrentCount
        tracker_count = $Summary.TrackerCount
        duplicate_count = $Summary.DuplicateCount
        removed_count = $Summary.RemovedCount
        preserved_application_count = $Summary.PreservedAppliedCount
        excluded_contract_count = $Summary.ExcludedContractCount
        source_stats = $sourceStats
    }

    Add-Content -LiteralPath $Path -Value (($entry | ConvertTo-Json -Depth 8 -Compress)) -Encoding UTF8

    if ($MaxEntries -gt 0 -and (Test-Path -LiteralPath $Path)) {
        $lines = @(Get-Content -LiteralPath $Path -Encoding UTF8)
        if ($lines.Count -gt $MaxEntries) {
            $lines | Select-Object -Last $MaxEntries | Set-Content -LiteralPath $Path -Encoding UTF8
        }
    }
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
    return Repair-DisplayText $text
}

function ConvertFrom-HtmlAttribute {
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ""
    }

    return Repair-DisplayText ([System.Net.WebUtility]::HtmlDecode($Value))
}

function Get-HtmlAttributeValue {
    param(
        [AllowNull()][string]$Html,
        [string]$Name
    )

    if ([string]::IsNullOrWhiteSpace($Html) -or [string]::IsNullOrWhiteSpace($Name)) {
        return ""
    }

    $pattern = "(?is)\b{0}\s*=\s*[""'](?<value>[^""']*)[""']" -f [regex]::Escape($Name)
    $match = [regex]::Match($Html, $pattern)
    if (-not $match.Success) {
        return ""
    }

    return ConvertFrom-HtmlAttribute $match.Groups["value"].Value
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

function ConvertTo-SafeCacheKey {
    param([string]$Value)

    $bytes = [Text.Encoding]::UTF8.GetBytes([string]$Value)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        return (($sha.ComputeHash($bytes) | ForEach-Object { $_.ToString("x2") }) -join "")
    }
    finally {
        $sha.Dispose()
    }
}

function Get-CacheFilePath {
    param(
        [string]$Scope,
        [string]$Key
    )

    $safeScope = ([regex]::Replace((ConvertTo-MatchText $Scope), "[^a-z0-9_-]+", "_")).Trim("_")
    if ([string]::IsNullOrWhiteSpace($safeScope)) {
        $safeScope = "default"
    }

    return Join-Path (Join-Path $CacheDirectory $safeScope) ("{0}.txt" -f (ConvertTo-SafeCacheKey $Key))
}

function Get-CachedText {
    param(
        [string]$Scope,
        [string]$Key,
        [int]$TtlHours = $CacheTtlHours
    )

    if ($DisableCache -or $TtlHours -le 0) {
        return $null
    }

    $path = Get-CacheFilePath -Scope $Scope -Key $Key
    if (-not (Test-Path -LiteralPath $path)) {
        return $null
    }

    $item = Get-Item -LiteralPath $path
    if ($item.LastWriteTime -lt (Get-Date).AddHours(-[Math]::Abs($TtlHours))) {
        return $null
    }

    return [IO.File]::ReadAllText($item.FullName, [Text.Encoding]::UTF8)
}

function Set-CachedText {
    param(
        [string]$Scope,
        [string]$Key,
        [AllowNull()][string]$Text
    )

    if ($DisableCache -or [string]::IsNullOrWhiteSpace($Text)) {
        return
    }

    $path = Get-CacheFilePath -Scope $Scope -Key $Key
    $directory = Split-Path -Parent $path
    if (-not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Force -Path $directory | Out-Null
    }

    [IO.File]::WriteAllText($path, [string]$Text, [Text.Encoding]::UTF8)
}

function ConvertTo-AbsoluteUrl {
    param(
        [string]$BaseUrl,
        [string]$Href
    )

    $cleanHref = ConvertFrom-HtmlAttribute $Href
    if ([string]::IsNullOrWhiteSpace($cleanHref)) {
        return ""
    }

    if ($cleanHref -match "^https?://") {
        return ConvertTo-CleanUrl $cleanHref
    }

    try {
        $baseUri = [Uri]::new($BaseUrl)
        return ConvertTo-CleanUrl ([Uri]::new($baseUri, $cleanHref).AbsoluteUri)
    }
    catch {
        return ConvertTo-CleanUrl $cleanHref
    }
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

function ConvertFrom-FrenchRelativeDateText {
    param([AllowNull()][string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $null
    }

    $clean = ConvertTo-MatchText (ConvertFrom-HtmlText $Text)
    $now = [DateTimeOffset]::Now
    if ($clean -match "aujourd.?hui|quelques\s+(secondes|minutes)|a\s+l.?instant") {
        return $now
    }
    if ($clean -match "\bhier\b") {
        return $now.AddDays(-1)
    }

    $hoursMatch = [regex]::Match($clean, "il\s+y\s+a\s+(?<value>\d+)\s+h")
    if ($hoursMatch.Success) {
        return $now.AddHours(-[int]$hoursMatch.Groups["value"].Value)
    }

    $dayMatch = [regex]::Match($clean, "il\s+y\s+a\s+(?<value>\d+)\s+j")
    if ($dayMatch.Success) {
        return $now.AddDays(-[int]$dayMatch.Groups["value"].Value)
    }

    $weekMatch = [regex]::Match($clean, "il\s+y\s+a\s+(?<value>\d+)\s+sem")
    if ($weekMatch.Success) {
        return $now.AddDays(-7 * [int]$weekMatch.Groups["value"].Value)
    }

    $dateMatch = [regex]::Match($clean, "(?<day>\d{1,2})[/-](?<month>\d{1,2})[/-](?<year>\d{4})")
    if ($dateMatch.Success) {
        $dateText = "{0}-{1}-{2}" -f $dateMatch.Groups["year"].Value, $dateMatch.Groups["month"].Value.PadLeft(2, "0"), $dateMatch.Groups["day"].Value.PadLeft(2, "0")
        return ConvertTo-DateTimeOffsetOrNull $dateText
    }

    return $null
}

function Test-IsRecent {
    param([AllowNull()]$PublishedAt)

    if ($null -eq $PublishedAt) {
        return $false
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

    return ($null -ne $publishedDateValue -and $publishedDateValue -ge $Cutoff)
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

function Invoke-TextRequest {
    param(
        [string]$Url,
        [hashtable]$Headers = @{},
        [int]$TimeoutSec = 30
    )

    $mergedHeaders = @{
        "User-Agent"      = $BrowserUserAgent
        "Accept"          = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
        "Accept-Language" = Get-JobCrawlerAcceptLanguage
    }

    foreach ($key in $Headers.Keys) {
        $mergedHeaders[$key] = $Headers[$key]
    }

    $response = Invoke-JobCrawlerRetry -Description $Url -Operation {
        Invoke-WebRequest -Uri $Url -UseBasicParsing -Headers $mergedHeaders -TimeoutSec $TimeoutSec
    }
    return [string]$response.Content
}

function Invoke-CachedTextRequest {
    param(
        [string]$Url,
        [string]$CacheScope,
        [hashtable]$Headers = @{},
        [int]$TimeoutSec = 30,
        [AllowNull()]$Stats = $null
    )

    $cached = Get-CachedText -Scope $CacheScope -Key $Url
    if ($null -ne $cached) {
        Add-SourceMetric -Stats $Stats -Name "CacheHits"
        return $cached
    }

    Add-SourceMetric -Stats $Stats -Name "DetailRequests"
    $text = Invoke-TextRequest -Url $Url -Headers $Headers -TimeoutSec $TimeoutSec
    Set-CachedText -Scope $CacheScope -Key $Url -Text $text
    return $text
}

function Invoke-JsonPostRequest {
    param(
        [string]$Url,
        [AllowNull()]$Body = $null,
        [hashtable]$Headers = @{},
        [int]$TimeoutSec = 30
    )

    $mergedHeaders = @{
        "User-Agent"      = $BrowserUserAgent
        "Accept"          = "application/json, text/plain, */*"
        "Accept-Language" = Get-JobCrawlerAcceptLanguage
        "Content-Type"    = "application/json"
    }

    foreach ($key in $Headers.Keys) {
        $mergedHeaders[$key] = $Headers[$key]
    }

    $jsonBody = ""
    if ($null -ne $Body) {
        $jsonBody = $Body | ConvertTo-Json -Depth 12 -Compress
    }

    $response = Invoke-JobCrawlerRetry -Description $Url -Operation {
        Invoke-WebRequest -Uri $Url -UseBasicParsing -Method Post -Headers $mergedHeaders -Body $jsonBody -TimeoutSec $TimeoutSec
    }
    if ([string]::IsNullOrWhiteSpace([string]$response.Content)) {
        return $null
    }

    return ([string]$response.Content | ConvertFrom-Json)
}

function Invoke-CurlTextRequest {
    param([string]$Url)

    return Invoke-JobCrawlerRetry -Description $Url -Operation {
        $curl = Get-Command "curl.exe" -ErrorAction Stop
        $lines = & $curl.Source -L -s --compressed $Url `
            -H "User-Agent: $BrowserUserAgent" `
            -H "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8" `
            -H ("Accept-Language: {0}" -f (Get-JobCrawlerAcceptLanguage)) `
            -H "Connection: keep-alive"

        if ($LASTEXITCODE -ne 0) {
            throw "curl.exe failed for $Url with exit code $LASTEXITCODE"
        }

        return ($lines -join "`n")
    }
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
    return Repair-DisplayText ([Globalization.CultureInfo]::CurrentCulture.TextInfo.ToTitleCase($slug))
}

function ConvertFrom-SlugToTitle {
    param([AllowNull()][string]$Slug)

    if ([string]::IsNullOrWhiteSpace($Slug)) {
        return ""
    }

    $clean = $Slug -replace "[-_]+", " "
    $clean = ([regex]::Replace($clean, "\s+", " ")).Trim()
    return Repair-DisplayText ([Globalization.CultureInfo]::CurrentCulture.TextInfo.ToTitleCase($clean))
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

    if ($Object -is [Collections.IDictionary]) {
        foreach ($name in $Names) {
            if ($Object.Contains($name)) {
                return $Object[$name]
            }
        }

        return $null
    }

    $properties = @($Object.PSObject.Properties)
    foreach ($name in $Names) {
        $property = @($properties | Where-Object { $_.Name -eq $name } | Select-Object -First 1)
        if ($property.Count -gt 0) {
            return $property[0].Value
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
        $location = $match.Groups["location"].Value.Trim()
        if (-not (Test-IsJunkLocationText $location)) {
            return $location
        }
    }

    return ""
}

function Test-IsJunkLocationText {
    param([AllowNull()][string]$Location)

    if ([string]::IsNullOrWhiteSpace($Location)) {
        return $true
    }

    $raw = ([string]$Location).Trim()
    $clean = ConvertTo-MatchText $raw
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return $true
    }
    if ($clean -match "^(h|f|m|x|nb|stage|internship|cdi|cdd|full\s*time|permanent)$") {
        return $true
    }
    if ($raw -match "^[A-Za-z0-9]{5,14}$" -and $raw -match "\d" -and $raw -match "[A-Z]" -and $raw -match "[a-z]") {
        return $true
    }
    if ($clean -match "^[a-z0-9]{8,}$" -and $clean -match "\d") {
        return $true
    }

    return $false
}

