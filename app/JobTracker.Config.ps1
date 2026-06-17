# Configuration helpers for the crawler. Values that may reasonably change over
# time live in config/*.json; structural workbook internals stay in code.

function Read-JobCrawlerJsonConfig {
    param(
        [string]$Path,
        [AllowNull()]$DefaultValue = $null
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return $DefaultValue
    }

    try {
        return Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    catch {
        throw "Could not read config file '$Path': $($_.Exception.Message)"
    }
}

function Get-ConfigProperty {
    param(
        [AllowNull()]$Object,
        [string]$Name,
        [AllowNull()]$DefaultValue = $null
    )

    if ($null -eq $Object) {
        return $DefaultValue
    }
    if ($Object -is [hashtable] -and $Object.ContainsKey($Name)) {
        return $Object[$Name]
    }
    if ($Object.PSObject.Properties.Name -contains $Name) {
        return $Object.$Name
    }

    return $DefaultValue
}

function Get-ConfigPathValue {
    param(
        [AllowNull()]$Object,
        [string]$Path,
        [AllowNull()]$DefaultValue = $null
    )

    $current = $Object
    foreach ($part in ($Path -split "\.")) {
        $current = Get-ConfigProperty -Object $current -Name $part -DefaultValue $null
        if ($null -eq $current) {
            return $DefaultValue
        }
    }

    return $current
}

function Get-ConfigStringArray {
    param(
        [AllowNull()]$Value,
        [string[]]$DefaultValue = @()
    )

    if ($null -eq $Value) {
        return @($DefaultValue)
    }
    if ($Value -is [string]) {
        if ([string]::IsNullOrWhiteSpace($Value)) {
            return @()
        }
        return @([string]$Value)
    }

    return @($Value | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}

function Resolve-JobCrawlerPath {
    param(
        [string]$BasePath,
        [AllowNull()][string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return ""
    }
    if ([IO.Path]::IsPathRooted($Path)) {
        return $Path
    }

    return Join-Path $BasePath $Path
}

function ConvertTo-ConfigHashtable {
    param([AllowNull()]$Value)

    if ($null -eq $Value) {
        return $null
    }
    if ($Value -is [string]) {
        return $Value
    }
    if ($Value -is [System.Collections.IDictionary]) {
        $hash = [ordered]@{}
        foreach ($key in $Value.Keys) {
            $hash[[string]$key] = ConvertTo-ConfigHashtable $Value[$key]
        }
        return $hash
    }
    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [pscustomobject])) {
        $items = New-Object System.Collections.Generic.List[object]
        foreach ($item in $Value) {
            $items.Add((ConvertTo-ConfigHashtable $item)) | Out-Null
        }
        return @($items.ToArray())
    }
    if ($Value -is [pscustomobject]) {
        $hash = [ordered]@{}
        foreach ($property in $Value.PSObject.Properties) {
            $hash[$property.Name] = ConvertTo-ConfigHashtable $property.Value
        }
        return $hash
    }

    return $Value
}

function Merge-ConfigHashtable {
    param(
        [AllowNull()]$Base,
        [AllowNull()]$Override
    )

    if ($null -eq $Override) {
        return $Base
    }
    if ($null -eq $Base) {
        return $Override
    }
    if ($Base -is [System.Collections.IDictionary] -and $Override -is [System.Collections.IDictionary]) {
        $merged = [ordered]@{}
        foreach ($key in $Base.Keys) {
            $merged[$key] = $Base[$key]
        }
        foreach ($key in $Override.Keys) {
            if ($merged.Contains($key)) {
                $merged[$key] = Merge-ConfigHashtable -Base $merged[$key] -Override $Override[$key]
            }
            else {
                $merged[$key] = $Override[$key]
            }
        }
        return $merged
    }

    return $Override
}

function Merge-JobCrawlerConfigObjects {
    param(
        [AllowNull()]$Base,
        [AllowNull()]$Override
    )

    $baseHash = ConvertTo-ConfigHashtable $Base
    $overrideHash = ConvertTo-ConfigHashtable $Override
    $merged = Merge-ConfigHashtable -Base $baseHash -Override $overrideHash
    if ($null -eq $merged) {
        return [PSCustomObject]@{}
    }

    return ($merged | ConvertTo-Json -Depth 100 | ConvertFrom-Json)
}

function Read-JobCrawlerLayeredConfig {
    param(
        [string]$Root,
        [string]$Name,
        [AllowNull()]$DefaultValue = $null,
        [System.Collections.Generic.List[string]]$AppliedOverrides
    )

    $value = Read-JobCrawlerJsonConfig -Path (Join-Path $Root ("{0}.json" -f $Name)) -DefaultValue $DefaultValue
    $overridePaths = @(
        (Join-Path $Root ("local.{0}.json" -f $Name)),
        (Join-Path (Join-Path $Root "local") ("{0}.json" -f $Name))
    )

    foreach ($overridePath in $overridePaths) {
        if (Test-Path -LiteralPath $overridePath) {
            $overrideValue = Read-JobCrawlerJsonConfig -Path $overridePath -DefaultValue ([PSCustomObject]@{})
            $value = Merge-JobCrawlerConfigObjects -Base $value -Override $overrideValue
            if ($null -ne $AppliedOverrides) {
                $AppliedOverrides.Add($overridePath) | Out-Null
            }
        }
    }

    return $value
}

function Get-JobCrawlerConfig {
    param([string]$ConfigDirectory)

    $configRoot = Resolve-Path -LiteralPath $ConfigDirectory -ErrorAction SilentlyContinue
    if ($null -eq $configRoot) {
        throw "Config directory not found: $ConfigDirectory"
    }

    $root = $configRoot.Path
    $appliedOverrides = New-Object System.Collections.Generic.List[string]
    $runtimeConfig = Read-JobCrawlerLayeredConfig -Root $root -Name "runtime" -DefaultValue ([PSCustomObject]@{}) -AppliedOverrides $appliedOverrides
    $crawlModesConfig = Read-JobCrawlerLayeredConfig -Root $root -Name "crawl_modes" -DefaultValue ([PSCustomObject]@{}) -AppliedOverrides $appliedOverrides
    $sourcesConfig = Read-JobCrawlerLayeredConfig -Root $root -Name "sources" -DefaultValue ([PSCustomObject]@{}) -AppliedOverrides $appliedOverrides
    $matchingRulesConfig = Read-JobCrawlerLayeredConfig -Root $root -Name "matching_rules" -DefaultValue ([PSCustomObject]@{}) -AppliedOverrides $appliedOverrides
    $workbookConfig = Read-JobCrawlerLayeredConfig -Root $root -Name "workbook" -DefaultValue ([PSCustomObject]@{}) -AppliedOverrides $appliedOverrides
    $preferencesConfig = Read-JobCrawlerLayeredConfig -Root $root -Name "preferences" -DefaultValue ([PSCustomObject]@{}) -AppliedOverrides $appliedOverrides

    return [PSCustomObject]@{
        Root           = $root
        LocalOverrides = @($appliedOverrides.ToArray())
        Runtime        = $runtimeConfig
        CrawlModes     = $crawlModesConfig
        Sources        = $sourcesConfig
        MatchingRules  = $matchingRulesConfig
        Workbook       = $workbookConfig
        Preferences    = $preferencesConfig
    }
}

function ConvertTo-ConfigBoolean {
    param(
        [AllowNull()]$Value,
        [bool]$DefaultValue = $false
    )

    if ($null -eq $Value) {
        return $DefaultValue
    }
    if ($Value -is [bool]) {
        return [bool]$Value
    }

    $text = ([string]$Value).Trim().ToLowerInvariant()
    if ($text -in @("1", "true", "yes", "y", "on")) {
        return $true
    }
    if ($text -in @("0", "false", "no", "n", "off")) {
        return $false
    }

    return $DefaultValue
}

function Test-JobCrawlerSourceEnabledByDefault {
    param(
        [AllowNull()]$SourcesConfig,
        [string]$SourceKey,
        [bool]$DefaultValue = $true
    )

    $value = Get-ConfigPathValue -Object $SourcesConfig -Path ("sources.{0}.enabled_by_default" -f $SourceKey) -DefaultValue $null
    return ConvertTo-ConfigBoolean -Value $value -DefaultValue $DefaultValue
}

function Test-JobCrawlerSourceRequiresCredentials {
    param(
        [AllowNull()]$SourcesConfig,
        [string]$SourceKey
    )

    $value = Get-ConfigPathValue -Object $SourcesConfig -Path ("sources.{0}.requires_credentials" -f $SourceKey) -DefaultValue $false
    return ConvertTo-ConfigBoolean -Value $value -DefaultValue $false
}

function Get-JobCrawlerSourceLabel {
    param(
        [AllowNull()]$SourcesConfig,
        [string]$SourceKey,
        [string]$DefaultValue
    )

    return [string](Get-ConfigPathValue -Object $SourcesConfig -Path ("sources.{0}.label" -f $SourceKey) -DefaultValue $DefaultValue)
}

function Get-JobCrawlerSourceDefinitions {
    param([AllowNull()]$SourcesConfig)

    $defaultSourceOrder = @("apec", "hellowork", "wttj_public", "linkedin", "france_travail", "adzuna", "welcome_kit")
    $sourceOrder = @(Get-ConfigStringArray (Get-ConfigPathValue -Object $SourcesConfig -Path "source_order" -DefaultValue $defaultSourceOrder))
    if ($sourceOrder.Count -eq 0) {
        $sourceOrder = $defaultSourceOrder
    }

    $defaults = @{
        apec = @{
            Label = "APEC"; Enabled = $true; Credentials = $false; Function = "Get-ApecJobs"; Skip = "SkipApec"; Enable = ""; FallbackFor = ""
        }
        hellowork = @{
            Label = "HelloWork"; Enabled = $true; Credentials = $false; Function = "Get-HelloWorkJobs"; Skip = "SkipHelloWork"; Enable = ""; FallbackFor = ""
        }
        wttj_public = @{
            Label = "Welcome to the Jungle public"; Enabled = $true; Credentials = $false; Function = "Get-WttjPublicFallbackJobs"; Skip = "DisableWttjPublicFallback"; Enable = ""; FallbackFor = "welcome_kit"
        }
        linkedin = @{
            Label = "LinkedIn public guest"; Enabled = $true; Credentials = $false; Function = "Get-LinkedInJobs"; Skip = "SkipLinkedIn"; Enable = ""; FallbackFor = ""
        }
        france_travail = @{
            Label = "France Travail API"; Enabled = $false; Credentials = $true; Function = "Get-FranceTravailJobs"; Skip = "SkipFranceTravail"; Enable = "EnableFranceTravail"; FallbackFor = ""
        }
        adzuna = @{
            Label = "Adzuna API"; Enabled = $false; Credentials = $true; Function = "Get-AdzunaJobs"; Skip = "SkipAdzuna"; Enable = "EnableAdzuna"; FallbackFor = ""
        }
        welcome_kit = @{
            Label = "WelcomeKit API"; Enabled = $false; Credentials = $true; Function = "Get-WelcomeKitJobs"; Skip = "DisableWelcomeKit"; Enable = "EnableWelcomeKit"; FallbackFor = ""
        }
    }

    $definitions = New-Object System.Collections.Generic.List[object]
    foreach ($sourceKey in $sourceOrder) {
        if ([string]::IsNullOrWhiteSpace($sourceKey)) {
            continue
        }

        $key = [string]$sourceKey
        $fallback = $defaults[$key]
        if ($null -eq $fallback) {
            $fallback = @{ Label = $key; Enabled = $true; Credentials = $false; Function = ""; Skip = ""; Enable = ""; FallbackFor = "" }
        }

        $definition = [PSCustomObject]@{
            Key                = $key
            Label              = [string](Get-ConfigPathValue -Object $SourcesConfig -Path ("sources.{0}.label" -f $key) -DefaultValue $fallback.Label)
            ShortLabel         = [string](Get-ConfigPathValue -Object $SourcesConfig -Path ("sources.{0}.short_label" -f $key) -DefaultValue $fallback.Label)
            EnabledByDefault   = Test-JobCrawlerSourceEnabledByDefault -SourcesConfig $SourcesConfig -SourceKey $key -DefaultValue ([bool]$fallback.Enabled)
            RequiresCredential = ConvertTo-ConfigBoolean -Value (Get-ConfigPathValue -Object $SourcesConfig -Path ("sources.{0}.requires_credentials" -f $key) -DefaultValue $fallback.Credentials) -DefaultValue ([bool]$fallback.Credentials)
            CrawlFunction      = [string](Get-ConfigPathValue -Object $SourcesConfig -Path ("sources.{0}.crawl_function" -f $key) -DefaultValue $fallback.Function)
            SkipSwitch         = [string](Get-ConfigPathValue -Object $SourcesConfig -Path ("sources.{0}.skip_switch" -f $key) -DefaultValue $fallback.Skip)
            EnableSwitch       = [string](Get-ConfigPathValue -Object $SourcesConfig -Path ("sources.{0}.enable_switch" -f $key) -DefaultValue $fallback.Enable)
            FallbackFor        = [string](Get-ConfigPathValue -Object $SourcesConfig -Path ("sources.{0}.fallback_for" -f $key) -DefaultValue $fallback.FallbackFor)
        }
        $definitions.Add($definition) | Out-Null
    }

    return @($definitions.ToArray())
}

function Get-JobCrawlerCredentialValue {
    param(
        [AllowNull()]$SourcesConfig,
        [string]$SourceKey,
        [string]$CredentialKey,
        [AllowNull()][string]$FallbackValue = ""
    )

    $envName = Get-ConfigPathValue -Object $SourcesConfig -Path ("credentials.{0}.{1}.env" -f $SourceKey, $CredentialKey) -DefaultValue ""
    if ([string]::IsNullOrWhiteSpace($envName)) {
        return $FallbackValue
    }

    $value = [Environment]::GetEnvironmentVariable($envName, "Process")
    if ([string]::IsNullOrWhiteSpace($value)) {
        $value = [Environment]::GetEnvironmentVariable($envName, "User")
    }
    if ([string]::IsNullOrWhiteSpace($value)) {
        $value = [Environment]::GetEnvironmentVariable($envName, "Machine")
    }
    if ([string]::IsNullOrWhiteSpace($value)) {
        return $FallbackValue
    }

    return $value
}

function Get-JobCrawlerCredentialStatuses {
    param([AllowNull()]$SourcesConfig)

    $rows = New-Object System.Collections.Generic.List[object]
    $credentials = Get-ConfigProperty -Object $SourcesConfig -Name "credentials" -DefaultValue $null
    if ($null -eq $credentials) {
        return @()
    }

    foreach ($sourceName in @($credentials.PSObject.Properties.Name)) {
        $sourceCredentials = $credentials.$sourceName
        foreach ($credentialName in @($sourceCredentials.PSObject.Properties.Name)) {
            $envName = Get-ConfigProperty -Object $sourceCredentials.$credentialName -Name "env" -DefaultValue ""
            $defaultValue = Get-ConfigProperty -Object $sourceCredentials.$credentialName -Name "default" -DefaultValue ""
            $status = "missing"
            if (-not [string]::IsNullOrWhiteSpace($envName)) {
                $processValue = [Environment]::GetEnvironmentVariable($envName, "Process")
                $userValue = [Environment]::GetEnvironmentVariable($envName, "User")
                $machineValue = [Environment]::GetEnvironmentVariable($envName, "Machine")
                if (-not [string]::IsNullOrWhiteSpace($processValue) -or -not [string]::IsNullOrWhiteSpace($userValue) -or -not [string]::IsNullOrWhiteSpace($machineValue)) {
                    $status = "set"
                }
                elseif (-not [string]::IsNullOrWhiteSpace($defaultValue)) {
                    $status = "default"
                }
            }
            $rows.Add([PSCustomObject]@{
                Source = $sourceName
                Credential = $credentialName
                EnvironmentVariable = $envName
                Status = $status
            }) | Out-Null
        }
    }

    return @($rows.ToArray())
}

function Test-JobCrawlerConfig {
    param([AllowNull()]$Config)

    $issues = New-Object System.Collections.Generic.List[string]
    foreach ($mode in @("Fast", "Default", "Deep")) {
        $modeConfig = Get-ConfigPathValue -Object $Config.CrawlModes -Path ("modes.{0}" -f $mode) -DefaultValue $null
        if ($null -eq $modeConfig) {
            $issues.Add("Missing crawl mode: $mode") | Out-Null
        }
    }

    $minimumScore = Get-ConfigPathValue -Object $Config.MatchingRules -Path "thresholds.minimum_match_score" -DefaultValue $null
    if ($null -eq $minimumScore) {
        $issues.Add("Missing matching_rules.thresholds.minimum_match_score") | Out-Null
    }

    $linkedInQueries = Get-ConfigStringArray (Get-ConfigPathValue -Object $Config.Sources -Path "queries.linkedin" -DefaultValue @())
    if ($linkedInQueries.Count -eq 0) {
        $issues.Add("Missing sources.queries.linkedin") | Out-Null
    }

    $apiQueries = Get-ConfigStringArray (Get-ConfigPathValue -Object $Config.Sources -Path "queries.api" -DefaultValue @())
    if ($apiQueries.Count -eq 0) {
        $issues.Add("Missing sources.queries.api") | Out-Null
    }

    $statusOptions = Get-ConfigStringArray (Get-ConfigPathValue -Object $Config.Workbook -Path "status_options" -DefaultValue @())
    if ($statusOptions.Count -gt 0 -and "new" -notin $statusOptions) {
        $issues.Add("workbook.status_options must include 'new'") | Out-Null
    }

    foreach ($sourceKey in @("apec", "hellowork", "wttj_public", "linkedin", "france_travail", "adzuna", "welcome_kit")) {
        $sourceConfig = Get-ConfigPathValue -Object $Config.Sources -Path ("sources.{0}" -f $sourceKey) -DefaultValue $null
        if ($null -eq $sourceConfig) {
            $issues.Add("Missing sources.$sourceKey metadata") | Out-Null
        }
    }

    foreach ($source in @(Get-JobCrawlerSourceDefinitions -SourcesConfig $Config.Sources)) {
        if ([string]::IsNullOrWhiteSpace([string]$source.CrawlFunction)) {
            $issues.Add("Missing crawl function for source '$($source.Key)'") | Out-Null
        }
        if ([string]::IsNullOrWhiteSpace([string]$source.Label)) {
            $issues.Add("Missing label for source '$($source.Key)'") | Out-Null
        }
    }

    return [PSCustomObject]@{
        IsValid = ($issues.Count -eq 0)
        Issues = @($issues.ToArray())
    }
}
