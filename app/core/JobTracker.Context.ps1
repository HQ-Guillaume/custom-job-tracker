function New-JobCrawlerContext {
    param(
        [string]$ProjectRoot,
        [string]$ConfigDirectory,
        [AllowNull()]$Config,
        [AllowNull()][hashtable]$Runtime = @{}
    )

    $runtimeValues = [ordered]@{}
    foreach ($key in @($Runtime.Keys)) {
        $runtimeValues[[string]$key] = $Runtime[$key]
    }

    return [PSCustomObject]@{
        ProjectRoot     = $ProjectRoot
        ConfigDirectory = $ConfigDirectory
        Config          = $Config
        Runtime         = [PSCustomObject]$runtimeValues
        CreatedAt       = [DateTimeOffset]::Now
    }
}

function Set-JobCrawlerScriptContext {
    param(
        [AllowNull()]$Context,
        [switch]$PassThru
    )

    if ($null -eq $Context) {
        return $null
    }

    $script:JobCrawlerContext = $Context
    $script:ProjectRoot = [string]$Context.ProjectRoot

    if ($null -ne $Context.Config) {
        $script:JobCrawlerConfig = $Context.Config
        $script:JobCrawlerRuntimeConfig = $Context.Config.Runtime
        $script:JobCrawlerSourcesConfig = $Context.Config.Sources
        $script:JobCrawlerMatchingRules = $Context.Config.MatchingRules
        $script:JobCrawlerWorkbookConfig = $Context.Config.Workbook
    }

    if ($null -ne $Context.Runtime) {
        foreach ($property in @($Context.Runtime.PSObject.Properties)) {
            Set-Variable -Name $property.Name -Value $property.Value -Scope Script
        }
    }

    if ($PassThru) {
        return $Context
    }

    return $null
}

function Get-JobCrawlerContextValue {
    param(
        [string]$Name,
        [AllowNull()]$DefaultValue = $null
    )

    if (-not (Get-Variable -Name JobCrawlerContext -Scope Script -ErrorAction SilentlyContinue)) {
        return $DefaultValue
    }
    if ($null -eq $script:JobCrawlerContext -or $null -eq $script:JobCrawlerContext.Runtime) {
        return $DefaultValue
    }

    $property = @($script:JobCrawlerContext.Runtime.PSObject.Properties | Where-Object { $_.Name -eq $Name } | Select-Object -First 1)
    if ($property.Count -eq 0) {
        return $DefaultValue
    }

    return $property[0].Value
}

