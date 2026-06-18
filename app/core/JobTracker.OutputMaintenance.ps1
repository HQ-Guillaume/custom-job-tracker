function Test-JobCrawlerPathInsideRoot {
    param(
        [string]$Path,
        [string]$RootPath
    )

    if ([string]::IsNullOrWhiteSpace($Path) -or [string]::IsNullOrWhiteSpace($RootPath)) {
        return $false
    }

    $resolvedRoot = (Resolve-Path -LiteralPath $RootPath -ErrorAction Stop).Path
    $resolvedPath = $null
    if (Test-Path -LiteralPath $Path) {
        $resolvedPath = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path
    }
    else {
        $parent = Split-Path -Parent $Path
        if ([string]::IsNullOrWhiteSpace($parent) -or -not (Test-Path -LiteralPath $parent)) {
            return $false
        }
        $resolvedPath = Join-Path ((Resolve-Path -LiteralPath $parent -ErrorAction Stop).Path) (Split-Path -Leaf $Path)
    }

    return ([string]$resolvedPath).StartsWith([string]$resolvedRoot, [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-JobCrawlerDirectoryStats {
    param(
        [string]$Path,
        [string]$Label = ""
    )

    $files = @()
    if (-not [string]::IsNullOrWhiteSpace($Path) -and (Test-Path -LiteralPath $Path)) {
        $files = @(Get-ChildItem -LiteralPath $Path -Recurse -File -ErrorAction SilentlyContinue)
    }

    $measurement = $files | Measure-Object Length -Sum
    $bytes = [int64]0
    if ($null -ne $measurement -and $null -ne $measurement.Sum) {
        $bytes = [int64]$measurement.Sum
    }
    return [PSCustomObject]@{
        Label = $Label
        Path = $Path
        FileCount = $files.Count
        Bytes = $bytes
        Megabytes = [Math]::Round(([double]$bytes / 1MB), 2)
    }
}

function Clear-JobCrawlerManagedFiles {
    param(
        [string]$Path,
        [string]$ProjectRoot,
        [string]$Pattern = "*",
        [int]$OlderThanDays = 0,
        [switch]$Recurse,
        [switch]$WhatIf
    )

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        return [PSCustomObject]@{ Path = $Path; RemovedFiles = 0; RemovedBytes = 0 }
    }
    if (-not (Test-JobCrawlerPathInsideRoot -Path $Path -RootPath $ProjectRoot)) {
        throw "Refusing to clean path outside project root: $Path"
    }

    $cutoff = (Get-Date).AddDays(-[Math]::Abs($OlderThanDays))
    $files = @(Get-ChildItem -LiteralPath $Path -Filter $Pattern -File -Recurse:$Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -le $cutoff })

    $removedFiles = 0
    $removedBytes = [int64]0
    foreach ($file in $files) {
        $removedFiles++
        $removedBytes += [int64]$file.Length
        if (-not $WhatIf) {
            Remove-Item -LiteralPath $file.FullName -Force -ErrorAction SilentlyContinue
        }
    }

    return [PSCustomObject]@{
        Path = $Path
        RemovedFiles = $removedFiles
        RemovedBytes = $removedBytes
    }
}

function Invoke-JobCrawlerOutputCleanup {
    param(
        [string]$ProjectRoot = $script:ProjectRoot,
        [string]$CacheDirectory = $script:CacheDirectory,
        [switch]$Cache,
        [switch]$Logs,
        [switch]$Diagnostics,
        [switch]$Backups,
        [switch]$All,
        [int]$OlderThanDays = -1,
        [switch]$WhatIf
    )

    if ($OlderThanDays -lt 0) {
        $OlderThanDays = [int](Get-ConfigPathValue -Object $script:JobCrawlerRuntimeConfig -Path "output.cleanup_default_age_days" -DefaultValue 14)
    }

    $trackerRelativePath = [string](Get-ConfigPathValue -Object $script:JobCrawlerRuntimeConfig -Path "defaults.tracker_path" -DefaultValue "output\jobs_tracker.xlsx")
    $outputDirectory = Resolve-JobCrawlerPath -BasePath $ProjectRoot -Path (Split-Path -Parent $trackerRelativePath)
    $paths = New-Object System.Collections.Generic.List[object]
    if ($All -or $Cache) {
        $paths.Add([PSCustomObject]@{ Label = "cache"; Path = $CacheDirectory; Pattern = "*"; Recurse = $true }) | Out-Null
    }
    if ($All -or $Logs) {
        $paths.Add([PSCustomObject]@{ Label = "launcher logs"; Path = (Join-Path $outputDirectory "launcher_logs"); Pattern = "launcher_run_*.log"; Recurse = $false }) | Out-Null
        $paths.Add([PSCustomObject]@{ Label = "launcher run scripts"; Path = $outputDirectory; Pattern = "launcher_run_*.ps1"; Recurse = $false }) | Out-Null
    }
    if ($All -or $Diagnostics) {
        $diagnosticsPath = Resolve-JobCrawlerPath -BasePath $ProjectRoot -Path ([string](Get-ConfigPathValue -Object $script:JobCrawlerRuntimeConfig -Path "output.diagnostics_directory" -DefaultValue "output\diagnostics"))
        $paths.Add([PSCustomObject]@{ Label = "diagnostics"; Path = $diagnosticsPath; Pattern = "*"; Recurse = $true }) | Out-Null
    }
    if ($All -or $Backups) {
        $paths.Add([PSCustomObject]@{ Label = "backups"; Path = (Join-Path $outputDirectory "backups"); Pattern = "*"; Recurse = $true }) | Out-Null
    }

    $results = New-Object System.Collections.Generic.List[object]
    foreach ($item in @($paths.ToArray())) {
        $result = Clear-JobCrawlerManagedFiles -Path $item.Path -ProjectRoot $ProjectRoot -Pattern $item.Pattern -OlderThanDays $OlderThanDays -Recurse:([bool]$item.Recurse) -WhatIf:$WhatIf
        $results.Add([PSCustomObject]@{
            Label = $item.Label
            Path = $result.Path
            RemovedFiles = $result.RemovedFiles
            RemovedMB = [Math]::Round(([double]$result.RemovedBytes / 1MB), 2)
            WhatIf = [bool]$WhatIf
        }) | Out-Null
    }

    return @($results.ToArray())
}
