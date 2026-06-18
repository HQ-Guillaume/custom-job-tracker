[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$tempBase = [System.IO.Path]::GetTempPath()
$tempRoot = Join-Path $tempBase ("custom-job-tracker-portable-{0}" -f ([guid]::NewGuid().ToString("N")))

function Get-TestPowerShellExecutable {
    $pwsh = Get-Command pwsh -ErrorAction SilentlyContinue
    if ($null -ne $pwsh) {
        return $pwsh.Source
    }

    $powershell = Get-Command powershell.exe -ErrorAction SilentlyContinue
    if ($null -ne $powershell) {
        return $powershell.Source
    }

    return "powershell"
}

try {
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $projectRoot "app") -Destination (Join-Path $tempRoot "app") -Recurse -Force
    Copy-Item -LiteralPath (Join-Path $projectRoot "config") -Destination (Join-Path $tempRoot "config") -Recurse -Force

    $powershell = Get-TestPowerShellExecutable
    $crawlerPath = Join-Path $tempRoot "app\cli\Find-AnalyticsJobs.ps1"
    $arguments = @("-NoLogo", "-NoProfile")
    if ([Environment]::OSVersion.Platform -eq [PlatformID]::Win32NT) {
        $arguments += @("-ExecutionPolicy", "Bypass")
    }
    $arguments += @("-File", $crawlerPath, "-ValidateConfig")
    & $powershell @arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Portable project-root validation failed with exit code $LASTEXITCODE."
    }
}
finally {
    if (Test-Path -LiteralPath $tempRoot) {
        $resolvedTempRoot = (Resolve-Path -LiteralPath $tempRoot).Path
        $resolvedTempBase = (Resolve-Path -LiteralPath $tempBase).Path
        if ($resolvedTempRoot.StartsWith($resolvedTempBase, [System.StringComparison]::OrdinalIgnoreCase)) {
            Remove-Item -LiteralPath $resolvedTempRoot -Recurse -Force
        }
    }
}

Write-Host "Portable project-root test passed."
