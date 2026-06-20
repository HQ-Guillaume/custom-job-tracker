[CmdletBinding()]
param(
    [switch]$SelfTest,
    [switch]$BuildSelfTest,
    [switch]$SmokeTest,
    [switch]$RunSmokeTest
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$script:ProjectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$script:ConfigDirectory = Join-Path $script:ProjectRoot "config"
$script:CliDirectory = Join-Path $script:ProjectRoot "app\cli"
$script:CoreDirectory = Join-Path $script:ProjectRoot "app\core"

. (Join-Path $script:CoreDirectory "JobTracker.Common.ps1")
. (Join-Path $script:CoreDirectory "JobTracker.Config.ps1")
. (Join-Path $script:CoreDirectory "JobTracker.Runtime.ps1")
. (Join-Path $script:CoreDirectory "JobTracker.OutputMaintenance.ps1")

function Set-LauncherCrawlerConfig {
    param([string]$ProfileId = "")

    $script:CrawlerConfig = Get-JobCrawlerConfig -ConfigDirectory $script:ConfigDirectory -ProfileId $ProfileId
    $script:ConfigValidation = Test-JobCrawlerConfig -Config $script:CrawlerConfig
    if (-not $script:ConfigValidation.IsValid) {
        throw ("Invalid crawler config:`n- {0}" -f (($script:ConfigValidation.Issues) -join "`n- "))
    }
}

Set-LauncherCrawlerConfig

function Test-LauncherProfileConfigured {
    return Test-JobCrawlerProfileConfigured -Config $script:CrawlerConfig
}

function Get-LauncherDefaultTrackerPath {
    return Get-JobCrawlerTrackerPath -ProjectRoot $script:ProjectRoot -Config $script:CrawlerConfig
}

function Get-LauncherDisplayPath {
    param([AllowNull()][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return ""
    }

    $root = [string]$script:ProjectRoot
    if (-not [string]::IsNullOrWhiteSpace($root) -and $Path.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) {
        $relative = [regex]::Replace($Path.Substring($root.Length), "^[\\/]+", "")
        if (-not [string]::IsNullOrWhiteSpace($relative)) {
            return $relative
        }
    }

    return $Path
}

function Get-LauncherOutputDirectory {
    $trackerPath = Get-LauncherDefaultTrackerPath
    return Split-Path -Parent $trackerPath
}

function Set-LauncherOutputPaths {
    $outputDirectory = Get-LauncherOutputDirectory
    $script:LauncherErrorLogPath = Join-Path $outputDirectory "launcher_error.log"
    $script:LauncherLastRunLogPath = Join-Path $outputDirectory "launcher_last_run.log"
    $script:LauncherRunLogDirectory = Join-Path $outputDirectory "launcher_logs"
    $script:LauncherRunLockPath = Join-Path $outputDirectory "crawler_run.lock.json"

    $isRunning = $false
    if ($null -ne $script:RunningProcess) {
        try {
            $isRunning = -not $script:RunningProcess.HasExited
        }
        catch {
            $isRunning = $false
        }
    }
    if (-not $isRunning) {
        $script:LauncherRunLogPath = $script:LauncherLastRunLogPath
    }
}

function ConvertTo-CommandLineArgument {
    param([string]$Value)

    if ($null -eq $Value) {
        return '""'
    }

    $text = [string]$Value
    if ($text -match '^[A-Za-z0-9_./:\\=-]+$') {
        return $text
    }

    return '"' + ($text.Replace('"', '\"')) + '"'
}

function ConvertTo-PowerShellLiteral {
    param([AllowNull()][string]$Value)

    if ($null -eq $Value) {
        return "''"
    }

    return "'" + ([string]$Value).Replace("'", "''") + "'"
}

function ConvertTo-PowerShellCommandToken {
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return "''"
    }

    $text = [string]$Value
    if ($text -match '^-[A-Za-z][A-Za-z0-9]*$') {
        return $text
    }

    return ConvertTo-PowerShellLiteral $text
}

function Get-LauncherPowerShellPath {
    $windowsPowerShell = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
    if (Test-Path -LiteralPath $windowsPowerShell) {
        return $windowsPowerShell
    }

    return Join-Path $PSHOME "powershell.exe"
}

if ($SelfTest) {
    Write-Host "WinForms launcher self-test passed."
    Write-Host ("Project root: {0}" -f $script:ProjectRoot)
    Write-Host ("Tracker: {0}" -f (Get-LauncherDefaultTrackerPath))
    Write-Host "Credential status:"
    Get-JobCrawlerCredentialStatuses -SourcesConfig $script:CrawlerConfig.Sources | Format-Table -AutoSize
    return
}

if ([Threading.Thread]::CurrentThread.ApartmentState -ne "STA") {
    throw "Start the GUI with powershell.exe -STA -File .\app\cli\Launch-AnalyticsJobCrawlerGui.ps1, or double-click Run-CustomJobTracker-GUI.cmd."
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:LauncherErrorLogPath = ""
$script:LauncherLastRunLogPath = ""
$script:LauncherRunLogDirectory = ""
$script:LauncherRunLogPath = ""
$script:LauncherRunLockPath = ""
$script:LauncherRunId = ""
$script:MainForm = $null
$script:LogTextBox = $null
$script:RunButton = $null
$script:StopButton = $null
$script:ProgressBar = $null
$script:StatusLabel = $null
$script:CredentialList = $null
$script:ReadinessList = $null
$script:RunningProcess = $null
$script:RunPollTimer = $null
$script:RunLogLineCount = 0
$script:RunScriptPath = $null
$script:RunLockHeld = $false
$script:LastCrawlerExitCode = $null
Set-LauncherOutputPaths
$script:SettingsGroup = $null
$script:SourceListView = $null
$script:SourceCheckboxes = @()
$script:IsRefreshingSourceList = $false
$script:ModeComboBox = $null
$script:DaysBackComboBox = $null
$script:ProfileComboBox = $null
$script:ProfileOptions = @()
$script:ProfileActionButtons = @()
$script:IsRefreshingProfileCombo = $false
$script:DryRunCheckBox = $null
$script:DiagnosticModeCheckBox = $null
$script:DisableCacheCheckBox = $null
$script:ExcelAutomationAvailable = $null
$script:ExcelAutomationCheckedAt = [DateTime]::MinValue
$script:WorkbookWriterPlan = $null
$script:CleanupButton = $null

function Write-LauncherError {
    param(
        [string]$Context,
        [System.Exception]$Exception
    )

    try {
        $outputDirectory = Get-LauncherOutputDirectory
        if (-not (Test-Path -LiteralPath $outputDirectory)) {
            New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
        }

        $message = @(
            ("[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Context),
            $Exception.Message,
            $Exception.ToString(),
            ""
        )
        Add-Content -LiteralPath $script:LauncherErrorLogPath -Value $message -Encoding UTF8
    }
    catch {
        Write-Host ("Could not write launcher error log: {0}" -f $_.Exception.Message)
    }
}

function Show-LauncherException {
    param(
        [string]$Context,
        [System.Exception]$Exception
    )

    Write-LauncherError -Context $Context -Exception $Exception

    $message = "{0}: {1}`n`nDetails were written to:`n{2}" -f $Context, $Exception.Message, $script:LauncherErrorLogPath
    if ($null -ne $script:LogTextBox) {
        Add-LogLine ("ERROR: {0}" -f $message.Replace("`n", " "))
    }

    [System.Windows.Forms.MessageBox]::Show($message, "Custom Job Tracker Launcher") | Out-Null
}

function Invoke-LauncherAction {
    param(
        [string]$Context,
        [scriptblock]$Action
    )

    try {
        & $Action
    }
    catch {
        Show-LauncherException -Context $Context -Exception $_.Exception
    }
}

[System.Windows.Forms.Application]::SetUnhandledExceptionMode([System.Windows.Forms.UnhandledExceptionMode]::CatchException)
[System.Windows.Forms.Application]::add_ThreadException({
    param($sender, $eventArgs)
    Show-LauncherException -Context "Unexpected launcher error" -Exception $eventArgs.Exception
})
[AppDomain]::CurrentDomain.add_UnhandledException({
    param($sender, $eventArgs)

    $exception = $eventArgs.ExceptionObject
    if ($exception -isnot [System.Exception]) {
        $exception = New-Object System.Exception([string]$eventArgs.ExceptionObject)
    }

    Write-LauncherError -Context "Fatal launcher error" -Exception $exception
})

function Add-LogLine {
    param([string]$Line)

    if ($null -eq $script:LogTextBox -or $null -eq $script:MainForm) {
        Write-Host $Line
        return
    }

    if ($script:MainForm.InvokeRequired) {
        $message = $Line
        [void]$script:MainForm.BeginInvoke([System.Action]{ Add-LogLine -Line $message })
        return
    }

    $script:LogTextBox.AppendText($Line + [Environment]::NewLine)
    $script:LogTextBox.SelectionStart = $script:LogTextBox.TextLength
    $script:LogTextBox.ScrollToCaret()
}

function Read-LauncherTextFileShared {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        return ""
    }

    for ($attempt = 1; $attempt -le 6; $attempt++) {
        $stream = $null
        $reader = $null
        try {
            $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
            return $reader.ReadToEnd()
        }
        catch {
            if ($attempt -ge 6) {
                throw
            }
            Start-Sleep -Milliseconds (60 * $attempt)
        }
        finally {
            if ($null -ne $reader) {
                $reader.Dispose()
            }
            elseif ($null -ne $stream) {
                $stream.Dispose()
            }
        }
    }

    return ""
}

function Set-LauncherRunningState {
    param([bool]$IsRunning)

    if ($null -eq $script:MainForm) {
        return
    }

    if ($script:MainForm.InvokeRequired) {
        $running = $IsRunning
        [void]$script:MainForm.BeginInvoke([System.Action]{ Set-LauncherRunningState -IsRunning $running })
        return
    }

    $script:RunButton.Enabled = -not $IsRunning
    $script:StopButton.Enabled = $IsRunning
    $script:ModeComboBox.Enabled = -not $IsRunning
    if ($null -ne $script:DaysBackComboBox) {
        $script:DaysBackComboBox.Enabled = -not $IsRunning
    }
    if ($null -ne $script:ProfileComboBox) {
        $script:ProfileComboBox.Enabled = -not $IsRunning
    }
    foreach ($button in @($script:ProfileActionButtons)) {
        $button.Enabled = -not $IsRunning
    }
    if ($null -ne $script:SourceListView) {
        $script:SourceListView.Enabled = -not $IsRunning
    }
    $script:DryRunCheckBox.Enabled = -not $IsRunning
    $script:DiagnosticModeCheckBox.Enabled = -not $IsRunning
    $script:DisableCacheCheckBox.Enabled = -not $IsRunning
    if ($null -ne $script:CleanupButton) {
        $script:CleanupButton.Enabled = -not $IsRunning
    }

    if ($IsRunning) {
        $script:ProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
        $script:ProgressBar.MarqueeAnimationSpeed = 35
        $script:StatusLabel.Text = "Running crawler..."
    }
    else {
        $script:ProgressBar.MarqueeAnimationSpeed = 0
        $script:ProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks
    }
}

function Set-LauncherStatus {
    param([string]$Text)

    if ($null -eq $script:MainForm) {
        return
    }

    if ($script:MainForm.InvokeRequired) {
        $message = $Text
        [void]$script:MainForm.BeginInvoke([System.Action]{ Set-LauncherStatus -Text $message })
        return
    }

    $script:StatusLabel.Text = $Text
}

function Stop-LauncherRunPollTimer {
    if ($null -ne $script:RunPollTimer) {
        try {
            $script:RunPollTimer.Stop()
            $script:RunPollTimer.Dispose()
        }
        catch {
            Write-LauncherError -Context "Could not stop launcher run timer" -Exception $_.Exception
        }
        finally {
            $script:RunPollTimer = $null
        }
    }
}

function Remove-LauncherRunScript {
    if (-not [string]::IsNullOrWhiteSpace($script:RunScriptPath) -and (Test-Path -LiteralPath $script:RunScriptPath)) {
        try {
            Remove-Item -LiteralPath $script:RunScriptPath -Force
        }
        catch {
            Write-LauncherError -Context "Could not remove temporary launcher run script" -Exception $_.Exception
        }
    }

    $script:RunScriptPath = $null
}

function New-LauncherRunLogPath {
    $runId = Get-Date -Format "yyyyMMdd_HHmmss_fff"
    if (-not (Test-Path -LiteralPath $script:LauncherRunLogDirectory)) {
        New-Item -ItemType Directory -Path $script:LauncherRunLogDirectory -Force | Out-Null
    }

    $script:LauncherRunId = $runId
    $script:LauncherRunLogPath = Join-Path $script:LauncherRunLogDirectory ("launcher_run_{0}.log" -f $runId)
    return $script:LauncherRunLogPath
}

function Copy-LauncherRunLogToLastRun {
    try {
        if (-not [string]::IsNullOrWhiteSpace($script:LauncherRunLogPath) -and
            (Test-Path -LiteralPath $script:LauncherRunLogPath) -and
            ([string]$script:LauncherRunLogPath -ne [string]$script:LauncherLastRunLogPath)) {
            Copy-Item -LiteralPath $script:LauncherRunLogPath -Destination $script:LauncherLastRunLogPath -Force
        }
    }
    catch {
        Write-LauncherError -Context "Could not copy launcher run log to last-run log" -Exception $_.Exception
    }
}

function Clear-StaleLauncherRunScripts {
    try {
        $outputDirectory = Get-LauncherOutputDirectory
        if (-not (Test-Path -LiteralPath $outputDirectory)) {
            return
        }

        $cutoff = (Get-Date).AddHours(-12)
        Get-ChildItem -LiteralPath $outputDirectory -Filter "launcher_run_*.ps1" -File -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -lt $cutoff -and [string]$_.FullName -ne [string]$script:RunScriptPath } |
            Remove-Item -Force

        if (Test-Path -LiteralPath $script:LauncherRunLogDirectory) {
            $logCutoff = (Get-Date).AddDays(-14)
            Get-ChildItem -LiteralPath $script:LauncherRunLogDirectory -Filter "launcher_run_*.log" -File -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTime -lt $logCutoff -and [string]$_.FullName -ne [string]$script:LauncherRunLogPath } |
                Remove-Item -Force
        }
    }
    catch {
        Write-LauncherError -Context "Could not remove stale launcher run scripts" -Exception $_.Exception
    }
}

function Test-LauncherProcessIdRunning {
    param([int]$ProcessId)

    if ($ProcessId -le 0) {
        return $false
    }

    try {
        $process = [System.Diagnostics.Process]::GetProcessById($ProcessId)
        return -not $process.HasExited
    }
    catch {
        return $false
    }
}

function Get-LauncherRunLock {
    if ([string]::IsNullOrWhiteSpace($script:LauncherRunLockPath) -or -not (Test-Path -LiteralPath $script:LauncherRunLockPath)) {
        return $null
    }

    try {
        return Get-Content -LiteralPath $script:LauncherRunLockPath -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    catch {
        Write-LauncherError -Context "Could not read launcher run lock" -Exception $_.Exception
        return $null
    }
}

function Remove-LauncherRunLock {
    if (-not $script:RunLockHeld -and -not (Test-Path -LiteralPath $script:LauncherRunLockPath)) {
        return
    }

    try {
        if (Test-Path -LiteralPath $script:LauncherRunLockPath) {
            Remove-Item -LiteralPath $script:LauncherRunLockPath -Force
        }
    }
    catch {
        Write-LauncherError -Context "Could not remove launcher run lock" -Exception $_.Exception
    }
    finally {
        $script:RunLockHeld = $false
    }
}

function Get-ActiveLauncherRunLock {
    $lock = Get-LauncherRunLock
    if ($null -eq $lock) {
        return $null
    }

    $processId = [int](Get-ConfigProperty -Object $lock -Name "process_id" -DefaultValue 0)
    $createdAtText = [string](Get-ConfigProperty -Object $lock -Name "created_at" -DefaultValue "")
    $createdAt = [DateTime]::MinValue
    if (-not [string]::IsNullOrWhiteSpace($createdAtText)) {
        [DateTime]::TryParse($createdAtText, [ref]$createdAt) | Out-Null
    }

    if ($processId -gt 0 -and (Test-LauncherProcessIdRunning -ProcessId $processId)) {
        return $lock
    }
    if ($processId -le 0 -and $createdAt -gt [DateTime]::MinValue -and $createdAt -gt (Get-Date).AddMinutes(-30)) {
        return $lock
    }

    Remove-LauncherRunLock
    return $null
}

function New-LauncherRunLock {
    param(
        [string]$RunId,
        [string]$LogPath
    )

    $activeLock = Get-ActiveLauncherRunLock
    if ($null -ne $activeLock) {
        $activeLog = [string](Get-ConfigProperty -Object $activeLock -Name "log_path" -DefaultValue "")
        if ([string]::IsNullOrWhiteSpace($activeLog)) {
            $activeLog = $script:LauncherRunLockPath
        }
        throw "Another crawler run appears to be active. Check or close it before starting a new run. Active log: $activeLog"
    }

    $lock = [ordered]@{
        run_id     = $RunId
        created_at = (Get-Date).ToString("o")
        process_id = 0
        log_path   = $LogPath
    }
    $json = $lock | ConvertTo-Json -Depth 5
    Set-Content -LiteralPath $script:LauncherRunLockPath -Value $json -Encoding UTF8
    $script:RunLockHeld = $true
}

function Set-LauncherRunLockProcessId {
    param([int]$ProcessId)

    if (-not $script:RunLockHeld -or -not (Test-Path -LiteralPath $script:LauncherRunLockPath)) {
        return
    }

    try {
        $lock = Get-LauncherRunLock
        if ($null -eq $lock) {
            return
        }

        $hash = ConvertTo-ConfigHashtable $lock
        $hash["process_id"] = $ProcessId
        $json = $hash | ConvertTo-Json -Depth 5
        Set-Content -LiteralPath $script:LauncherRunLockPath -Value $json -Encoding UTF8
    }
    catch {
        Write-LauncherError -Context "Could not update launcher run lock" -Exception $_.Exception
    }
}

function Read-LauncherRunLog {
    if ([string]::IsNullOrWhiteSpace($script:LauncherRunLogPath) -or -not (Test-Path -LiteralPath $script:LauncherRunLogPath)) {
        return
    }

    try {
        $content = Read-LauncherTextFileShared -Path $script:LauncherRunLogPath
        $lines = @($content -split "\r?\n")
        if ($lines.Count -le $script:RunLogLineCount) {
            return
        }

        for ($index = $script:RunLogLineCount; $index -lt $lines.Count; $index++) {
            if (-not [string]::IsNullOrWhiteSpace($lines[$index])) {
                Add-LogLine $lines[$index]
            }
        }
        $script:RunLogLineCount = $lines.Count
    }
    catch {
        Write-LauncherError -Context "Could not read launcher run log" -Exception $_.Exception
    }
}

function New-LauncherRunScript {
    param(
        [string]$CrawlerPath,
        [string[]]$CrawlerArguments
    )

    $outputDirectory = Get-LauncherOutputDirectory
    if (-not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    }

    Clear-StaleLauncherRunScripts
    $scriptPath = Join-Path $outputDirectory ("launcher_run_{0}.ps1" -f (Get-Date -Format "yyyyMMdd_HHmmss_fff"))
    $argumentTokens = @($CrawlerArguments | ForEach-Object { ConvertTo-PowerShellCommandToken $_ })
    if ($argumentTokens.Count -eq 0) {
        $argumentText = ""
    }
    else {
        $argumentText = " " + ($argumentTokens -join " ")
    }

    $body = @"
`$ErrorActionPreference = 'Stop'
`$ProgressPreference = 'SilentlyContinue'
`$crawlerPath = $(ConvertTo-PowerShellLiteral $CrawlerPath)
`$logPath = $(ConvertTo-PowerShellLiteral $script:LauncherRunLogPath)

function Add-LauncherLogLine {
    param([AllowNull()][string]`$Line)

    `$encoding = New-Object System.Text.UTF8Encoding(`$false)
    `$text = [string]`$Line + [Environment]::NewLine
    `$bytes = `$encoding.GetBytes(`$text)
    for (`$attempt = 1; `$attempt -le 8; `$attempt++) {
        `$stream = `$null
        try {
            `$stream = [System.IO.File]::Open(`$logPath, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
            `$stream.Write(`$bytes, 0, `$bytes.Length)
            return
        }
        catch {
            if (`$attempt -ge 8) {
                [Console]::Error.WriteLine("Could not write launcher log: {0}" -f `$_.Exception.Message)
                throw
            }
            Start-Sleep -Milliseconds (80 * `$attempt)
        }
        finally {
            if (`$null -ne `$stream) {
                `$stream.Dispose()
            }
        }
    }
}

try {
    Add-LauncherLogLine ("[{0}] Launcher started crawler." -f (Get-Date -Format "HH:mm:ss"))
    & `$crawlerPath$argumentText *>&1 | ForEach-Object {
        if (`$null -eq `$_) {
            `$text = ""
        }
        else {
            `$text = [string]`$_
        }
        `$text = `$text.TrimEnd()
        if (-not [string]::IsNullOrWhiteSpace(`$text)) {
            Add-LauncherLogLine `$text
        }
    }
    Add-LauncherLogLine ("[{0}] Crawler process completed." -f (Get-Date -Format "HH:mm:ss"))
    exit 0
}
catch {
    try {
        Add-LauncherLogLine ("[{0}] ERROR: {1}" -f (Get-Date -Format "HH:mm:ss"), `$_.Exception.Message)
        Add-LauncherLogLine (`$_ | Out-String)
    }
    catch {
        [Console]::Error.WriteLine("Crawler launcher failed and could not write the log: {0}" -f `$_.Exception.Message)
    }
    exit 1
}
"@

    Set-Content -LiteralPath $scriptPath -Value $body -Encoding UTF8
    return $scriptPath
}

function Start-LauncherRunPollTimer {
    Stop-LauncherRunPollTimer

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 1000
    $timer.Add_Tick({
        Invoke-LauncherAction -Context "Poll crawler run" -Action {
            Read-LauncherRunLog
            if ($null -ne $script:RunningProcess -and $script:RunningProcess.HasExited) {
                $exitCode = $script:RunningProcess.ExitCode
                Read-LauncherRunLog
                Stop-LauncherRunPollTimer
                Complete-CrawlerRun -ExitCode $exitCode
            }
        }
    })
    $script:RunPollTimer = $timer
    $timer.Start()
}

function Complete-CrawlerRun {
    param([int]$ExitCode)

    if ($null -ne $script:MainForm -and $script:MainForm.InvokeRequired) {
        $completedExitCode = $ExitCode
        try {
            [void]$script:MainForm.BeginInvoke([System.Action]{ Complete-CrawlerRun -ExitCode $completedExitCode })
        }
        catch {
            Write-LauncherError -Context "Could not update launcher after crawl completion" -Exception $_.Exception
        }
        return
    }

    try {
        $script:LastCrawlerExitCode = $ExitCode
        if ($ExitCode -eq 0) {
            Add-LogLine "Crawler finished successfully."
            Set-LauncherStatus "Finished successfully."
        }
        else {
            Add-LogLine ("Crawler finished with error code {0}." -f $ExitCode)
            Set-LauncherStatus ("Finished with error code {0}." -f $ExitCode)
        }

        Set-LauncherRunningState -IsRunning $false
        $script:RunningProcess = $null
        Stop-LauncherRunPollTimer
        Copy-LauncherRunLogToLastRun
        Remove-LauncherRunScript
        Remove-LauncherRunLock
        Refresh-ReadinessChecklist
    }
    catch {
        Show-LauncherException -Context "Could not finalize crawler run" -Exception $_.Exception
    }
}

function Refresh-CredentialList {
    if ($null -eq $script:CredentialList) {
        return
    }

    $script:CredentialList.Items.Clear()
    $rows = @(Get-JobCrawlerCredentialStatuses -SourcesConfig $script:CrawlerConfig.Sources)
    foreach ($row in $rows) {
        $item = New-Object System.Windows.Forms.ListViewItem([string]$row.Source)
        [void]$item.SubItems.Add([string]$row.Credential)
        [void]$item.SubItems.Add([string]$row.EnvironmentVariable)
        [void]$item.SubItems.Add([string]$row.Status)

        if ($row.Status -eq "set") {
            $item.ForeColor = [System.Drawing.Color]::FromArgb(30, 104, 72)
        }
        elseif ($row.Status -eq "default") {
            $item.ForeColor = [System.Drawing.Color]::FromArgb(77, 91, 114)
        }
        else {
            $item.ForeColor = [System.Drawing.Color]::FromArgb(154, 69, 59)
        }

        [void]$script:CredentialList.Items.Add($item)
    }

    if ($null -ne $script:SourceListView -and -not $script:IsRefreshingSourceList) {
        Refresh-SourceCheckboxes
    }
    elseif (Get-Command Refresh-ReadinessChecklist -ErrorAction SilentlyContinue) {
        Refresh-ReadinessChecklist
    }
}

. (Join-Path $script:CliDirectory "JobCrawler.GuiConfig.ps1")

function Add-ReadinessItem {
    param(
        [string]$Area,
        [string]$Status,
        [string]$Detail,
        [string]$Level = "ok"
    )

    if ($null -eq $script:ReadinessList) {
        return
    }

    $item = New-Object System.Windows.Forms.ListViewItem($Area)
    [void]$item.SubItems.Add($Status)
    [void]$item.SubItems.Add($Detail)
    switch ($Level) {
        "error" { $item.ForeColor = [System.Drawing.Color]::FromArgb(154, 69, 59) }
        "warning" { $item.ForeColor = [System.Drawing.Color]::FromArgb(145, 92, 35) }
        default { $item.ForeColor = [System.Drawing.Color]::FromArgb(30, 104, 72) }
    }
    [void]$script:ReadinessList.Items.Add($item)
}

function Test-LauncherExcelAutomationAvailable {
    if ($null -ne $script:ExcelAutomationAvailable -and $script:ExcelAutomationCheckedAt -gt (Get-Date).AddMinutes(-5)) {
        return [bool]$script:ExcelAutomationAvailable
    }

    $excel = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $script:ExcelAutomationAvailable = $true
        $script:ExcelAutomationCheckedAt = Get-Date
        return $true
    }
    catch {
        $script:ExcelAutomationAvailable = $false
        $script:ExcelAutomationCheckedAt = Get-Date
        return $false
    }
    finally {
        if ($null -ne $excel) {
            try {
                $excel.DisplayAlerts = $false
                $excel.Quit()
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
            }
            catch {
            }
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }
}

function Get-LauncherWorkbookOutputBackend {
    $backend = [string](Get-ConfigPathValue -Object $script:CrawlerConfig.Workbook -Path "output_backend" -DefaultValue "auto")
    if ([string]::IsNullOrWhiteSpace($backend)) {
        return "auto"
    }
    $backend = $backend.Trim().ToLowerInvariant()
    if ($backend -notin @("auto", "excel", "openxml")) {
        return "auto"
    }

    return $backend
}

function Get-LauncherWorkbookWriterPlan {
    $backend = Get-LauncherWorkbookOutputBackend
    $excelAvailable = Test-LauncherExcelAutomationAvailable
    $openXmlAvailable = Test-Path -LiteralPath (Join-Path $script:CoreDirectory "JobTracker.OpenXml.ps1")
    $htmlFallbackEnabled = ConvertTo-ConfigBoolean -Value (Get-ConfigPathValue -Object $script:CrawlerConfig.Workbook -Path "html_fallback_enabled" -DefaultValue $true) -DefaultValue $true

    if ($backend -eq "excel") {
        if ($excelAvailable) {
            return [PSCustomObject]@{ Level = "ok"; Status = "Excel XLSX"; Detail = "Configured to use desktop Excel COM for the formatted tracker." }
        }
        return [PSCustomObject]@{ Level = "error"; Status = "Excel missing"; Detail = "Workbook backend is set to excel, but desktop Excel COM is not available." }
    }

    if ($backend -eq "openxml") {
        if ($openXmlAvailable) {
            return [PSCustomObject]@{ Level = "ok"; Status = "No-Excel XLSX"; Detail = "Configured to create the formatted XLSX with the built-in OpenXML writer." }
        }
        return [PSCustomObject]@{ Level = "error"; Status = "Writer missing"; Detail = "Workbook backend is set to openxml, but the OpenXML writer module is missing." }
    }

    if ($excelAvailable) {
        return [PSCustomObject]@{ Level = "ok"; Status = "Excel XLSX"; Detail = "Auto mode will use desktop Excel COM for the richest formatting." }
    }
    if ($openXmlAvailable) {
        return [PSCustomObject]@{ Level = "ok"; Status = "No-Excel XLSX"; Detail = "Auto mode will create the profile tracker workbook with the built-in OpenXML writer." }
    }
    if ($htmlFallbackEnabled) {
        return [PSCustomObject]@{ Level = "warning"; Status = "HTML fallback"; Detail = "No XLSX writer is available; only a readable HTML fallback can be written if export fails." }
    }

    return [PSCustomObject]@{ Level = "error"; Status = "No writer"; Detail = "No workbook writer is available." }
}

function Get-SelectedSourceDefinitionsFromGui {
    $definitions = New-Object System.Collections.Generic.List[object]
    if ($null -ne $script:SourceListView) {
        foreach ($item in $script:SourceListView.Items) {
            if ((Get-GuiCheckedValue $item) -and $null -ne $item.Tag) {
                $definitions.Add($item.Tag) | Out-Null
            }
        }
        return @($definitions.ToArray())
    }

    return @(Get-JobCrawlerSourceDefinitions -SourcesConfig $script:CrawlerConfig.Sources | Where-Object { $_.EnabledByDefault })
}

function Get-EnabledSourceCredentialIssues {
    $issues = New-Object System.Collections.Generic.List[string]
    foreach ($definition in @(Get-SelectedSourceDefinitionsFromGui)) {
        if (-not [bool]$definition.RequiresCredential) {
            continue
        }

        $missing = @(Get-JobCrawlerCredentialStatuses -SourcesConfig $script:CrawlerConfig.Sources |
            Where-Object { $_.Source -eq [string]$definition.Key -and $_.Status -eq "missing" })
        if ($missing.Count -gt 0) {
            $issues.Add(("{0}: {1}" -f $definition.ShortLabel, (($missing | ForEach-Object { $_.Credential }) -join ", "))) | Out-Null
        }
    }

    return @($issues.ToArray())
}

function Test-TrackerWorkbookWritable {
    $trackerPath = Get-LauncherDefaultTrackerPath
    $displayPath = Get-LauncherDisplayPath -Path $trackerPath
    if (-not (Test-Path -LiteralPath $trackerPath)) {
        return [PSCustomObject]@{ Level = "warning"; Status = "Not found"; Detail = ("Will be created: {0}" -f $displayPath) }
    }

    try {
        $stream = [System.IO.File]::Open($trackerPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        $stream.Close()
        return [PSCustomObject]@{ Level = "ok"; Status = "Ready"; Detail = ("Workbook is not locked: {0}" -f $displayPath) }
    }
    catch {
        return [PSCustomObject]@{ Level = "error"; Status = "Locked"; Detail = ("Close this workbook before running: {0}" -f $displayPath) }
    }
}

function Refresh-ReadinessChecklist {
    if ($null -eq $script:ReadinessList) {
        return
    }

    if ($null -ne $script:MainForm -and $script:MainForm.InvokeRequired) {
        try {
            [void]$script:MainForm.BeginInvoke([System.Action]{ Refresh-ReadinessChecklist })
        }
        catch {
            Write-LauncherError -Context "Could not refresh readiness checklist" -Exception $_.Exception
        }
        return
    }

    $script:ReadinessList.BeginUpdate()
    try {
        $script:ReadinessList.Items.Clear()

        $validation = Test-JobCrawlerConfig -Config $script:CrawlerConfig
        if (-not (Test-LauncherProfileConfigured)) {
            Add-ReadinessItem -Area "Profile" -Status "Create profile" -Detail "Click New in the Profile section before crawling." -Level "error"
        }
        elseif ($validation.IsValid) {
            Add-ReadinessItem -Area "Config" -Status "Ready" -Detail ("Profile: {0}" -f $script:CrawlerConfig.Profile.Label)
            $builder = Get-ConfigProperty -Object $script:CrawlerConfig.Profile -Name "profile_builder" -DefaultValue $null
            if ($null -ne $builder) {
                $quality = Get-JobCrawlerProfileQuality `
                    -Label ([string]$script:CrawlerConfig.Profile.Label) `
                    -TargetTitles (Get-ConfigStringArray (Get-ConfigProperty -Object $builder -Name "target_titles" -DefaultValue @())) `
                    -SearchQueries (Get-ConfigStringArray (Get-ConfigProperty -Object $builder -Name "search_queries" -DefaultValue @())) `
                    -ImportantSkills (Get-ConfigStringArray (Get-ConfigProperty -Object $builder -Name "important_skills" -DefaultValue @())) `
                    -ExclusionKeywords (Get-ConfigStringArray (Get-ConfigProperty -Object $builder -Name "exclusion_keywords" -DefaultValue @())) `
                    -TargetLocations (Get-ConfigStringArray (Get-ConfigProperty -Object $builder -Name "target_locations" -DefaultValue @())) `
                    -ExcludedContracts (Get-ConfigStringArray (Get-ConfigProperty -Object $builder -Name "excluded_contracts" -DefaultValue @()))
                $qualityLevel = $(if ($quality.Score -ge 70) { "ok" } elseif ($quality.Score -ge 50) { "warning" } else { "error" })
                $qualityDetail = ("{0}/100 - {1}" -f $quality.Score, $quality.Level)
                $firstFinding = @($quality.Findings | Select-Object -First 1)
                if ($firstFinding.Count -gt 0) {
                    $qualityDetail = "{0}. {1}" -f $qualityDetail, $firstFinding[0]
                }
                Add-ReadinessItem -Area "Profile quality" -Status $quality.Level -Detail $qualityDetail -Level $qualityLevel
            }
        }
        else {
            Add-ReadinessItem -Area "Config" -Status "Review" -Detail (($validation.Issues) -join "; ") -Level "error"
        }

        $workbook = Test-TrackerWorkbookWritable
        Add-ReadinessItem -Area "Tracker" -Status $workbook.Status -Detail $workbook.Detail -Level $workbook.Level

        $writerPlan = Get-LauncherWorkbookWriterPlan
        $script:WorkbookWriterPlan = $writerPlan
        Add-ReadinessItem -Area "Workbook writer" -Status $writerPlan.Status -Detail $writerPlan.Detail -Level $writerPlan.Level

        $outputPath = Get-LauncherOutputDirectory
        if (Test-Path -LiteralPath $outputPath) {
            Add-ReadinessItem -Area "Output" -Status "Ready" -Detail (Get-LauncherOutputStatsText)
        }
        else {
            Add-ReadinessItem -Area "Output" -Status "Will create" -Detail $outputPath -Level "warning"
        }

        $selectedSources = @(Get-SelectedSourceDefinitionsFromGui)
        if ($selectedSources.Count -eq 0) {
            Add-ReadinessItem -Area "Sources" -Status "None selected" -Detail "Select at least one source." -Level "error"
        }
        else {
            Add-ReadinessItem -Area "Sources" -Status ("{0} enabled" -f $selectedSources.Count) -Detail (($selectedSources | ForEach-Object { $_.ShortLabel }) -join ", ")
        }

        $credentialIssues = @(Get-EnabledSourceCredentialIssues)
        if ($credentialIssues.Count -gt 0) {
            Add-ReadinessItem -Area "Credentials" -Status "Missing" -Detail ($credentialIssues -join "; ") -Level "warning"
        }
        else {
            Add-ReadinessItem -Area "Credentials" -Status "Ready" -Detail "Selected sources are usable."
        }
    }
    finally {
        $script:ReadinessList.EndUpdate()
    }
}

function Show-CredentialDialog {
    $rows = @(Get-JobCrawlerCredentialStatuses -SourcesConfig $script:CrawlerConfig.Sources | Where-Object { -not [string]::IsNullOrWhiteSpace($_.EnvironmentVariable) })
    if ($rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No credential environment variables are configured.", "Credentials") | Out-Null
        return
    }

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Set crawler credential"
    $dialog.StartPosition = "CenterParent"
    $dialog.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $dialog.MinimumSize = New-Object System.Drawing.Size(520, 280)
    $dialog.ClientSize = New-Object System.Drawing.Size(560, 280)
    $dialog.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $dialog.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 251)

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $layout.Padding = New-Object System.Windows.Forms.Padding(14)
    $layout.ColumnCount = 1
    $layout.RowCount = 5
    Add-ProfileTableColumn -Table $layout -Width 100
    Add-ProfileTableRow -Table $layout -Height 30 -Absolute
    Add-ProfileTableRow -Table $layout -Height 62 -Absolute
    Add-ProfileTableRow -Table $layout -Height 62 -Absolute
    Add-ProfileTableRow -Table $layout -Height 40 -Absolute
    Add-ProfileTableRow -Table $layout -Height 46 -Absolute
    $dialog.Controls.Add($layout)

    $intro = New-Object System.Windows.Forms.Label
    $intro.Text = "Select the variable to save in your Windows user environment."
    $intro.Dock = [System.Windows.Forms.DockStyle]::Fill
    $intro.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $layout.Controls.Add($intro, 0, 0)

    $combo = New-Object System.Windows.Forms.ComboBox
    $combo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $combo.Dock = [System.Windows.Forms.DockStyle]::Fill
    $combo.DisplayMember = "Display"
    foreach ($row in $rows) {
        $option = [PSCustomObject]@{
            Display = ("{0} ({1}/{2})" -f $row.EnvironmentVariable, $row.Source, $row.Credential)
            EnvironmentVariable = [string]$row.EnvironmentVariable
        }
        [void]$combo.Items.Add($option)
    }
    $missingIndex = 0
    for ($i = 0; $i -lt $rows.Count; $i++) {
        if ($rows[$i].Status -eq "missing") {
            $missingIndex = $i
            break
        }
    }
    $combo.SelectedIndex = $missingIndex
    $comboPanel = New-ProfileEditorFieldPanel -Label "Environment variable" -Control $combo
    $layout.Controls.Add($comboPanel, 0, 1)

    $valueBox = New-Object System.Windows.Forms.TextBox
    $valueBox.Dock = [System.Windows.Forms.DockStyle]::Fill
    $valueBox.UseSystemPasswordChar = $true

    $showCheckBox = New-Object System.Windows.Forms.CheckBox
    $showCheckBox.Text = "Show"
    $showCheckBox.Dock = [System.Windows.Forms.DockStyle]::Fill
    $showCheckBox.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $showCheckBox.Add_CheckedChanged({ $valueBox.UseSystemPasswordChar = -not $showCheckBox.Checked })

    $valueLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $valueLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $valueLayout.ColumnCount = 2
    $valueLayout.RowCount = 1
    Add-ProfileTableColumn -Table $valueLayout -Width 82
    Add-ProfileTableColumn -Table $valueLayout -Width 18
    Add-ProfileTableRow -Table $valueLayout -Height 100
    $valueLayout.Controls.Add($valueBox, 0, 0)
    $valueLayout.Controls.Add($showCheckBox, 1, 0)
    $valuePanel = New-ProfileEditorFieldPanel -Label "Value" -Control $valueLayout
    $layout.Controls.Add($valuePanel, 0, 2)

    $hint = New-Object System.Windows.Forms.Label
    $hint.Text = "The value is stored in Windows User environment variables, not in config JSON."
    $hint.Dock = [System.Windows.Forms.DockStyle]::Fill
    $hint.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $hint.ForeColor = [System.Drawing.Color]::FromArgb(77, 91, 114)
    $layout.Controls.Add($hint, 0, 3)

    $buttonBar = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttonBar.Dock = [System.Windows.Forms.DockStyle]::Fill
    $buttonBar.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
    $buttonBar.WrapContents = $false
    $layout.Controls.Add($buttonBar, 0, 4)

    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = "Save"
    $saveButton.Size = New-Object System.Drawing.Size(84, 30)
    $saveButton.Add_Click({ Invoke-LauncherAction -Context "Save credential" -Action {
        if ([string]::IsNullOrWhiteSpace($valueBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Enter a value before saving.", "Credentials") | Out-Null
            return
        }

        $selected = $combo.SelectedItem
        $envName = [string]$selected.EnvironmentVariable
        [Environment]::SetEnvironmentVariable($envName, $valueBox.Text, "User")
        [Environment]::SetEnvironmentVariable($envName, $valueBox.Text, "Process")
        Add-LogLine ("Credential saved to Windows User environment variable: {0}" -f $envName)
        Refresh-CredentialList
        $dialog.Close()
    } })

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Size = New-Object System.Drawing.Size(84, 30)
    $cancelButton.Add_Click({ $dialog.Close() })
    [void]$buttonBar.Controls.Add($cancelButton)
    [void]$buttonBar.Controls.Add($saveButton)

    $dialog.AcceptButton = $saveButton
    $dialog.CancelButton = $cancelButton
    [void]$dialog.ShowDialog($script:MainForm)
}

function Start-CrawlerFromGui {
    if ($null -ne $script:RunningProcess -and -not $script:RunningProcess.HasExited) {
        Add-LogLine "Crawler is already running."
        return
    }

    Refresh-ReadinessChecklist

    if (-not (Test-LauncherProfileConfigured)) {
        [System.Windows.Forms.MessageBox]::Show("Create a job profile first. Click New in the Profile section, save it, then run the crawler.", "Profile required") | Out-Null
        return
    }

    $validation = Test-JobCrawlerConfig -Config $script:CrawlerConfig
    if (-not $validation.IsValid) {
        [System.Windows.Forms.MessageBox]::Show((($validation.Issues) -join [Environment]::NewLine), "Config needs review") | Out-Null
        return
    }

    $workbook = Test-TrackerWorkbookWritable
    if ($workbook.Level -eq "error") {
        [System.Windows.Forms.MessageBox]::Show($workbook.Detail, "Workbook is open") | Out-Null
        Add-LogLine $workbook.Detail
        return
    }
    $writerPlan = Get-LauncherWorkbookWriterPlan
    if ($writerPlan.Level -eq "error") {
        [System.Windows.Forms.MessageBox]::Show($writerPlan.Detail, "Workbook writer") | Out-Null
        Add-LogLine $writerPlan.Detail
        return
    }

    $selectedSources = @(Get-SelectedSourceDefinitionsFromGui)
    if ($selectedSources.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select at least one source before running.", "Sources") | Out-Null
        return
    }

    $credentialIssues = @(Get-EnabledSourceCredentialIssues)
    if (-not $RunSmokeTest -and $credentialIssues.Count -gt 0) {
        $answer = [System.Windows.Forms.MessageBox]::Show(
            ("Some selected API sources are missing credentials:`n`n{0}`n`nContinue anyway?" -f ($credentialIssues -join "`n")),
            "Missing credentials",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }
    }

    $crawlerPath = Join-Path $script:CliDirectory "Find-AnalyticsJobs.ps1"
    $crawlerArguments = New-Object System.Collections.Generic.List[string]
    [void]$crawlerArguments.Add("-Profile")
    [void]$crawlerArguments.Add((Get-SelectedProfileId))
    [void]$crawlerArguments.Add("-DaysBack")
    [void]$crawlerArguments.Add(([string](Get-SelectedDaysBack)))
    [void]$crawlerArguments.Add("-CrawlMode")
    [void]$crawlerArguments.Add(([string]$script:ModeComboBox.SelectedItem))

    $enabledSourceKeys = New-Object System.Collections.Generic.List[string]
    $skippedSourceKeys = New-Object System.Collections.Generic.List[string]
    foreach ($sourceItem in $script:SourceListView.Items) {
        $metadata = $sourceItem.Tag
        $sourceChecked = Get-GuiCheckedValue $sourceItem
        $sourceKey = [string]$metadata.Key
        if ([string]::IsNullOrWhiteSpace($sourceKey)) {
            continue
        }

        if ($sourceChecked) {
            $enabledSourceKeys.Add($sourceKey) | Out-Null
        }
        else {
            $skippedSourceKeys.Add($sourceKey) | Out-Null
        }
    }
    if ($enabledSourceKeys.Count -gt 0) {
        [void]$crawlerArguments.Add("-EnableSource")
        [void]$crawlerArguments.Add((@($enabledSourceKeys.ToArray()) -join ","))
    }
    if ($skippedSourceKeys.Count -gt 0) {
        [void]$crawlerArguments.Add("-SkipSource")
        [void]$crawlerArguments.Add((@($skippedSourceKeys.ToArray()) -join ","))
    }
    if ($script:DryRunCheckBox.Checked) {
        [void]$crawlerArguments.Add("-DryRun")
    }
    if ($script:DiagnosticModeCheckBox.Checked) {
        [void]$crawlerArguments.Add("-DiagnosticMode")
    }
    if ($script:DisableCacheCheckBox.Checked) {
        [void]$crawlerArguments.Add("-DisableCache")
    }
    if ($RunSmokeTest) {
        [void]$crawlerArguments.Add("-SelfTest")
    }

    $script:LogTextBox.Clear()
    $script:RunLogLineCount = 0
    $script:LastCrawlerExitCode = $null
    $runLogPath = New-LauncherRunLogPath
    New-LauncherRunLock -RunId $script:LauncherRunId -LogPath $runLogPath

    Add-LogLine ("Starting crawl in {0} mode." -f $script:ModeComboBox.SelectedItem)
    Add-LogLine ("Published window: last {0} days." -f (Get-SelectedDaysBack))
    Add-LogLine ("Profile: {0} ({1})" -f $script:CrawlerConfig.Profile.Label, $script:CrawlerConfig.Profile.Id)
    Add-LogLine ("Output: {0} - {1}" -f $writerPlan.Status, $writerPlan.Detail)
    Add-LogLine ("Project: {0}" -f $script:ProjectRoot)
    Add-LogLine ("Live log: {0}" -f $script:LauncherRunLogPath)

    $script:RunScriptPath = New-LauncherRunScript -CrawlerPath $crawlerPath -CrawlerArguments @($crawlerArguments.ToArray())

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = Get-LauncherPowerShellPath
    $startInfo.WorkingDirectory = $script:ProjectRoot
    $startInfo.Arguments = "-NoLogo -NoProfile -ExecutionPolicy Bypass -File {0}" -f (ConvertTo-CommandLineArgument $script:RunScriptPath)
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $false
    $startInfo.RedirectStandardError = $false
    $startInfo.CreateNoWindow = $true

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo

    try {
        $script:RunningProcess = $process
        Set-LauncherRunningState -IsRunning $true
        [void]$process.Start()
        Set-LauncherRunLockProcessId -ProcessId $process.Id
        Start-LauncherRunPollTimer
    }
    catch {
        Stop-LauncherRunPollTimer
        Remove-LauncherRunScript
        Remove-LauncherRunLock
        $script:RunningProcess = $null
        Set-LauncherRunningState -IsRunning $false
        Set-LauncherStatus "Failed to start crawler."
        Add-LogLine ("Failed to start crawler: {0}" -f $_.Exception.Message)
    }
}

function Stop-CrawlerFromGui {
    if ($null -eq $script:RunningProcess -or $script:RunningProcess.HasExited) {
        Stop-LauncherRunPollTimer
        return
    }

    try {
        $script:RunningProcess.Kill()
        [void]$script:RunningProcess.WaitForExit(3000)
        Read-LauncherRunLog
        Add-LogLine "Crawler stopped by user."
        Set-LauncherStatus "Stopped."
    }
    catch {
        Add-LogLine ("Could not stop crawler: {0}" -f $_.Exception.Message)
    }
    finally {
        Stop-LauncherRunPollTimer
        Copy-LauncherRunLogToLastRun
        Remove-LauncherRunScript
        Remove-LauncherRunLock
        $script:RunningProcess = $null
        Set-LauncherRunningState -IsRunning $false
        Refresh-ReadinessChecklist
    }
}

function Test-ConfigFromGui {
    $validation = Test-JobCrawlerConfig -Config $script:CrawlerConfig
    if ($validation.IsValid) {
        Add-LogLine "Config validation passed."
        Set-LauncherStatus "Config validation passed."
    }
    else {
        Add-LogLine ("Config validation failed: {0}" -f (($validation.Issues) -join "; "))
        Set-LauncherStatus "Config validation failed."
    }
    Refresh-ReadinessChecklist
}

function Initialize-TrackerFromGui {
    if (-not (Test-LauncherProfileConfigured)) {
        [System.Windows.Forms.MessageBox]::Show("Create a job profile first. Click New in the Profile section, save it, then create the tracker.", "Profile required") | Out-Null
        return
    }

    $trackerPath = Get-LauncherDefaultTrackerPath
    if (Test-Path -LiteralPath $trackerPath) {
        [System.Windows.Forms.MessageBox]::Show("Tracker already exists: $trackerPath", "Create tracker") | Out-Null
        return
    }
    $writerPlan = Get-LauncherWorkbookWriterPlan
    if ($writerPlan.Level -eq "error") {
        [System.Windows.Forms.MessageBox]::Show($writerPlan.Detail, "Workbook writer") | Out-Null
        Add-LogLine $writerPlan.Detail
        return
    }

    $initializerPath = Join-Path $script:CliDirectory "Initialize-JobTracker.ps1"
    $arguments = @(
        "-NoLogo",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        (ConvertTo-CommandLineArgument $initializerPath),
        "-Profile",
        (ConvertTo-CommandLineArgument (Get-SelectedProfileId))
    ) -join " "

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = Get-LauncherPowerShellPath
    $startInfo.WorkingDirectory = $script:ProjectRoot
    $startInfo.Arguments = $arguments
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.CreateNoWindow = $true

    Add-LogLine ("Creating empty tracker workbook with {0}." -f $writerPlan.Status)
    try {
        $process = [System.Diagnostics.Process]::Start($startInfo)
        $stdout = $process.StandardOutput.ReadToEnd()
        $stderr = $process.StandardError.ReadToEnd()
        $process.WaitForExit()

        if (-not [string]::IsNullOrWhiteSpace($stdout)) {
            foreach ($line in ($stdout -split "`r?`n")) {
                if (-not [string]::IsNullOrWhiteSpace($line)) {
                    Add-LogLine $line
                }
            }
        }
        if (-not [string]::IsNullOrWhiteSpace($stderr)) {
            foreach ($line in ($stderr -split "`r?`n")) {
                if (-not [string]::IsNullOrWhiteSpace($line)) {
                    Add-LogLine ("ERROR: {0}" -f $line)
                }
            }
        }

        if ($process.ExitCode -eq 0) {
            Set-LauncherStatus "Tracker initialized."
            Refresh-ReadinessChecklist
        }
        else {
            Set-LauncherStatus ("Tracker initialization failed with error code {0}." -f $process.ExitCode)
        }
    }
    catch {
        Add-LogLine ("Could not initialize tracker: {0}" -f $_.Exception.Message)
        Set-LauncherStatus "Tracker initialization failed."
    }
}

function Get-LauncherOutputStatsText {
    $outputDirectory = Get-LauncherOutputDirectory
    $cacheDirectory = Resolve-JobCrawlerPath -BasePath $script:ProjectRoot -Path ([string](Get-ConfigPathValue -Object $script:CrawlerConfig.Runtime -Path "defaults.cache_directory" -DefaultValue "output\cache"))
    $cacheStats = Get-JobCrawlerDirectoryStats -Label "cache" -Path $cacheDirectory
    $logStats = Get-JobCrawlerDirectoryStats -Label "logs" -Path (Join-Path $outputDirectory "launcher_logs")
    $backupStats = Get-JobCrawlerDirectoryStats -Label "backups" -Path (Join-Path $outputDirectory "backups")
    return ("{0} | cache {1:N2} MB/{2} files | logs {3:N2} MB/{4} files | backups {5:N2} MB/{6} files" -f $outputDirectory, [double]$cacheStats.Megabytes, [int]$cacheStats.FileCount, [double]$logStats.Megabytes, [int]$logStats.FileCount, [double]$backupStats.Megabytes, [int]$backupStats.FileCount)
}

function Clean-ManagedOutputFromGui {
    $cacheDirectory = Resolve-JobCrawlerPath -BasePath $script:ProjectRoot -Path ([string](Get-ConfigPathValue -Object $script:CrawlerConfig.Runtime -Path "defaults.cache_directory" -DefaultValue "output\cache"))
    $script:JobCrawlerConfig = $script:CrawlerConfig
    $script:JobCrawlerRuntimeConfig = $script:CrawlerConfig.Runtime
    $script:CacheDirectory = $cacheDirectory
    $ageDays = 0
    Add-LogLine "Cleaning managed cache, logs, diagnostics, and backups now..."

    $results = @(Invoke-JobCrawlerOutputCleanup -ProjectRoot $script:ProjectRoot -CacheDirectory $cacheDirectory -All -OlderThanDays $ageDays)
    $removedFiles = 0
    $removedMb = [double]0
    foreach ($result in $results) {
        Add-LogLine ("Cleanup {0}: removed {1} file(s), {2:N2} MB." -f $result.Label, $result.RemovedFiles, [double]$result.RemovedMB)
        $removedFiles += [int]$result.RemovedFiles
        $removedMb += [double]$result.RemovedMB
    }

    Set-LauncherStatus ("Output cleanup finished: removed {0} file(s), {1:N2} MB." -f $removedFiles, $removedMb)
    Refresh-ReadinessChecklist
}

function Open-TrackerWorkbook {
    $path = Get-LauncherDefaultTrackerPath
    if (-not (Test-Path -LiteralPath $path)) {
        [System.Windows.Forms.MessageBox]::Show("Tracker does not exist yet: $path", "Open tracker") | Out-Null
        return
    }
    Start-Process -FilePath $path
}

function Open-OutputFolder {
    $path = Get-LauncherOutputDirectory
    if (-not (Test-Path -LiteralPath $path)) {
        New-Item -ItemType Directory -Path $path -Force | Out-Null
    }
    Start-Process -FilePath $path
}

$form = New-Object System.Windows.Forms.Form
$script:MainForm = $form
$form.Text = "Custom Job Tracker"
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(1040, 760)
$form.Size = New-Object System.Drawing.Size(1180, 820)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 251)

$mainLayout = New-Object System.Windows.Forms.TableLayoutPanel
$mainLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
$mainLayout.Padding = New-Object System.Windows.Forms.Padding(18, 12, 18, 16)
$mainLayout.ColumnCount = 1
$mainLayout.RowCount = 3
Add-ProfileTableColumn -Table $mainLayout -Width 100
Add-ProfileTableRow -Table $mainLayout -Height 72 -Absolute
Add-ProfileTableRow -Table $mainLayout -Height 164 -Absolute
Add-ProfileTableRow -Table $mainLayout -Height 100
$form.Controls.Add($mainLayout)

$headerPanel = New-Object System.Windows.Forms.TableLayoutPanel
$headerPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$headerPanel.ColumnCount = 1
$headerPanel.RowCount = 2
Add-ProfileTableColumn -Table $headerPanel -Width 100
Add-ProfileTableRow -Table $headerPanel -Height 40 -Absolute
Add-ProfileTableRow -Table $headerPanel -Height 28 -Absolute
$mainLayout.Controls.Add($headerPanel, 0, 0)

$header = New-Object System.Windows.Forms.Label
$header.Text = "Custom Job Tracker"
$header.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 18)
$header.Dock = [System.Windows.Forms.DockStyle]::Fill
$header.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$headerPanel.Controls.Add($header, 0, 0)

$subHeader = New-Object System.Windows.Forms.Label
$subHeader.Text = "Creates one tracker workbook per profile. Auto: Excel if available, otherwise built-in XLSX writer."
$subHeader.ForeColor = [System.Drawing.Color]::FromArgb(77, 91, 114)
$subHeader.Dock = [System.Windows.Forms.DockStyle]::Fill
$subHeader.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$headerPanel.Controls.Add($subHeader, 0, 1)

$runGroup = New-Object System.Windows.Forms.GroupBox
$script:SettingsGroup = $runGroup
$runGroup.Text = "Run setup"
$runGroup.Dock = [System.Windows.Forms.DockStyle]::Fill
$mainLayout.Controls.Add($runGroup, 0, 1)

$runTable = New-Object System.Windows.Forms.TableLayoutPanel
$runTable.Dock = [System.Windows.Forms.DockStyle]::Fill
$runTable.Padding = New-Object System.Windows.Forms.Padding(10, 8, 10, 10)
$runTable.ColumnCount = 5
$runTable.RowCount = 3
foreach ($width in @(32, 14, 16, 16, 22)) {
    Add-ProfileTableColumn -Table $runTable -Width $width
}
Add-ProfileTableRow -Table $runTable -Height 58 -Absolute
Add-ProfileTableRow -Table $runTable -Height 40 -Absolute
Add-ProfileTableRow -Table $runTable -Height 38 -Absolute
$runGroup.Controls.Add($runTable)

$profileCombo = New-Object System.Windows.Forms.ComboBox
$script:ProfileComboBox = $profileCombo
$profileCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$profileCombo.DisplayMember = "Display"
$profileCombo.Dock = [System.Windows.Forms.DockStyle]::Fill
$profileCombo.Add_SelectedIndexChanged({ Invoke-LauncherAction -Context "Select profile" -Action { Select-ProfileFromGui } })
Add-ProfileEditorField -Table $runTable -Label "Profile" -Control $profileCombo -Column 0 -Row 0 -ColumnSpan 2

$modeCombo = New-Object System.Windows.Forms.ComboBox
$script:ModeComboBox = $modeCombo
$modeCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$modeCombo.Dock = [System.Windows.Forms.DockStyle]::Fill
Add-ProfileEditorField -Table $runTable -Label "Crawl mode" -Control $modeCombo -Column 2 -Row 0

$daysBackCombo = New-Object System.Windows.Forms.ComboBox
$script:DaysBackComboBox = $daysBackCombo
$daysBackCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$daysBackCombo.DisplayMember = "Label"
$daysBackCombo.Dock = [System.Windows.Forms.DockStyle]::Fill
Add-ProfileEditorField -Table $runTable -Label "Published since" -Control $daysBackCombo -Column 3 -Row 0

$runButtons = New-Object System.Windows.Forms.FlowLayoutPanel
$runButtons.Dock = [System.Windows.Forms.DockStyle]::Fill
$runButtons.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$runButtons.WrapContents = $false
$runButtons.Padding = New-Object System.Windows.Forms.Padding(8, 22, 8, 4)
$runTable.Controls.Add($runButtons, 4, 0)

$script:RunButton = New-Object System.Windows.Forms.Button
$script:RunButton.Text = "Run crawl"
$script:RunButton.Size = New-Object System.Drawing.Size(104, 30)
$script:RunButton.BackColor = [System.Drawing.Color]::FromArgb(30, 95, 155)
$script:RunButton.ForeColor = [System.Drawing.Color]::White
$script:RunButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$script:RunButton.Add_Click({ Invoke-LauncherAction -Context "Run crawl" -Action { Start-CrawlerFromGui } })
[void]$runButtons.Controls.Add($script:RunButton)

$script:StopButton = New-Object System.Windows.Forms.Button
$script:StopButton.Text = "Stop"
$script:StopButton.Enabled = $false
$script:StopButton.Size = New-Object System.Drawing.Size(76, 30)
$script:StopButton.Add_Click({ Invoke-LauncherAction -Context "Stop crawl" -Action { Stop-CrawlerFromGui } })
[void]$runButtons.Controls.Add($script:StopButton)

$profileActions = New-Object System.Windows.Forms.FlowLayoutPanel
$profileActions.Dock = [System.Windows.Forms.DockStyle]::Fill
$profileActions.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$profileActions.WrapContents = $false
$profileActions.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 4)
$runTable.Controls.Add($profileActions, 0, 1)
$runTable.SetColumnSpan($profileActions, 2)

$newProfileButton = New-Object System.Windows.Forms.Button
$newProfileButton.Text = "New"
$newProfileButton.Size = New-Object System.Drawing.Size(74, 28)
$newProfileButton.Add_Click({ Invoke-LauncherAction -Context "Create profile" -Action { Show-NewProfileDialog } })
[void]$profileActions.Controls.Add($newProfileButton)
$script:ProfileActionButtons += $newProfileButton

$editProfileButton = New-Object System.Windows.Forms.Button
$editProfileButton.Text = "Edit"
$editProfileButton.Size = New-Object System.Drawing.Size(74, 28)
$editProfileButton.Add_Click({ Invoke-LauncherAction -Context "Edit profile" -Action { Show-EditProfileDialog } })
[void]$profileActions.Controls.Add($editProfileButton)
$script:ProfileActionButtons += $editProfileButton

$duplicateProfileButton = New-Object System.Windows.Forms.Button
$duplicateProfileButton.Text = "Duplicate"
$duplicateProfileButton.Size = New-Object System.Drawing.Size(94, 28)
$duplicateProfileButton.Add_Click({ Invoke-LauncherAction -Context "Duplicate profile" -Action { Show-DuplicateProfileDialog } })
[void]$profileActions.Controls.Add($duplicateProfileButton)
$script:ProfileActionButtons += $duplicateProfileButton

$defaultProfileButton = New-Object System.Windows.Forms.Button
$defaultProfileButton.Text = "Set default"
$defaultProfileButton.Size = New-Object System.Drawing.Size(96, 28)
$defaultProfileButton.Add_Click({ Invoke-LauncherAction -Context "Set default profile" -Action { Set-SelectedProfileAsDefault } })
[void]$profileActions.Controls.Add($defaultProfileButton)
$script:ProfileActionButtons += $defaultProfileButton

$deleteProfileButton = New-Object System.Windows.Forms.Button
$deleteProfileButton.Text = "Delete"
$deleteProfileButton.Size = New-Object System.Drawing.Size(74, 28)
$deleteProfileButton.Add_Click({ Invoke-LauncherAction -Context "Delete profile" -Action { Remove-SelectedProfileFromGui } })
[void]$profileActions.Controls.Add($deleteProfileButton)
$script:ProfileActionButtons += $deleteProfileButton

$optionsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$optionsPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$optionsPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$optionsPanel.WrapContents = $false
$optionsPanel.Padding = New-Object System.Windows.Forms.Padding(8, 5, 8, 4)
$runTable.Controls.Add($optionsPanel, 2, 1)
$runTable.SetColumnSpan($optionsPanel, 3)

$script:DryRunCheckBox = New-Object System.Windows.Forms.CheckBox
$script:DryRunCheckBox.Text = "Dry run"
$script:DryRunCheckBox.AutoSize = $true
$script:DryRunCheckBox.Visible = $false

$script:DiagnosticModeCheckBox = New-Object System.Windows.Forms.CheckBox
$script:DiagnosticModeCheckBox.Text = "Diagnostics"
$script:DiagnosticModeCheckBox.AutoSize = $true
$script:DiagnosticModeCheckBox.Visible = $false

$script:DisableCacheCheckBox = New-Object System.Windows.Forms.CheckBox
$script:DisableCacheCheckBox.Text = "Fresh fetch"
$script:DisableCacheCheckBox.AutoSize = $true
[void]$optionsPanel.Controls.Add($script:DisableCacheCheckBox)

$utilityButtons = New-Object System.Windows.Forms.FlowLayoutPanel
$utilityButtons.Dock = [System.Windows.Forms.DockStyle]::Fill
$utilityButtons.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$utilityButtons.WrapContents = $false
$utilityButtons.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 4)
$runTable.Controls.Add($utilityButtons, 0, 2)
$runTable.SetColumnSpan($utilityButtons, 5)

$validateButton = New-Object System.Windows.Forms.Button
$validateButton.Text = "Validate"
$validateButton.Size = New-Object System.Drawing.Size(78, 28)
$validateButton.Add_Click({ Invoke-LauncherAction -Context "Validate configuration" -Action { Test-ConfigFromGui } })
[void]$utilityButtons.Controls.Add($validateButton)

$createTrackerButton = New-Object System.Windows.Forms.Button
$createTrackerButton.Text = "Create tracker"
$createTrackerButton.Size = New-Object System.Drawing.Size(108, 28)
$createTrackerButton.Add_Click({ Invoke-LauncherAction -Context "Create tracker" -Action { Initialize-TrackerFromGui } })
[void]$utilityButtons.Controls.Add($createTrackerButton)

$openTrackerButton = New-Object System.Windows.Forms.Button
$openTrackerButton.Text = "Open tracker"
$openTrackerButton.Size = New-Object System.Drawing.Size(104, 28)
$openTrackerButton.Add_Click({ Invoke-LauncherAction -Context "Open tracker" -Action { Open-TrackerWorkbook } })
[void]$utilityButtons.Controls.Add($openTrackerButton)

$openOutputButton = New-Object System.Windows.Forms.Button
$openOutputButton.Text = "Output"
$openOutputButton.Size = New-Object System.Drawing.Size(74, 28)
$openOutputButton.Add_Click({ Invoke-LauncherAction -Context "Open output folder" -Action { Open-OutputFolder } })
[void]$utilityButtons.Controls.Add($openOutputButton)

$script:CleanupButton = New-Object System.Windows.Forms.Button
$script:CleanupButton.Text = "Clean output"
$script:CleanupButton.Size = New-Object System.Drawing.Size(104, 28)
$script:CleanupButton.Add_Click({ Invoke-LauncherAction -Context "Clean output" -Action { Clean-ManagedOutputFromGui } })
[void]$utilityButtons.Controls.Add($script:CleanupButton)

$bodySplit = New-Object System.Windows.Forms.SplitContainer
$bodySplit.Dock = [System.Windows.Forms.DockStyle]::Fill
$bodySplit.FixedPanel = [System.Windows.Forms.FixedPanel]::Panel1
$bodySplit.Panel1MinSize = 100
$bodySplit.Panel2MinSize = 100
$mainLayout.Controls.Add($bodySplit, 0, 2)

$leftTabs = New-Object System.Windows.Forms.TabControl
$leftTabs.Dock = [System.Windows.Forms.DockStyle]::Fill
$bodySplit.Panel1.Controls.Add($leftTabs)

$sourcesTab = New-Object System.Windows.Forms.TabPage
$sourcesTab.Text = "Sources"
$credentialsTab = New-Object System.Windows.Forms.TabPage
$credentialsTab.Text = "Credentials"
$readinessTab = New-Object System.Windows.Forms.TabPage
$readinessTab.Text = "Readiness"
[void]$leftTabs.TabPages.Add($sourcesTab)
[void]$leftTabs.TabPages.Add($credentialsTab)
[void]$leftTabs.TabPages.Add($readinessTab)

$sourceList = New-Object System.Windows.Forms.ListView
$script:SourceListView = $sourceList
$sourceList.Dock = [System.Windows.Forms.DockStyle]::Fill
$sourceList.View = [System.Windows.Forms.View]::Details
$sourceList.CheckBoxes = $true
$sourceList.FullRowSelect = $true
$sourceList.GridLines = $false
$sourceList.ShowItemToolTips = $true
[void]$sourceList.Columns.Add("Source", 100)
[void]$sourceList.Columns.Add("Type", 82)
[void]$sourceList.Columns.Add("Credentials", 126)
[void]$sourceList.Columns.Add("State", 98)
$sourceList.Add_ItemChecked({
    param($sender, $eventArgs)
    Invoke-LauncherAction -Context "Update source selection" -Action {
        if ($script:IsRefreshingSourceList) {
            return
        }
        $changedItem = $null
        try {
            $changedItem = $eventArgs.Item
        }
        catch {
            return
        }
        Update-GuiSourceListItem -Item $changedItem
        Refresh-ReadinessChecklist
    }
})
$sourcesTab.Controls.Add($sourceList)

$credentialsLayout = New-Object System.Windows.Forms.TableLayoutPanel
$credentialsLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
$credentialsLayout.Padding = New-Object System.Windows.Forms.Padding(10)
$credentialsLayout.ColumnCount = 1
$credentialsLayout.RowCount = 3
Add-ProfileTableColumn -Table $credentialsLayout -Width 100
Add-ProfileTableRow -Table $credentialsLayout -Height 32 -Absolute
Add-ProfileTableRow -Table $credentialsLayout -Height 100
Add-ProfileTableRow -Table $credentialsLayout -Height 44 -Absolute
$credentialsTab.Controls.Add($credentialsLayout)

$credentialHint = New-Object System.Windows.Forms.Label
$credentialHint.Text = "Secrets are stored in Windows User environment variables."
$credentialHint.Dock = [System.Windows.Forms.DockStyle]::Fill
$credentialHint.ForeColor = [System.Drawing.Color]::FromArgb(77, 91, 114)
$credentialsLayout.Controls.Add($credentialHint, 0, 0)

$credentialList = New-Object System.Windows.Forms.ListView
$script:CredentialList = $credentialList
$credentialList.View = [System.Windows.Forms.View]::Details
$credentialList.FullRowSelect = $true
$credentialList.GridLines = $false
$credentialList.Dock = [System.Windows.Forms.DockStyle]::Fill
[void]$credentialList.Columns.Add("Source", 108)
[void]$credentialList.Columns.Add("Name", 88)
[void]$credentialList.Columns.Add("Env var", 144)
[void]$credentialList.Columns.Add("Status", 70)
$credentialsLayout.Controls.Add($credentialList, 0, 1)

$credentialButtons = New-Object System.Windows.Forms.FlowLayoutPanel
$credentialButtons.Dock = [System.Windows.Forms.DockStyle]::Fill
$credentialButtons.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$credentialButtons.WrapContents = $false
$credentialsLayout.Controls.Add($credentialButtons, 0, 2)

$refreshCredentialsButton = New-Object System.Windows.Forms.Button
$refreshCredentialsButton.Text = "Refresh"
$refreshCredentialsButton.Size = New-Object System.Drawing.Size(82, 28)
$refreshCredentialsButton.Add_Click({ Invoke-LauncherAction -Context "Refresh credentials" -Action { Refresh-CredentialList } })
[void]$credentialButtons.Controls.Add($refreshCredentialsButton)

$setCredentialButton = New-Object System.Windows.Forms.Button
$setCredentialButton.Text = "Set credential"
$setCredentialButton.Size = New-Object System.Drawing.Size(116, 28)
$setCredentialButton.Add_Click({ Invoke-LauncherAction -Context "Set credential" -Action { Show-CredentialDialog } })
[void]$credentialButtons.Controls.Add($setCredentialButton)

$readinessList = New-Object System.Windows.Forms.ListView
$script:ReadinessList = $readinessList
$readinessList.Dock = [System.Windows.Forms.DockStyle]::Fill
$readinessList.View = [System.Windows.Forms.View]::Details
$readinessList.FullRowSelect = $true
$readinessList.GridLines = $false
[void]$readinessList.Columns.Add("Check", 86)
[void]$readinessList.Columns.Add("Status", 86)
[void]$readinessList.Columns.Add("Detail", 234)
$readinessTab.Controls.Add($readinessList)

$logGroup = New-Object System.Windows.Forms.GroupBox
$logGroup.Text = "Run log"
$logGroup.Dock = [System.Windows.Forms.DockStyle]::Fill
$bodySplit.Panel2.Controls.Add($logGroup)

$logLayout = New-Object System.Windows.Forms.TableLayoutPanel
$logLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
$logLayout.Padding = New-Object System.Windows.Forms.Padding(10, 8, 10, 10)
$logLayout.ColumnCount = 1
$logLayout.RowCount = 2
Add-ProfileTableColumn -Table $logLayout -Width 100
Add-ProfileTableRow -Table $logLayout -Height 100
Add-ProfileTableRow -Table $logLayout -Height 34 -Absolute
$logGroup.Controls.Add($logLayout)

$script:LogTextBox = New-Object System.Windows.Forms.TextBox
$script:LogTextBox.Multiline = $true
$script:LogTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$script:LogTextBox.ReadOnly = $true
$script:LogTextBox.WordWrap = $false
$script:LogTextBox.BackColor = [System.Drawing.Color]::White
$script:LogTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$script:LogTextBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$logLayout.Controls.Add($script:LogTextBox, 0, 0)

$statusPanel = New-Object System.Windows.Forms.TableLayoutPanel
$statusPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$statusPanel.ColumnCount = 2
$statusPanel.RowCount = 1
Add-ProfileTableColumn -Table $statusPanel -Width 65
Add-ProfileTableColumn -Table $statusPanel -Width 35
Add-ProfileTableRow -Table $statusPanel -Height 100
$logLayout.Controls.Add($statusPanel, 0, 1)

$script:StatusLabel = New-Object System.Windows.Forms.Label
$script:StatusLabel.Text = "Ready."
$script:StatusLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
$script:StatusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$statusPanel.Controls.Add($script:StatusLabel, 0, 0)

$script:ProgressBar = New-Object System.Windows.Forms.ProgressBar
$script:ProgressBar.Dock = [System.Windows.Forms.DockStyle]::Fill
$script:ProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks
$statusPanel.Controls.Add($script:ProgressBar, 1, 0)

$form.Add_FormClosing({
    param($sender, $eventArgs)

    Invoke-LauncherAction -Context "Close launcher" -Action {
        if ($null -ne $script:RunningProcess -and -not $script:RunningProcess.HasExited) {
            $answer = [System.Windows.Forms.MessageBox]::Show("The crawler is still running. Stop it and close?", "Crawler running", [System.Windows.Forms.MessageBoxButtons]::YesNo)
            if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
                Stop-CrawlerFromGui
            }
            else {
                $eventArgs.Cancel = $true
            }
        }
    }
})

$form.Add_Shown({
    Invoke-LauncherAction -Context "Show launcher" -Action {
        $panel1Min = 360
        $panel2Min = 420
        if ($bodySplit.Width -gt ($panel1Min + $panel2Min + 20)) {
            $targetDistance = [Math]::Min(420, ($bodySplit.Width - $panel2Min))
            $targetDistance = [Math]::Max($panel1Min, $targetDistance)
            $bodySplit.SplitterDistance = $targetDistance
            $bodySplit.Panel1MinSize = $panel1Min
            $bodySplit.Panel2MinSize = $panel2Min
        }
    }
})

Refresh-ProfileComboBox
Refresh-ModeComboBox
Refresh-DaysBackComboBox
Refresh-SourceCheckboxes
Refresh-CredentialList
Add-LogLine "Launcher ready."
Add-LogLine ("Profile: {0} ({1})" -f $script:CrawlerConfig.Profile.Label, $script:CrawlerConfig.Profile.Id)
Add-LogLine ("Tracker: {0}" -f (Get-LauncherDefaultTrackerPath))
if (-not (Test-Path -LiteralPath (Get-LauncherDefaultTrackerPath))) {
    Add-LogLine "First run: tracker workbook not found yet. Use Create tracker or Run crawl to create it."
}
if ($BuildSelfTest) {
    Write-Host "WinForms launcher build self-test passed."
    return
}
if ($SmokeTest) {
    $smokeTimer = New-Object System.Windows.Forms.Timer
    $smokeTimer.Interval = 750
    $smokeTimer.Add_Tick({
        $smokeTimer.Stop()
        $form.Close()
    })
    $form.Add_Shown({ $smokeTimer.Start() })
}
if ($RunSmokeTest) {
    $script:RunSmokeTestResult = ""
    $script:RunSmokeTestTicks = 0

    $runSmokeWatchTimer = New-Object System.Windows.Forms.Timer
    $runSmokeWatchTimer.Interval = 500
    $runSmokeWatchTimer.Add_Tick({
        Invoke-LauncherAction -Context "Run smoke test watch" -Action {
            $script:RunSmokeTestTicks++
            if ($null -eq $script:RunningProcess -and $script:StatusLabel.Text -like "Finished*") {
                if ($script:LastCrawlerExitCode -ne 0) {
                    $script:RunSmokeTestResult = "failed: crawler exited with code $script:LastCrawlerExitCode"
                }
                elseif ($form.Visible -and -not $form.IsDisposed) {
                    $script:RunSmokeTestResult = "passed"
                }
                else {
                    $script:RunSmokeTestResult = "failed: form closed before run completion could be observed"
                }
                $runSmokeWatchTimer.Stop()
                $form.Close()
                return
            }

            if ($script:RunSmokeTestTicks -gt 40) {
                $script:RunSmokeTestResult = "failed: timed out waiting for GUI run completion"
                $runSmokeWatchTimer.Stop()
                $form.Close()
            }
        }
    })

    $runSmokeStartTimer = New-Object System.Windows.Forms.Timer
    $runSmokeStartTimer.Interval = 500
    $runSmokeStartTimer.Add_Tick({
        Invoke-LauncherAction -Context "Run smoke test start" -Action {
            $runSmokeStartTimer.Stop()
            $script:DryRunCheckBox.Checked = $true
            Start-CrawlerFromGui
            $runSmokeWatchTimer.Start()
        }
    })
    $form.Add_Shown({ $runSmokeStartTimer.Start() })
}
[void][System.Windows.Forms.Application]::Run($form)
Copy-LauncherRunLogToLastRun
Remove-LauncherRunScript
Remove-LauncherRunLock
Clear-StaleLauncherRunScripts
if ($SmokeTest) {
    Write-Host "WinForms launcher smoke test passed."
}
if ($RunSmokeTest) {
    if ($script:RunSmokeTestResult -ne "passed") {
        if ([string]::IsNullOrWhiteSpace($script:RunSmokeTestResult)) {
            $script:RunSmokeTestResult = "failed: no result captured"
        }
        throw "WinForms launcher run smoke test $script:RunSmokeTestResult."
    }
    Write-Host "WinForms launcher run smoke test passed."
}
