[CmdletBinding()]
param(
    [switch]$SelfTest
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$script:ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:ConfigDirectory = Join-Path $script:ProjectRoot "config"

. (Join-Path $script:ProjectRoot "JobTracker.Common.ps1")
. (Join-Path $script:ProjectRoot "app\JobTracker.Config.ps1")

$script:CrawlerConfig = Get-JobCrawlerConfig -ConfigDirectory $script:ConfigDirectory
$script:ConfigValidation = Test-JobCrawlerConfig -Config $script:CrawlerConfig
if (-not $script:ConfigValidation.IsValid) {
    throw ("Invalid crawler config:`n- {0}" -f (($script:ConfigValidation.Issues) -join "`n- "))
}

function Get-LauncherDefaultTrackerPath {
    $relativePath = [string](Get-ConfigPathValue -Object $script:CrawlerConfig.Runtime -Path "defaults.tracker_path" -DefaultValue "output\jobs_tracker.xlsx")
    return Resolve-JobCrawlerPath -BasePath $script:ProjectRoot -Path $relativePath
}

function Get-LauncherOutputDirectory {
    $trackerPath = Get-LauncherDefaultTrackerPath
    return Split-Path -Parent $trackerPath
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

function ConvertTo-CommandLineSwitch {
    param([AllowNull()][string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return ""
    }
    if ($Name.StartsWith("-")) {
        return $Name
    }

    return "-$Name"
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
    throw "Start the GUI with powershell.exe -STA -File .\Launch-AnalyticsJobCrawlerGui.ps1, or double-click Run-AnalyticsJobCrawler-GUI.cmd."
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:MainForm = $null
$script:LogTextBox = $null
$script:RunButton = $null
$script:StopButton = $null
$script:ProgressBar = $null
$script:StatusLabel = $null
$script:CredentialList = $null
$script:RunningProcess = $null
$script:SourceCheckboxes = @()
$script:ModeComboBox = $null
$script:DryRunCheckBox = $null
$script:DiagnosticModeCheckBox = $null
$script:DisableCacheCheckBox = $null

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
    foreach ($checkbox in $script:SourceCheckboxes) {
        $checkbox.Enabled = -not $IsRunning
    }
    $script:DryRunCheckBox.Enabled = -not $IsRunning
    $script:DiagnosticModeCheckBox.Enabled = -not $IsRunning
    $script:DisableCacheCheckBox.Enabled = -not $IsRunning

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
    $dialog.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false
    $dialog.ClientSize = New-Object System.Drawing.Size(520, 220)
    $dialog.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Select the variable to save in your Windows user environment."
    $label.Location = New-Object System.Drawing.Point(16, 16)
    $label.Size = New-Object System.Drawing.Size(488, 22)
    $dialog.Controls.Add($label)

    $combo = New-Object System.Windows.Forms.ComboBox
    $combo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $combo.Location = New-Object System.Drawing.Point(16, 46)
    $combo.Size = New-Object System.Drawing.Size(488, 24)
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
    $dialog.Controls.Add($combo)

    $valueLabel = New-Object System.Windows.Forms.Label
    $valueLabel.Text = "Value"
    $valueLabel.Location = New-Object System.Drawing.Point(16, 86)
    $valueLabel.Size = New-Object System.Drawing.Size(80, 20)
    $dialog.Controls.Add($valueLabel)

    $valueBox = New-Object System.Windows.Forms.TextBox
    $valueBox.Location = New-Object System.Drawing.Point(16, 108)
    $valueBox.Size = New-Object System.Drawing.Size(400, 24)
    $valueBox.UseSystemPasswordChar = $true
    $dialog.Controls.Add($valueBox)

    $showCheckBox = New-Object System.Windows.Forms.CheckBox
    $showCheckBox.Text = "Show"
    $showCheckBox.Location = New-Object System.Drawing.Point(426, 109)
    $showCheckBox.Size = New-Object System.Drawing.Size(80, 22)
    $showCheckBox.Add_CheckedChanged({ $valueBox.UseSystemPasswordChar = -not $showCheckBox.Checked })
    $dialog.Controls.Add($showCheckBox)

    $hint = New-Object System.Windows.Forms.Label
    $hint.Text = "The value is stored in Windows User environment variables, not in config JSON."
    $hint.Location = New-Object System.Drawing.Point(16, 144)
    $hint.Size = New-Object System.Drawing.Size(488, 22)
    $hint.ForeColor = [System.Drawing.Color]::FromArgb(77, 91, 114)
    $dialog.Controls.Add($hint)

    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = "Save"
    $saveButton.Location = New-Object System.Drawing.Point(328, 178)
    $saveButton.Size = New-Object System.Drawing.Size(84, 30)
    $saveButton.Add_Click({
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
    })
    $dialog.Controls.Add($saveButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(420, 178)
    $cancelButton.Size = New-Object System.Drawing.Size(84, 30)
    $cancelButton.Add_Click({ $dialog.Close() })
    $dialog.Controls.Add($cancelButton)

    $dialog.AcceptButton = $saveButton
    $dialog.CancelButton = $cancelButton
    [void]$dialog.ShowDialog($script:MainForm)
}

function Start-CrawlerFromGui {
    if ($null -ne $script:RunningProcess -and -not $script:RunningProcess.HasExited) {
        Add-LogLine "Crawler is already running."
        return
    }

    $trackerPath = Get-LauncherDefaultTrackerPath
    if (Test-Path -LiteralPath $trackerPath) {
        try {
            $stream = [System.IO.File]::Open($trackerPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
            $stream.Close()
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Close jobs_tracker.xlsx before running the crawler.", "Workbook is open") | Out-Null
            Add-LogLine "Workbook appears to be open in Excel. Close it and run again."
            return
        }
    }

    $crawlerPath = Join-Path $script:ProjectRoot "Find-AnalyticsJobs.ps1"
    $arguments = New-Object System.Collections.Generic.List[string]
    [void]$arguments.Add("-NoLogo")
    [void]$arguments.Add("-NoProfile")
    [void]$arguments.Add("-ExecutionPolicy")
    [void]$arguments.Add("Bypass")
    [void]$arguments.Add("-File")
    [void]$arguments.Add((ConvertTo-CommandLineArgument $crawlerPath))
    [void]$arguments.Add("-CrawlMode")
    [void]$arguments.Add((ConvertTo-CommandLineArgument ([string]$script:ModeComboBox.SelectedItem)))

    $wttjPublicChecked = $false
    $welcomeKitChecked = $false
    foreach ($checkbox in $script:SourceCheckboxes) {
        $metadata = $checkbox.Tag
        $sourceKey = [string]$metadata.Key
        if ($sourceKey -eq "wttj_public") {
            $wttjPublicChecked = $checkbox.Checked
            $skipSwitch = ConvertTo-CommandLineSwitch ([string]$metadata.SkipSwitch)
            if (-not $checkbox.Checked -and -not [string]::IsNullOrWhiteSpace($skipSwitch)) {
                [void]$arguments.Add($skipSwitch)
            }
            continue
        }
        if ($sourceKey -eq "welcome_kit") {
            $welcomeKitChecked = $checkbox.Checked
            $skipSwitch = ConvertTo-CommandLineSwitch ([string]$metadata.SkipSwitch)
            $enableSwitch = ConvertTo-CommandLineSwitch ([string]$metadata.EnableSwitch)
            if ($checkbox.Checked -and -not [string]::IsNullOrWhiteSpace($enableSwitch)) {
                [void]$arguments.Add($enableSwitch)
            }
            elseif (-not $checkbox.Checked -and -not [string]::IsNullOrWhiteSpace($skipSwitch)) {
                [void]$arguments.Add($skipSwitch)
            }
            continue
        }

        $skipSwitch = ConvertTo-CommandLineSwitch ([string]$metadata.SkipSwitch)
        $enableSwitch = ConvertTo-CommandLineSwitch ([string]$metadata.EnableSwitch)
        if ($checkbox.Checked -and -not [string]::IsNullOrWhiteSpace($enableSwitch)) {
            [void]$arguments.Add($enableSwitch)
        }
        elseif (-not $checkbox.Checked -and -not [string]::IsNullOrWhiteSpace($skipSwitch)) {
            [void]$arguments.Add($skipSwitch)
        }
    }
    if (-not $wttjPublicChecked -and -not $welcomeKitChecked) {
        [void]$arguments.Add("-SkipWttj")
    }
    if ($script:DryRunCheckBox.Checked) {
        [void]$arguments.Add("-DryRun")
    }
    if ($script:DiagnosticModeCheckBox.Checked) {
        [void]$arguments.Add("-DiagnosticMode")
    }
    if ($script:DisableCacheCheckBox.Checked) {
        [void]$arguments.Add("-DisableCache")
    }

    $script:LogTextBox.Clear()
    Add-LogLine ("Starting crawl in {0} mode." -f $script:ModeComboBox.SelectedItem)
    Add-LogLine ("Project: {0}" -f $script:ProjectRoot)

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = Get-LauncherPowerShellPath
    $startInfo.WorkingDirectory = $script:ProjectRoot
    $startInfo.Arguments = ($arguments -join " ")
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.CreateNoWindow = $true

    try {
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        $startInfo.StandardOutputEncoding = $utf8
        $startInfo.StandardErrorEncoding = $utf8
    }
    catch {
    }

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    $process.EnableRaisingEvents = $true
    $process.add_OutputDataReceived({
        param($sender, $eventArgs)
        if (-not [string]::IsNullOrWhiteSpace($eventArgs.Data)) {
            Add-LogLine $eventArgs.Data
        }
    })
    $process.add_ErrorDataReceived({
        param($sender, $eventArgs)
        if (-not [string]::IsNullOrWhiteSpace($eventArgs.Data)) {
            Add-LogLine ("ERROR: {0}" -f $eventArgs.Data)
        }
    })
    $process.add_Exited({
        param($sender, $eventArgs)
        $exitCode = $sender.ExitCode
        if ($exitCode -eq 0) {
            Add-LogLine "Crawler finished successfully."
            Set-LauncherStatus "Finished successfully."
        }
        else {
            Add-LogLine ("Crawler finished with error code {0}." -f $exitCode)
            Set-LauncherStatus ("Finished with error code {0}." -f $exitCode)
        }
        Set-LauncherRunningState -IsRunning $false
    })

    try {
        $script:RunningProcess = $process
        Set-LauncherRunningState -IsRunning $true
        [void]$process.Start()
        $process.BeginOutputReadLine()
        $process.BeginErrorReadLine()
    }
    catch {
        Set-LauncherRunningState -IsRunning $false
        Set-LauncherStatus "Failed to start crawler."
        Add-LogLine ("Failed to start crawler: {0}" -f $_.Exception.Message)
    }
}

function Stop-CrawlerFromGui {
    if ($null -eq $script:RunningProcess -or $script:RunningProcess.HasExited) {
        return
    }

    try {
        $script:RunningProcess.Kill()
        Add-LogLine "Crawler stopped by user."
        Set-LauncherStatus "Stopped."
    }
    catch {
        Add-LogLine ("Could not stop crawler: {0}" -f $_.Exception.Message)
    }
    finally {
        Set-LauncherRunningState -IsRunning $false
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
}

function Initialize-TrackerFromGui {
    $trackerPath = Get-LauncherDefaultTrackerPath
    if (Test-Path -LiteralPath $trackerPath) {
        [System.Windows.Forms.MessageBox]::Show("Tracker already exists: $trackerPath", "Create tracker") | Out-Null
        return
    }

    $initializerPath = Join-Path $script:ProjectRoot "Initialize-JobTracker.ps1"
    $arguments = @(
        "-NoLogo",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        (ConvertTo-CommandLineArgument $initializerPath)
    ) -join " "

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = Get-LauncherPowerShellPath
    $startInfo.WorkingDirectory = $script:ProjectRoot
    $startInfo.Arguments = $arguments
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.CreateNoWindow = $true

    Add-LogLine "Creating empty tracker workbook..."
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
$form.Text = "Analytics Job Crawler"
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(980, 700)
$form.Size = New-Object System.Drawing.Size(1040, 760)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 251)

$header = New-Object System.Windows.Forms.Label
$header.Text = "Analytics Job Crawler"
$header.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 18)
$header.Location = New-Object System.Drawing.Point(18, 16)
$header.Size = New-Object System.Drawing.Size(500, 34)
$form.Controls.Add($header)

$subHeader = New-Object System.Windows.Forms.Label
$subHeader.Text = "Manual crawl launcher for output\jobs_tracker.xlsx. Close the workbook before running."
$subHeader.ForeColor = [System.Drawing.Color]::FromArgb(77, 91, 114)
$subHeader.Location = New-Object System.Drawing.Point(21, 52)
$subHeader.Size = New-Object System.Drawing.Size(760, 22)
$form.Controls.Add($subHeader)

$settingsGroup = New-Object System.Windows.Forms.GroupBox
$settingsGroup.Text = "Crawl setup"
$settingsGroup.Location = New-Object System.Drawing.Point(18, 88)
$settingsGroup.Size = New-Object System.Drawing.Size(350, 256)
$settingsGroup.Anchor = "Top,Left"
$form.Controls.Add($settingsGroup)

$modeLabel = New-Object System.Windows.Forms.Label
$modeLabel.Text = "Mode"
$modeLabel.Location = New-Object System.Drawing.Point(16, 30)
$modeLabel.Size = New-Object System.Drawing.Size(70, 22)
$settingsGroup.Controls.Add($modeLabel)

$modeCombo = New-Object System.Windows.Forms.ComboBox
$script:ModeComboBox = $modeCombo
$modeCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$modeCombo.Location = New-Object System.Drawing.Point(92, 27)
$modeCombo.Size = New-Object System.Drawing.Size(130, 24)
$modeNames = @($script:CrawlerConfig.CrawlModes.modes.PSObject.Properties.Name)
if ($modeNames.Count -eq 0) {
    $modeNames = @("Fast", "Default", "Deep")
}
foreach ($modeName in $modeNames) {
    [void]$modeCombo.Items.Add([string]$modeName)
}
$defaultMode = [string](Get-ConfigPathValue -Object $script:CrawlerConfig.Runtime -Path "defaults.crawl_mode" -DefaultValue "Default")
if ($modeCombo.Items.Contains($defaultMode)) {
    $modeCombo.SelectedItem = $defaultMode
}
else {
    $modeCombo.SelectedItem = "Default"
}
$settingsGroup.Controls.Add($modeCombo)

$sourceLabel = New-Object System.Windows.Forms.Label
$sourceLabel.Text = "Sources"
$sourceLabel.Location = New-Object System.Drawing.Point(16, 66)
$sourceLabel.Size = New-Object System.Drawing.Size(100, 22)
$settingsGroup.Controls.Add($sourceLabel)

$sourceDefinitions = @(Get-JobCrawlerSourceDefinitions -SourcesConfig $script:CrawlerConfig.Sources)

$sourceTop = 92
for ($i = 0; $i -lt $sourceDefinitions.Count; $i++) {
    $definition = $sourceDefinitions[$i]
    $checkbox = New-Object System.Windows.Forms.CheckBox
    $checkbox.Text = [string]$definition.ShortLabel
    $checkbox.Tag = $definition
    $checkbox.Checked = [bool]$definition.EnabledByDefault
    $checkbox.Location = New-Object System.Drawing.Point((16 + (($i % 2) * 160)), ($sourceTop + ([Math]::Floor($i / 2) * 26)))
    $checkbox.Size = New-Object System.Drawing.Size(154, 24)
    $settingsGroup.Controls.Add($checkbox)
    $script:SourceCheckboxes += $checkbox
}

$script:DryRunCheckBox = New-Object System.Windows.Forms.CheckBox
$script:DryRunCheckBox.Text = "Dry run"
$script:DryRunCheckBox.Location = New-Object System.Drawing.Point(16, 190)
$script:DryRunCheckBox.Size = New-Object System.Drawing.Size(90, 24)
$settingsGroup.Controls.Add($script:DryRunCheckBox)

$script:DiagnosticModeCheckBox = New-Object System.Windows.Forms.CheckBox
$script:DiagnosticModeCheckBox.Text = "Diagnostics"
$script:DiagnosticModeCheckBox.Location = New-Object System.Drawing.Point(116, 190)
$script:DiagnosticModeCheckBox.Size = New-Object System.Drawing.Size(100, 24)
$settingsGroup.Controls.Add($script:DiagnosticModeCheckBox)

$script:DisableCacheCheckBox = New-Object System.Windows.Forms.CheckBox
$script:DisableCacheCheckBox.Text = "Fresh fetch"
$script:DisableCacheCheckBox.Location = New-Object System.Drawing.Point(226, 190)
$script:DisableCacheCheckBox.Size = New-Object System.Drawing.Size(100, 24)
$settingsGroup.Controls.Add($script:DisableCacheCheckBox)

$script:RunButton = New-Object System.Windows.Forms.Button
$script:RunButton.Text = "Run crawl"
$script:RunButton.Location = New-Object System.Drawing.Point(16, 220)
$script:RunButton.Size = New-Object System.Drawing.Size(96, 28)
$script:RunButton.BackColor = [System.Drawing.Color]::FromArgb(30, 95, 155)
$script:RunButton.ForeColor = [System.Drawing.Color]::White
$script:RunButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$script:RunButton.Add_Click({ Start-CrawlerFromGui })
$settingsGroup.Controls.Add($script:RunButton)

$script:StopButton = New-Object System.Windows.Forms.Button
$script:StopButton.Text = "Stop"
$script:StopButton.Enabled = $false
$script:StopButton.Location = New-Object System.Drawing.Point(120, 220)
$script:StopButton.Size = New-Object System.Drawing.Size(72, 28)
$script:StopButton.Add_Click({ Stop-CrawlerFromGui })
$settingsGroup.Controls.Add($script:StopButton)

$validateButton = New-Object System.Windows.Forms.Button
$validateButton.Text = "Validate config"
$validateButton.Location = New-Object System.Drawing.Point(200, 220)
$validateButton.Size = New-Object System.Drawing.Size(116, 28)
$validateButton.Add_Click({ Test-ConfigFromGui })
$settingsGroup.Controls.Add($validateButton)

$credentialsGroup = New-Object System.Windows.Forms.GroupBox
$credentialsGroup.Text = "Credentials"
$credentialsGroup.Location = New-Object System.Drawing.Point(18, 356)
$credentialsGroup.Size = New-Object System.Drawing.Size(350, 244)
$form.Controls.Add($credentialsGroup)

$credentialHint = New-Object System.Windows.Forms.Label
$credentialHint.Text = "Secrets are stored in Windows User environment variables."
$credentialHint.Location = New-Object System.Drawing.Point(14, 24)
$credentialHint.Size = New-Object System.Drawing.Size(318, 20)
$credentialHint.ForeColor = [System.Drawing.Color]::FromArgb(77, 91, 114)
$credentialsGroup.Controls.Add($credentialHint)

$credentialList = New-Object System.Windows.Forms.ListView
$script:CredentialList = $credentialList
$credentialList.View = [System.Windows.Forms.View]::Details
$credentialList.FullRowSelect = $true
$credentialList.GridLines = $false
$credentialList.Location = New-Object System.Drawing.Point(14, 50)
$credentialList.Size = New-Object System.Drawing.Size(318, 140)
[void]$credentialList.Columns.Add("Source", 88)
[void]$credentialList.Columns.Add("Name", 74)
[void]$credentialList.Columns.Add("Env var", 106)
[void]$credentialList.Columns.Add("Status", 50)
$credentialsGroup.Controls.Add($credentialList)

$refreshCredentialsButton = New-Object System.Windows.Forms.Button
$refreshCredentialsButton.Text = "Refresh"
$refreshCredentialsButton.Location = New-Object System.Drawing.Point(14, 202)
$refreshCredentialsButton.Size = New-Object System.Drawing.Size(76, 28)
$refreshCredentialsButton.Add_Click({ Refresh-CredentialList })
$credentialsGroup.Controls.Add($refreshCredentialsButton)

$setCredentialButton = New-Object System.Windows.Forms.Button
$setCredentialButton.Text = "Set credential"
$setCredentialButton.Location = New-Object System.Drawing.Point(98, 202)
$setCredentialButton.Size = New-Object System.Drawing.Size(110, 28)
$setCredentialButton.Add_Click({ Show-CredentialDialog })
$credentialsGroup.Controls.Add($setCredentialButton)

$openTrackerButton = New-Object System.Windows.Forms.Button
$openTrackerButton.Text = "Open tracker"
$openTrackerButton.Location = New-Object System.Drawing.Point(18, 614)
$openTrackerButton.Size = New-Object System.Drawing.Size(108, 30)
$openTrackerButton.Add_Click({ Open-TrackerWorkbook })
$form.Controls.Add($openTrackerButton)

$openOutputButton = New-Object System.Windows.Forms.Button
$openOutputButton.Text = "Open output"
$openOutputButton.Location = New-Object System.Drawing.Point(136, 614)
$openOutputButton.Size = New-Object System.Drawing.Size(108, 30)
$openOutputButton.Add_Click({ Open-OutputFolder })
$form.Controls.Add($openOutputButton)

$createTrackerButton = New-Object System.Windows.Forms.Button
$createTrackerButton.Text = "Create tracker"
$createTrackerButton.Location = New-Object System.Drawing.Point(254, 614)
$createTrackerButton.Size = New-Object System.Drawing.Size(114, 30)
$createTrackerButton.Add_Click({ Initialize-TrackerFromGui })
$form.Controls.Add($createTrackerButton)

$logGroup = New-Object System.Windows.Forms.GroupBox
$logGroup.Text = "Run log"
$logGroup.Location = New-Object System.Drawing.Point(386, 88)
$logGroup.Size = New-Object System.Drawing.Size(620, 556)
$logGroup.Anchor = "Top,Left,Right,Bottom"
$form.Controls.Add($logGroup)

$script:LogTextBox = New-Object System.Windows.Forms.TextBox
$script:LogTextBox.Multiline = $true
$script:LogTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$script:LogTextBox.ReadOnly = $true
$script:LogTextBox.WordWrap = $false
$script:LogTextBox.BackColor = [System.Drawing.Color]::White
$script:LogTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$script:LogTextBox.Location = New-Object System.Drawing.Point(14, 24)
$script:LogTextBox.Size = New-Object System.Drawing.Size(592, 486)
$script:LogTextBox.Anchor = "Top,Left,Right,Bottom"
$logGroup.Controls.Add($script:LogTextBox)

$script:StatusLabel = New-Object System.Windows.Forms.Label
$script:StatusLabel.Text = "Ready."
$script:StatusLabel.Location = New-Object System.Drawing.Point(14, 518)
$script:StatusLabel.Size = New-Object System.Drawing.Size(360, 22)
$script:StatusLabel.Anchor = "Left,Right,Bottom"
$logGroup.Controls.Add($script:StatusLabel)

$script:ProgressBar = New-Object System.Windows.Forms.ProgressBar
$script:ProgressBar.Location = New-Object System.Drawing.Point(382, 518)
$script:ProgressBar.Size = New-Object System.Drawing.Size(224, 18)
$script:ProgressBar.Anchor = "Right,Bottom"
$script:ProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks
$logGroup.Controls.Add($script:ProgressBar)

$form.Add_FormClosing({
    param($sender, $eventArgs)

    if ($null -ne $script:RunningProcess -and -not $script:RunningProcess.HasExited) {
        $answer = [System.Windows.Forms.MessageBox]::Show("The crawler is still running. Stop it and close?", "Crawler running", [System.Windows.Forms.MessageBoxButtons]::YesNo)
        if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
            Stop-CrawlerFromGui
        }
        else {
            $eventArgs.Cancel = $true
        }
    }
})

Refresh-CredentialList
Add-LogLine "Launcher ready."
Add-LogLine ("Tracker: {0}" -f (Get-LauncherDefaultTrackerPath))
if (-not (Test-Path -LiteralPath (Get-LauncherDefaultTrackerPath))) {
    Add-LogLine "First run: tracker workbook not found yet. Use Create tracker or Run crawl to create it."
}
[void][System.Windows.Forms.Application]::Run($form)
