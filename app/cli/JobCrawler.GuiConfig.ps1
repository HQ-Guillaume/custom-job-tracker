# Profile, mode, and source-control helpers for the WinForms launcher.

function Get-SelectedProfileId {
    if ($null -ne $script:ProfileComboBox -and $null -ne $script:ProfileComboBox.SelectedItem) {
        return [string]$script:ProfileComboBox.SelectedItem.Id
    }

    return [string]$script:CrawlerConfig.Profile.Id
}

function Get-GuiCheckedValue {
    param([AllowNull()]$Control)

    if ($null -eq $Control) {
        return $false
    }

    try {
        return [bool]$Control.Checked
    }
    catch {
        return $false
    }
}

function Refresh-ModeComboBox {
    if ($null -eq $script:ModeComboBox) {
        return
    }

    $selectedMode = [string]$script:ModeComboBox.SelectedItem
    $script:ModeComboBox.Items.Clear()
    $modeNames = @($script:CrawlerConfig.CrawlModes.modes.PSObject.Properties.Name)
    if ($modeNames.Count -eq 0) {
        $modeNames = @("Fast", "Default", "Deep")
    }
    foreach ($modeName in $modeNames) {
        [void]$script:ModeComboBox.Items.Add([string]$modeName)
    }

    $defaultMode = [string](Get-ConfigPathValue -Object $script:CrawlerConfig.Runtime -Path "defaults.crawl_mode" -DefaultValue "Default")
    if (-not [string]::IsNullOrWhiteSpace($selectedMode) -and $script:ModeComboBox.Items.Contains($selectedMode)) {
        $script:ModeComboBox.SelectedItem = $selectedMode
    }
    elseif ($script:ModeComboBox.Items.Contains($defaultMode)) {
        $script:ModeComboBox.SelectedItem = $defaultMode
    }
    else {
        $script:ModeComboBox.SelectedIndex = 0
    }
}

function Get-GuiDaysBackOptions {
    $rawValues = @(Get-ConfigPathValue -Object $script:CrawlerConfig.Runtime -Path "defaults.allowed_days_back" -DefaultValue @(7, 14, 30))
    $values = New-Object System.Collections.Generic.List[int]
    foreach ($rawValue in $rawValues) {
        try {
            $value = [int]$rawValue
        }
        catch {
            continue
        }
        if ($value -lt 1 -or $value -gt 365 -or $values.Contains($value)) {
            continue
        }
        $values.Add($value) | Out-Null
    }
    if ($values.Count -eq 0) {
        foreach ($fallback in @(7, 14, 30)) {
            $values.Add([int]$fallback) | Out-Null
        }
    }

    $options = New-Object System.Collections.Generic.List[object]
    foreach ($value in @($values.ToArray() | Sort-Object)) {
        $label = "{0} days" -f $value
        if ($value -eq 14) {
            $label = "14 days (2 weeks)"
        }
        elseif ($value -eq 30) {
            $label = "30 days (1 month)"
        }
        [void]$options.Add([PSCustomObject]@{ Label = $label; Value = $value })
    }

    return @($options.ToArray())
}

function Get-SelectedDaysBack {
    $defaultDays = [int](Get-ConfigPathValue -Object $script:CrawlerConfig.Runtime -Path "defaults.days_back" -DefaultValue 7)
    if ($null -ne $script:DaysBackComboBox -and $null -ne $script:DaysBackComboBox.SelectedItem) {
        $selectedValue = Get-ConfigProperty -Object $script:DaysBackComboBox.SelectedItem -Name "Value" -DefaultValue $defaultDays
        try {
            return [int]$selectedValue
        }
        catch {
            return $defaultDays
        }
    }

    return $defaultDays
}

function Refresh-DaysBackComboBox {
    if ($null -eq $script:DaysBackComboBox) {
        return
    }

    $previousDays = Get-SelectedDaysBack
    $defaultDays = [int](Get-ConfigPathValue -Object $script:CrawlerConfig.Runtime -Path "defaults.days_back" -DefaultValue 7)
    $targetDays = $(if ($previousDays -gt 0) { $previousDays } else { $defaultDays })

    $script:DaysBackComboBox.Items.Clear()
    foreach ($option in @(Get-GuiDaysBackOptions)) {
        [void]$script:DaysBackComboBox.Items.Add($option)
    }
    if ($script:DaysBackComboBox.Items.Count -eq 0) {
        return
    }

    $selectedIndex = -1
    for ($i = 0; $i -lt $script:DaysBackComboBox.Items.Count; $i++) {
        $value = [int](Get-ConfigProperty -Object $script:DaysBackComboBox.Items[$i] -Name "Value" -DefaultValue 0)
        if ($value -eq $targetDays) {
            $selectedIndex = $i
            break
        }
    }
    if ($selectedIndex -lt 0) {
        for ($i = 0; $i -lt $script:DaysBackComboBox.Items.Count; $i++) {
            $value = [int](Get-ConfigProperty -Object $script:DaysBackComboBox.Items[$i] -Name "Value" -DefaultValue 0)
            if ($value -eq $defaultDays) {
                $selectedIndex = $i
                break
            }
        }
    }
    $script:DaysBackComboBox.SelectedIndex = $(if ($selectedIndex -ge 0) { $selectedIndex } else { 0 })
}

function Get-GuiSourceCredentialSummary {
    param($Definition)

    if (-not [bool]$Definition.RequiresCredential) {
        return "Not required"
    }

    $rows = @(Get-JobCrawlerCredentialStatuses -SourcesConfig $script:CrawlerConfig.Sources | Where-Object { $_.Source -eq [string]$Definition.Key })
    if ($rows.Count -eq 0) {
        return "No variables"
    }

    $missing = @($rows | Where-Object { $_.Status -eq "missing" })
    if ($missing.Count -gt 0) {
        return "Missing: {0}" -f (($missing | ForEach-Object { $_.Credential }) -join ", ")
    }

    $defaults = @($rows | Where-Object { $_.Status -eq "default" })
    if ($defaults.Count -gt 0) {
        return "Ready, defaults used"
    }

    return "Ready"
}

function Get-GuiSourceKind {
    param($Definition)

    if ([bool]$Definition.RequiresCredential) {
        return "API"
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$Definition.FallbackFor)) {
        return "Public fallback"
    }

    return "Public"
}

function Test-GuiSourceMissingCredentials {
    param($Definition)

    if (-not [bool]$Definition.RequiresCredential) {
        return $false
    }

    $rows = @(Get-JobCrawlerCredentialStatuses -SourcesConfig $script:CrawlerConfig.Sources | Where-Object { $_.Source -eq [string]$Definition.Key })
    return @($rows | Where-Object { $_.Status -eq "missing" }).Count -gt 0
}

function Update-GuiSourceListItem {
    param($Item)

    if ($null -eq $Item -or $null -eq $Item.Tag) {
        return
    }

    $definition = $Item.Tag
    while ($Item.SubItems.Count -lt 4) {
        [void]$Item.SubItems.Add("")
    }

    $credentialSummary = Get-GuiSourceCredentialSummary -Definition $definition
    $missingCredentials = Test-GuiSourceMissingCredentials -Definition $definition
    $kind = Get-GuiSourceKind -Definition $definition
    $isChecked = Get-GuiCheckedValue $Item
    $state = $(if ($isChecked) { "Enabled" } else { "Disabled" })
    if ($isChecked -and $missingCredentials) {
        $state = "Missing credentials"
    }
    elseif ($isChecked -and -not [string]::IsNullOrWhiteSpace([string]$definition.FallbackFor)) {
        $state = "Fallback enabled"
    }
    elseif (-not $isChecked -and [bool]$definition.RequiresCredential -and -not $missingCredentials) {
        $state = "Available, disabled"
    }
    elseif (-not $isChecked -and [bool]$definition.RequiresCredential -and $missingCredentials) {
        $state = "Disabled, missing credentials"
    }

    $Item.SubItems[1].Text = $kind
    $Item.SubItems[2].Text = $credentialSummary
    $Item.SubItems[3].Text = $state
    $Item.ToolTipText = ("{0} | {1} | {2}" -f $definition.Label, $kind, $credentialSummary)

    if ($isChecked -and $missingCredentials) {
        $Item.ForeColor = [System.Drawing.Color]::FromArgb(154, 69, 59)
    }
    elseif (-not $isChecked) {
        $Item.ForeColor = [System.Drawing.Color]::FromArgb(94, 110, 130)
    }
    else {
        $Item.ForeColor = [System.Drawing.Color]::FromArgb(30, 104, 72)
    }
}

function Refresh-SourceCheckboxes {
    if ($null -eq $script:SourceListView) {
        return
    }

    $script:IsRefreshingSourceList = $true
    try {
        $script:SourceListView.BeginUpdate()
        $script:SourceListView.Items.Clear()
        $script:SourceCheckboxes = @()

        foreach ($definition in @(Get-JobCrawlerSourceDefinitions -SourcesConfig $script:CrawlerConfig.Sources)) {
            $item = New-Object System.Windows.Forms.ListViewItem([string]$definition.ShortLabel)
            $item.Tag = $definition
            $item.Checked = [bool]$definition.EnabledByDefault
            [void]$item.SubItems.Add("")
            [void]$item.SubItems.Add("")
            [void]$item.SubItems.Add("")
            Update-GuiSourceListItem -Item $item
            [void]$script:SourceListView.Items.Add($item)
        }
    }
    finally {
        $script:SourceListView.EndUpdate()
        $script:IsRefreshingSourceList = $false
    }

    if (Get-Command Refresh-ReadinessChecklist -ErrorAction SilentlyContinue) {
        Refresh-ReadinessChecklist
    }
}

function Refresh-ProfileComboBox {
    param([string]$SelectedProfileId = "")

    if ($null -eq $script:ProfileComboBox) {
        return
    }

    $selectedId = ConvertTo-JobCrawlerProfileId $(if ([string]::IsNullOrWhiteSpace($SelectedProfileId)) { [string]$script:CrawlerConfig.Profile.Id } else { $SelectedProfileId })
    $script:IsRefreshingProfileCombo = $true
    try {
        $script:ProfileOptions = @(Get-JobCrawlerProfileSummaries -ConfigDirectory $script:ConfigDirectory | Sort-Object Label, Id)
        $script:ProfileComboBox.Items.Clear()
        foreach ($profile in $script:ProfileOptions) {
            [void]$script:ProfileComboBox.Items.Add($profile)
        }

        for ($i = 0; $i -lt $script:ProfileComboBox.Items.Count; $i++) {
            if ([string]$script:ProfileComboBox.Items[$i].Id -eq $selectedId) {
                $script:ProfileComboBox.SelectedIndex = $i
                break
            }
        }
        if ($script:ProfileComboBox.SelectedIndex -lt 0 -and $script:ProfileComboBox.Items.Count -gt 0) {
            $script:ProfileComboBox.SelectedIndex = 0
        }
    }
    finally {
        $script:IsRefreshingProfileCombo = $false
    }
}

function Select-ProfileFromGui {
    if ($script:IsRefreshingProfileCombo) {
        return
    }

    $profileId = Get-SelectedProfileId
    try {
        Set-LauncherCrawlerConfig -ProfileId $profileId
        Set-LauncherOutputPaths
        Refresh-ModeComboBox
        Refresh-DaysBackComboBox
        Refresh-SourceCheckboxes
        Refresh-CredentialList
        Set-LauncherStatus ("Profile selected: {0}" -f $script:CrawlerConfig.Profile.Label)
        Add-LogLine ("Profile selected: {0} ({1})" -f $script:CrawlerConfig.Profile.Label, $script:CrawlerConfig.Profile.Id)
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Profile") | Out-Null
        Add-LogLine ("Profile selection failed: {0}" -f $_.Exception.Message)
    }
}

function Set-SelectedProfileAsDefault {
    $profileId = Get-SelectedProfileId
    if ([string]::IsNullOrWhiteSpace($profileId)) {
        [System.Windows.Forms.MessageBox]::Show("Create or select a profile first.", "Profile") | Out-Null
        return
    }

    try {
        $path = Set-JobCrawlerDefaultProfile -ConfigDirectory $script:ConfigDirectory -ProfileId $profileId
        Add-LogLine ("Default profile saved: {0}" -f $profileId)
        Add-LogLine ("Local config: {0}" -f $path)
        Set-LauncherStatus "Default profile saved."
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Profile") | Out-Null
    }
}

function Remove-SelectedProfileFromGui {
    if ($null -eq $script:ProfileComboBox -or $null -eq $script:ProfileComboBox.SelectedItem) {
        [System.Windows.Forms.MessageBox]::Show("Select a profile first.", "Profile") | Out-Null
        return
    }

    $selectedProfile = $script:ProfileComboBox.SelectedItem
    $profileId = [string]$selectedProfile.Id
    $profileLabel = [string]$selectedProfile.Label
    if ([string]::IsNullOrWhiteSpace($profileId)) {
        return
    }
    if (-not [bool]$selectedProfile.IsLocal) {
        [System.Windows.Forms.MessageBox]::Show("Only local profiles can be deleted. Public profiles are part of the installed project.", "Profile") | Out-Null
        return
    }

    $answer = [System.Windows.Forms.MessageBox]::Show(
        ("Delete profile '{0}'?`n`nThe profile file will be removed, but its tracker workbook will be kept." -f $profileLabel),
        "Delete profile",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    try {
        $deletedPath = Remove-JobCrawlerLocalProfile -ConfigDirectory $script:ConfigDirectory -ProfileId $profileId
        $remainingProfiles = @(Get-JobCrawlerProfileSummaries -ConfigDirectory $script:ConfigDirectory | Sort-Object Label, Id)
        $nextProfileId = ""
        if ($remainingProfiles.Count -gt 0) {
            $nextProfileId = [string]$remainingProfiles[0].Id
            [void](Set-JobCrawlerDefaultProfile -ConfigDirectory $script:ConfigDirectory -ProfileId $nextProfileId)
        }
        else {
            [void](Clear-JobCrawlerDefaultProfile -ConfigDirectory $script:ConfigDirectory)
        }

        Set-LauncherCrawlerConfig -ProfileId $nextProfileId
        Set-LauncherOutputPaths
        Refresh-ProfileComboBox -SelectedProfileId $nextProfileId
        Refresh-ModeComboBox
        Refresh-DaysBackComboBox
        Refresh-SourceCheckboxes
        Refresh-CredentialList
        Add-LogLine ("Profile deleted: {0}" -f $deletedPath)
        Set-LauncherStatus $(if ([string]::IsNullOrWhiteSpace($nextProfileId)) { "Profile deleted. Create a new profile before crawling." } else { "Profile deleted." })
        Refresh-ReadinessChecklist
    }
    catch {
        if (Get-Command Show-LauncherException -ErrorAction SilentlyContinue) {
            Show-LauncherException -Context "Delete profile" -Exception $_.Exception
        }
        else {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Profile") | Out-Null
        }
    }
}

function Get-ProfileBuilderArray {
    param(
        [AllowNull()]$Builder,
        [string]$Name
    )

    return @(Get-ConfigStringArray (Get-ConfigProperty -Object $Builder -Name $Name -DefaultValue @()))
}

function Set-TextBoxLines {
    param(
        $TextBox,
        [AllowNull()][string[]]$Lines
    )

    $TextBox.Text = (@($Lines) -join [Environment]::NewLine)
}

function Get-TextBoxLines {
    param($TextBox)

    return @(ConvertTo-JobCrawlerProfileLineArray ([string]$TextBox.Text))
}

function Add-ProfileTableColumn {
    param(
        $Table,
        [float]$Width
    )

    $style = New-Object System.Windows.Forms.ColumnStyle
    $style.SizeType = [System.Windows.Forms.SizeType]::Percent
    $style.Width = $Width
    [void]$Table.ColumnStyles.Add($style)
}

function Add-ProfileTableRow {
    param(
        $Table,
        [float]$Height,
        [switch]$Absolute
    )

    $style = New-Object System.Windows.Forms.RowStyle
    if ($Absolute) {
        $style.SizeType = [System.Windows.Forms.SizeType]::Absolute
    }
    else {
        $style.SizeType = [System.Windows.Forms.SizeType]::Percent
    }
    $style.Height = $Height
    [void]$Table.RowStyles.Add($style)
}

function New-ProfileEditorTextBox {
    param([switch]$Multiline)

    $box = New-Object System.Windows.Forms.TextBox
    $box.Dock = [System.Windows.Forms.DockStyle]::Fill
    $box.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    if ($Multiline) {
        $box.Multiline = $true
        $box.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
        $box.AcceptsReturn = $true
        $box.WordWrap = $false
    }
    return $box
}

function New-ProfileEditorFieldPanel {
    param(
        [string]$Label,
        $Control
    )

    $panel = New-Object System.Windows.Forms.TableLayoutPanel
    $panel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panel.Padding = New-Object System.Windows.Forms.Padding(8)
    $panel.Margin = New-Object System.Windows.Forms.Padding(0)
    $panel.ColumnCount = 1
    $panel.RowCount = 2
    Add-ProfileTableColumn -Table $panel -Width 100
    Add-ProfileTableRow -Table $panel -Height 24 -Absolute
    Add-ProfileTableRow -Table $panel -Height 100

    $labelControl = New-Object System.Windows.Forms.Label
    $labelControl.Text = $Label
    $labelControl.Dock = [System.Windows.Forms.DockStyle]::Fill
    $labelControl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $labelControl.ForeColor = [System.Drawing.Color]::FromArgb(51, 65, 85)

    $Control.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panel.Controls.Add($labelControl, 0, 0)
    $panel.Controls.Add($Control, 0, 1)
    return $panel
}

function New-ProfileEditorTable {
    param(
        [int]$Columns,
        [int]$Rows
    )

    $table = New-Object System.Windows.Forms.TableLayoutPanel
    $table.Dock = [System.Windows.Forms.DockStyle]::Fill
    $table.Padding = New-Object System.Windows.Forms.Padding(10)
    $table.ColumnCount = $Columns
    $table.RowCount = $Rows

    for ($i = 0; $i -lt $Columns; $i++) {
        Add-ProfileTableColumn -Table $table -Width (100 / $Columns)
    }
    for ($i = 0; $i -lt $Rows; $i++) {
        Add-ProfileTableRow -Table $table -Height (100 / $Rows)
    }

    return $table
}

function Add-ProfileEditorField {
    param(
        $Table,
        [string]$Label,
        $Control,
        [int]$Column,
        [int]$Row,
        [int]$ColumnSpan = 1
    )

    $panel = New-ProfileEditorFieldPanel -Label $Label -Control $Control
    $Table.Controls.Add($panel, $Column, $Row)
    if ($ColumnSpan -gt 1) {
        $Table.SetColumnSpan($panel, $ColumnSpan)
    }
}

function Show-ProfileEditorDialog {
    param(
        [AllowNull()]$SourceProfileSummary = $null,
        [switch]$Duplicate
    )

    $sourceProfile = $null
    $builder = $null
    $sourceId = ""
    if ($null -ne $SourceProfileSummary) {
        $sourceId = [string]$SourceProfileSummary.Id
        $configRoot = (Resolve-Path -LiteralPath $script:ConfigDirectory).Path
        $sourceProfile = Read-JobCrawlerProfile -Root $configRoot -ProfileId $sourceId -AppliedOverrides $null
        $builder = Get-ConfigProperty -Object $sourceProfile -Name "profile_builder" -DefaultValue $null
    }

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = $(if ($null -eq $SourceProfileSummary) { "New profile" } elseif ($Duplicate) { "Duplicate profile" } else { "Edit profile" })
    $dialog.StartPosition = "CenterParent"
    $dialog.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $dialog.MinimumSize = New-Object System.Drawing.Size(760, 580)
    $dialog.ClientSize = New-Object System.Drawing.Size(900, 680)
    $dialog.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $dialog.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 251)

    $buttonBar = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttonBar.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $buttonBar.Height = 56
    $buttonBar.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
    $buttonBar.Padding = New-Object System.Windows.Forms.Padding(10)
    $buttonBar.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 251)
    $dialog.Controls.Add($buttonBar)

    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Dock = [System.Windows.Forms.DockStyle]::Fill
    $tabs.Padding = New-Object System.Drawing.Point(12, 5)
    $dialog.Controls.Add($tabs)

    $identityTab = New-Object System.Windows.Forms.TabPage
    $identityTab.Text = "Identity"
    $searchTab = New-Object System.Windows.Forms.TabPage
    $searchTab.Text = "Search intent"
    $qualityTab = New-Object System.Windows.Forms.TabPage
    $qualityTab.Text = "Quality"
    $filtersTab = New-Object System.Windows.Forms.TabPage
    $filtersTab.Text = "Filters"
    $preferencesTab = New-Object System.Windows.Forms.TabPage
    $preferencesTab.Text = "Preferences"
    [void]$tabs.TabPages.Add($identityTab)
    [void]$tabs.TabPages.Add($searchTab)
    [void]$tabs.TabPages.Add($qualityTab)
    [void]$tabs.TabPages.Add($filtersTab)
    [void]$tabs.TabPages.Add($preferencesTab)

    $identityTable = New-ProfileEditorTable -Columns 1 -Rows 2
    $identityTab.Controls.Add($identityTable)

    $nameBox = New-ProfileEditorTextBox
    $defaultName = [string](Get-ConfigProperty -Object $sourceProfile -Name "label" -DefaultValue "")
    if ($Duplicate -and -not [string]::IsNullOrWhiteSpace($defaultName)) {
        $defaultName = "Copy of $defaultName"
    }
    $nameBox.Text = $defaultName
    Add-ProfileEditorField -Table $identityTable -Label "Profile name" -Control $nameBox -Column 0 -Row 0

    $descriptionBox = New-ProfileEditorTextBox -Multiline
    $descriptionBox.Text = [string](Get-ConfigProperty -Object $sourceProfile -Name "description" -DefaultValue "")
    Add-ProfileEditorField -Table $identityTable -Label "Description" -Control $descriptionBox -Column 0 -Row 1

    $searchTable = New-ProfileEditorTable -Columns 2 -Rows 2
    $searchTab.Controls.Add($searchTable)

    $titlesBox = New-ProfileEditorTextBox -Multiline
    Set-TextBoxLines -TextBox $titlesBox -Lines (Get-ProfileBuilderArray -Builder $builder -Name "target_titles")
    Add-ProfileEditorField -Table $searchTable -Label "Target titles" -Control $titlesBox -Column 0 -Row 0

    $queriesBox = New-ProfileEditorTextBox -Multiline
    Set-TextBoxLines -TextBox $queriesBox -Lines (Get-ProfileBuilderArray -Builder $builder -Name "search_queries")
    Add-ProfileEditorField -Table $searchTable -Label "Search queries" -Control $queriesBox -Column 1 -Row 0

    $skillsBox = New-ProfileEditorTextBox -Multiline
    Set-TextBoxLines -TextBox $skillsBox -Lines (Get-ProfileBuilderArray -Builder $builder -Name "important_skills")
    Add-ProfileEditorField -Table $searchTable -Label "Important skills" -Control $skillsBox -Column 0 -Row 1 -ColumnSpan 2

    $filtersTable = New-ProfileEditorTable -Columns 2 -Rows 2
    $filtersTab.Controls.Add($filtersTable)

    $targetLocationsBox = New-ProfileEditorTextBox -Multiline
    Set-TextBoxLines -TextBox $targetLocationsBox -Lines (Get-ProfileBuilderArray -Builder $builder -Name "target_locations")
    Add-ProfileEditorField -Table $filtersTable -Label "Target locations" -Control $targetLocationsBox -Column 0 -Row 0

    $excludedLocationsBox = New-ProfileEditorTextBox -Multiline
    Set-TextBoxLines -TextBox $excludedLocationsBox -Lines (Get-ProfileBuilderArray -Builder $builder -Name "excluded_locations")
    Add-ProfileEditorField -Table $filtersTable -Label "Excluded locations" -Control $excludedLocationsBox -Column 1 -Row 0

    $exclusionsBox = New-ProfileEditorTextBox -Multiline
    Set-TextBoxLines -TextBox $exclusionsBox -Lines (Get-ProfileBuilderArray -Builder $builder -Name "exclusion_keywords")
    Add-ProfileEditorField -Table $filtersTable -Label "Excluded keywords" -Control $exclusionsBox -Column 0 -Row 1

    $existingContracts = @(Get-ProfileBuilderArray -Builder $builder -Name "excluded_contracts")
    if ($existingContracts.Count -eq 0) {
        $existingContracts = @("CDD", "Apprenticeship", "Internship", "Freelance")
    }
    $contractBoxes = @()
    $contractsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $contractsPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $contractsPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
    $contractsPanel.WrapContents = $false
    $contractOptions = @("CDD", "Apprenticeship", "Internship", "Freelance")
    foreach ($contractOption in $contractOptions) {
        $contractBox = New-Object System.Windows.Forms.CheckBox
        $contractBox.Text = $contractOption
        $contractBox.Checked = $contractOption -in $existingContracts
        $contractBox.AutoSize = $true
        $contractBox.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 6)
        [void]$contractsPanel.Controls.Add($contractBox)
        $contractBoxes += $contractBox
    }
    Add-ProfileEditorField -Table $filtersTable -Label "Excluded contracts" -Control $contractsPanel -Column 1 -Row 1

    $preferencesTable = New-ProfileEditorTable -Columns 1 -Rows 2
    $preferencesTab.Controls.Add($preferencesTable)

    $employerCombo = New-Object System.Windows.Forms.ComboBox
    $employerCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $employerCombo.DisplayMember = "Label"
    $employerCombo.Dock = [System.Windows.Forms.DockStyle]::Top
    $employerOptions = @(
        [PSCustomObject]@{ Label = "Neutral"; Value = "neutral" },
        [PSCustomObject]@{ Label = "Prefer annonceur"; Value = "annonceur" },
        [PSCustomObject]@{ Label = "Agency / consulting OK"; Value = "agency_consulting_ok" }
    )
    foreach ($option in $employerOptions) {
        [void]$employerCombo.Items.Add($option)
    }
    $currentPreference = [string](Get-ConfigProperty -Object $builder -Name "employer_preference" -DefaultValue "neutral")
    $employerCombo.SelectedIndex = 0
    for ($i = 0; $i -lt $employerCombo.Items.Count; $i++) {
        if ([string]$employerCombo.Items[$i].Value -eq $currentPreference) {
            $employerCombo.SelectedIndex = $i
            break
        }
    }
    Add-ProfileEditorField -Table $preferencesTable -Label "Employer preference" -Control $employerCombo -Column 0 -Row 0

    $defaultCheckBox = New-Object System.Windows.Forms.CheckBox
    $defaultCheckBox.Text = "Use as default"
    $defaultCheckBox.Dock = [System.Windows.Forms.DockStyle]::Top
    $defaultCheckBox.AutoSize = $true
    $defaultCheckBox.Checked = ($null -ne $SourceProfileSummary -and [string]$SourceProfileSummary.Id -eq [string]$script:CrawlerConfig.Profile.Id -and -not $Duplicate)
    Add-ProfileEditorField -Table $preferencesTable -Label "Default profile" -Control $defaultCheckBox -Column 0 -Row 1

    $qualityTable = New-Object System.Windows.Forms.TableLayoutPanel
    $qualityTable.Dock = [System.Windows.Forms.DockStyle]::Fill
    $qualityTable.Padding = New-Object System.Windows.Forms.Padding(14)
    $qualityTable.ColumnCount = 1
    $qualityTable.RowCount = 4
    Add-ProfileTableColumn -Table $qualityTable -Width 100
    Add-ProfileTableRow -Table $qualityTable -Height 48 -Absolute
    Add-ProfileTableRow -Table $qualityTable -Height 44 -Absolute
    Add-ProfileTableRow -Table $qualityTable -Height 100
    Add-ProfileTableRow -Table $qualityTable -Height 46 -Absolute
    $qualityTab.Controls.Add($qualityTable)

    $qualityScoreLabel = New-Object System.Windows.Forms.Label
    $qualityScoreLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $qualityScoreLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 14)
    $qualityScoreLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $qualityTable.Controls.Add($qualityScoreLabel, 0, 0)

    $qualityHintLabel = New-Object System.Windows.Forms.Label
    $qualityHintLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $qualityHintLabel.ForeColor = [System.Drawing.Color]::FromArgb(77, 91, 114)
    $qualityHintLabel.Text = "Quality checks whether the profile has enough search breadth before filtering. Search queries bring jobs in; titles, skills, exclusions, locations, and contracts decide what stays."
    $qualityHintLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $qualityTable.Controls.Add($qualityHintLabel, 0, 1)

    $qualityDetailsBox = New-ProfileEditorTextBox -Multiline
    $qualityDetailsBox.ReadOnly = $true
    $qualityDetailsBox.WordWrap = $true
    $qualityTable.Controls.Add($qualityDetailsBox, 0, 2)

    $qualityButtons = New-Object System.Windows.Forms.FlowLayoutPanel
    $qualityButtons.Dock = [System.Windows.Forms.DockStyle]::Fill
    $qualityButtons.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    $qualityButtons.WrapContents = $false
    $qualityButtons.Padding = New-Object System.Windows.Forms.Padding(0, 6, 0, 0)
    $qualityTable.Controls.Add($qualityButtons, 0, 3)

    $improveQueriesButton = New-Object System.Windows.Forms.Button
    $improveQueriesButton.Text = "Improve queries"
    $improveQueriesButton.Size = New-Object System.Drawing.Size(116, 30)
    [void]$qualityButtons.Controls.Add($improveQueriesButton)

    $refreshQualityButton = New-Object System.Windows.Forms.Button
    $refreshQualityButton.Text = "Refresh score"
    $refreshQualityButton.Size = New-Object System.Drawing.Size(112, 30)
    [void]$qualityButtons.Controls.Add($refreshQualityButton)

    $getExcludedContractsFromEditor = {
        return @($contractBoxes | Where-Object { $_.Checked } | ForEach-Object { [string]$_.Text })
    }
    $getQualityFromEditor = {
        return Get-JobCrawlerProfileQuality `
            -Label $nameBox.Text `
            -TargetTitles (Get-TextBoxLines $titlesBox) `
            -SearchQueries (Get-TextBoxLines $queriesBox) `
            -ImportantSkills (Get-TextBoxLines $skillsBox) `
            -ExclusionKeywords (Get-TextBoxLines $exclusionsBox) `
            -TargetLocations (Get-TextBoxLines $targetLocationsBox) `
            -ExcludedContracts (& $getExcludedContractsFromEditor)
    }
    $updateQualityPreview = {
        try {
            $quality = & $getQualityFromEditor
            $qualityScoreLabel.Text = ("Profile quality: {0}/100 - {1}" -f $quality.Score, $quality.Level)
            if ($quality.Score -ge 85) {
                $qualityScoreLabel.ForeColor = [System.Drawing.Color]::FromArgb(30, 104, 72)
            }
            elseif ($quality.Score -ge 70) {
                $qualityScoreLabel.ForeColor = [System.Drawing.Color]::FromArgb(40, 88, 141)
            }
            elseif ($quality.Score -ge 50) {
                $qualityScoreLabel.ForeColor = [System.Drawing.Color]::FromArgb(132, 86, 18)
            }
            else {
                $qualityScoreLabel.ForeColor = [System.Drawing.Color]::FromArgb(154, 69, 59)
            }

            $detailLines = New-Object System.Collections.Generic.List[string]
            foreach ($finding in @($quality.Findings)) {
                $detailLines.Add(("- {0}" -f $finding)) | Out-Null
            }
            $suggestions = @($quality.QuerySuggestions | Select-Object -First 12)
            if ($suggestions.Count -gt 0) {
                $detailLines.Add("") | Out-Null
                $detailLines.Add("Suggested search queries:") | Out-Null
                foreach ($suggestion in $suggestions) {
                    $detailLines.Add(("- {0}" -f $suggestion)) | Out-Null
                }
            }
            $qualityDetailsBox.Text = ($detailLines.ToArray() -join [Environment]::NewLine)
        }
        catch {
            $qualityScoreLabel.Text = "Profile quality: unavailable"
            $qualityScoreLabel.ForeColor = [System.Drawing.Color]::FromArgb(154, 69, 59)
            $qualityDetailsBox.Text = $_.Exception.Message
        }
    }
    $mergeQuerySuggestions = {
        $quality = & $getQualityFromEditor
        $mergedQueries = New-Object System.Collections.Generic.List[string]
        foreach ($query in @((Get-TextBoxLines $queriesBox) + @($quality.QuerySuggestions))) {
            Add-JobCrawlerProfileUniqueString -List $mergedQueries -Value $query
        }
        Set-TextBoxLines -TextBox $queriesBox -Lines @($mergedQueries.ToArray())
        & $updateQualityPreview
    }
    $improveQueriesButton.Add_Click({ & $mergeQuerySuggestions })
    $refreshQualityButton.Add_Click({ & $updateQualityPreview })
    foreach ($profileEditorControl in @($nameBox, $titlesBox, $queriesBox, $skillsBox, $exclusionsBox, $targetLocationsBox)) {
        $profileEditorControl.Add_TextChanged({ & $updateQualityPreview })
    }
    foreach ($contractBox in @($contractBoxes)) {
        $contractBox.Add_CheckedChanged({ & $updateQualityPreview })
    }
    $qualityTab.Add_Enter({ & $updateQualityPreview })

    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = "Save"
    $saveButton.Size = New-Object System.Drawing.Size(88, 30)
    $saveButton.Add_Click({
        try {
            if (@(Get-TextBoxLines $queriesBox).Count -lt 4) {
                & $mergeQuerySuggestions
            }
            $quality = & $getQualityFromEditor
            if ($quality.Score -lt 45) {
                $answer = [System.Windows.Forms.MessageBox]::Show(
                    ("This profile quality is {0}/100 ({1}). It may return 0 results or noisy results.`n`nSave anyway?" -f $quality.Score, $quality.Level),
                    "Profile quality",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
                if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
                    return
                }
            }

            $contracts = @($contractBoxes | Where-Object { $_.Checked } | ForEach-Object { [string]$_.Text })
            $profile = New-JobCrawlerProfileFromBuilder `
                -Label $nameBox.Text `
                -Id $(if ($null -ne $SourceProfileSummary -and -not $Duplicate) { $sourceId } else { "" }) `
                -Description $descriptionBox.Text `
                -TargetTitles (Get-TextBoxLines $titlesBox) `
                -ImportantSkills (Get-TextBoxLines $skillsBox) `
                -ExclusionKeywords (Get-TextBoxLines $exclusionsBox) `
                -SearchQueries (Get-TextBoxLines $queriesBox) `
                -TargetLocations (Get-TextBoxLines $targetLocationsBox) `
                -ExcludedLocations (Get-TextBoxLines $excludedLocationsBox) `
                -ExcludedContracts $contracts `
                -EmployerPreference ([string]$employerCombo.SelectedItem.Value) `
                -Compact

            $path = Save-JobCrawlerLocalProfile -ConfigDirectory $script:ConfigDirectory -Profile $profile
            $profileId = [string](Get-ConfigProperty -Object $profile -Name "id" -DefaultValue "")
            $profileLabel = [string](Get-ConfigProperty -Object $profile -Name "label" -DefaultValue $profileId)
            if ($defaultCheckBox.Checked) {
                [void](Set-JobCrawlerDefaultProfile -ConfigDirectory $script:ConfigDirectory -ProfileId $profileId)
            }
            Set-LauncherCrawlerConfig -ProfileId $profileId
            Set-LauncherOutputPaths
            Refresh-ProfileComboBox -SelectedProfileId $profileId
            Refresh-ModeComboBox
            Refresh-DaysBackComboBox
            Refresh-SourceCheckboxes
            Refresh-CredentialList
            Add-LogLine ("Profile saved: {0}" -f $path)
            Set-LauncherStatus ("Profile saved: {0}" -f $profileLabel)
            $dialog.Close()
        }
        catch {
            if (Get-Command Show-LauncherException -ErrorAction SilentlyContinue) {
                Show-LauncherException -Context "Save profile" -Exception $_.Exception
            }
            else {
                [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Profile") | Out-Null
            }
        }
    })

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Size = New-Object System.Drawing.Size(88, 30)
    $cancelButton.Add_Click({ $dialog.Close() })
    [void]$buttonBar.Controls.Add($cancelButton)
    [void]$buttonBar.Controls.Add($saveButton)

    $dialog.AcceptButton = $saveButton
    $dialog.CancelButton = $cancelButton
    & $updateQualityPreview
    [void]$dialog.ShowDialog($script:MainForm)
}

function Show-NewProfileDialog {
    Show-ProfileEditorDialog
}

function Show-EditProfileDialog {
    if ($null -eq $script:ProfileComboBox -or $null -eq $script:ProfileComboBox.SelectedItem) {
        return
    }
    Show-ProfileEditorDialog -SourceProfileSummary $script:ProfileComboBox.SelectedItem
}

function Show-DuplicateProfileDialog {
    if ($null -eq $script:ProfileComboBox -or $null -eq $script:ProfileComboBox.SelectedItem) {
        return
    }
    Show-ProfileEditorDialog -SourceProfileSummary $script:ProfileComboBox.SelectedItem -Duplicate
}
