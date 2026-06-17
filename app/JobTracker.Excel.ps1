# Auto-extracted from Find-AnalyticsJobs.ps1. Keep dot-sourced execution order in the main script.

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

function New-OrderedJobRecord {
    param([hashtable]$Values)

    $ordered = [ordered]@{}
    foreach ($column in $MasterColumns) {
        if ($Values.ContainsKey($column) -and $null -ne $Values[$column]) {
            $ordered[$column] = Repair-DisplayText ([string]$Values[$column])
        }
        else {
            $ordered[$column] = ""
        }
    }

    return [PSCustomObject]$ordered
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
                $value = Repair-DisplayText ([string]$sheet.Cells.Item($row, $column).Text)
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

function Get-WorkbookSheetName {
    param(
        [string]$Key,
        [string]$DefaultValue
    )

    return [string](Get-ConfigPathValue -Object $script:JobCrawlerWorkbookConfig -Path ("sheets.{0}" -f $Key) -DefaultValue $DefaultValue)
}

function Get-FeedbackQualityRows {
    param([object[]]$Rows)

    $findings = New-Object System.Collections.Generic.List[object]
    $knownReasons = @(Get-JobTrackerIgnoreReasonKeys)

    foreach ($row in @($Rows)) {
        $status = ConvertTo-MatchText (Get-RowValue -Row $row -Name "status")
        $notes = Get-RowValue -Row $row -Name "notes"
        $reason = Get-IgnoreReasonFromNotes $notes
        $title = Get-RowValue -Row $row -Name "job_title"
        $company = Get-RowValue -Row $row -Name "company_name"
        $jobId = Get-RowValue -Row $row -Name "job_id"

        if ($status -eq "ignored") {
            if ([string]::IsNullOrWhiteSpace($notes)) {
                $findings.Add([PSCustomObject]@{ Severity = "High"; Status = $status; Job = $title; Company = $company; Issue = "Ignored row has no Apply notes"; Action = "Add ignore_reason=...; detail=..."; JobID = $jobId }) | Out-Null
            }
            elseif ([string]::IsNullOrWhiteSpace($reason) -or $reason -notin $knownReasons) {
                $findings.Add([PSCustomObject]@{ Severity = "Medium"; Status = $status; Job = $title; Company = $company; Issue = "Ignored row has an unrecognized ignore reason"; Action = "Use one of the ignore_reason templates"; JobID = $jobId }) | Out-Null
            }
        }
        elseif ($status -in @("interesting", "applied", "interview", "offer")) {
            if ([string]::IsNullOrWhiteSpace($notes)) {
                $findings.Add([PSCustomObject]@{ Severity = "Low"; Status = $status; Job = $title; Company = $company; Issue = "Positive/application row has no Apply notes"; Action = "Add why it is interesting or what happened"; JobID = $jobId }) | Out-Null
            }
        }

        if ($status -in @("applied", "interview", "offer", "rejected", "withdrawn") -and [string]::IsNullOrWhiteSpace((Get-RowValue -Row $row -Name "applied_date"))) {
            $findings.Add([PSCustomObject]@{ Severity = "Medium"; Status = $status; Job = $title; Company = $company; Issue = "Application-related row has no Applied date"; Action = "Fill Applied date"; JobID = $jobId }) | Out-Null
        }
    }

    if ($findings.Count -eq 0) {
        $findings.Add([PSCustomObject]@{ Severity = "OK"; Status = ""; Job = ""; Company = ""; Issue = "No feedback quality issue found"; Action = ""; JobID = "" }) | Out-Null
    }

    return @($findings.ToArray())
}

function Write-KeyValueSheet {
    param(
        $Sheet,
        [string]$Title,
        [object[]]$Pairs
    )

    $Sheet.Cells.Item(1, 1).Value2 = $Title
    $Sheet.Cells.Item(1, 1).Font.Bold = $true
    $Sheet.Cells.Item(1, 1).Font.Size = 16
    $Sheet.Cells.Font.Name = "Segoe UI"
    $Sheet.Cells.Font.Size = 10
    $Sheet.Cells.Font.Color = Get-ExcelColor 40 47 52

    $rowNumber = 3
    foreach ($pair in @($Pairs)) {
        $Sheet.Cells.Item($rowNumber, 1).Value2 = [string]$pair[0]
        $Sheet.Cells.Item($rowNumber, 2).Value2 = [string]$pair[1]
        $rowNumber++
    }

    $Sheet.Columns.Item(1).ColumnWidth = 34
    $Sheet.Columns.Item(2).ColumnWidth = 90
    $Sheet.Columns.Item(2).WrapText = $true
    if ($rowNumber -gt 3) {
        $Sheet.Range("A3:A$($rowNumber - 1)").Font.Bold = $true
        $Sheet.Range("A3:B$($rowNumber - 1)").Borders.LineStyle = 1
        $Sheet.Range("A3:B$($rowNumber - 1)").Borders.Color = Get-ExcelColor 226 232 240
    }
}

function Write-ObjectListSheet {
    param(
        $Sheet,
        [string]$Title,
        [object[]]$Rows
    )

    $Sheet.Cells.Item(1, 1).Value2 = $Title
    $Sheet.Cells.Item(1, 1).Font.Bold = $true
    $Sheet.Cells.Item(1, 1).Font.Size = 16
    $Sheet.Cells.Font.Name = "Segoe UI"
    $Sheet.Cells.Font.Size = 10
    $Sheet.Cells.Font.Color = Get-ExcelColor 40 47 52

    $rowsArray = @($Rows)
    if ($rowsArray.Count -eq 0) {
        $Sheet.Cells.Item(3, 1).Value2 = "No data"
        return
    }

    $properties = @($rowsArray[0].PSObject.Properties.Name)
    for ($column = 0; $column -lt $properties.Count; $column++) {
        $cell = $Sheet.Cells.Item(3, $column + 1)
        $cell.Value2 = [string]$properties[$column]
        $cell.Font.Bold = $true
        $cell.Font.Color = Get-ExcelColor 255 255 255
        $cell.Interior.Color = Get-ExcelColor 38 50 56
    }

    for ($rowIndex = 0; $rowIndex -lt $rowsArray.Count; $rowIndex++) {
        $excelRow = $rowIndex + 4
        for ($column = 0; $column -lt $properties.Count; $column++) {
            $value = $rowsArray[$rowIndex].PSObject.Properties[$properties[$column]].Value
            $Sheet.Cells.Item($excelRow, $column + 1).Value2 = Repair-DisplayText ([string]$value)
        }
    }

    $Sheet.Columns.AutoFit() | Out-Null
    $Sheet.Rows.Item(3).AutoFilter() | Out-Null
    $Sheet.Range($Sheet.Cells.Item(3, 1), $Sheet.Cells.Item($rowsArray.Count + 3, $properties.Count)).Borders.LineStyle = 1
    $Sheet.Range($Sheet.Cells.Item(3, 1), $Sheet.Cells.Item($rowsArray.Count + 3, $properties.Count)).Borders.Color = Get-ExcelColor 226 232 240
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
    $settingsSheet = $null
    $sourceHealthSheet = $null
    $feedbackQualitySheet = $null
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

        $jobsSheetName = Get-WorkbookSheetName -Key "jobs" -DefaultValue "Jobs"
        $summarySheetName = Get-WorkbookSheetName -Key "summary" -DefaultValue "Summary"
        $settingsSheetName = Get-WorkbookSheetName -Key "settings" -DefaultValue "Settings"
        $sourceHealthSheetName = Get-WorkbookSheetName -Key "source_health" -DefaultValue "Source Health"
        $feedbackQualitySheetName = Get-WorkbookSheetName -Key "feedback_quality" -DefaultValue "Feedback Quality"

        $jobsSheet = $workbook.Worksheets.Item(1)
        $jobsSheet.Name = $jobsSheetName
        $summarySheet = $workbook.Worksheets.Add([System.Type]::Missing, $jobsSheet)
        $summarySheet.Name = $summarySheetName
        $settingsSheet = $workbook.Worksheets.Add([System.Type]::Missing, $summarySheet)
        $settingsSheet.Name = $settingsSheetName
        $sourceHealthSheet = $workbook.Worksheets.Add([System.Type]::Missing, $settingsSheet)
        $sourceHealthSheet.Name = $sourceHealthSheetName
        $feedbackQualitySheet = $workbook.Worksheets.Add([System.Type]::Missing, $sourceHealthSheet)
        $feedbackQualitySheet.Name = $feedbackQualitySheetName

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
                $value = Repair-DisplayText (Get-RowValue -Row $row -Name $columnName)

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
                elseif ($columnName -in @("match_score", "role_score", "employer_fit", "location_fit", "seniority_fit", "contract_fit", "days_since_published", "days_since_first_seen", "days_since_last_seen", "feedback_adjustment")) {
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

        foreach ($columnName in @("duplicate_reason", "job_title", "matched_keywords", "fit_notes", "job_url_raw", "notes")) {
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
        foreach ($columnName in @("review_priority", "status", "employer_type", "contract_type", "platform", "source_count", "published_date", "days_since_published", "job_url", "applied_date", "match_level", "match_score", "role_score", "employer_fit", "location_fit", "seniority_fit", "contract_fit", "seen_in_current_crawl", "first_seen_date", "last_seen_date", "is_new")) {
            if ($columnIndex.ContainsKey($columnName)) {
                $jobsSheet.Columns.Item([int]$columnIndex[$columnName]).HorizontalAlignment = -4108
            }
        }
        foreach ($columnName in @("published_date", "applied_date", "first_seen_date", "last_seen_date")) {
            if ($columnIndex.ContainsKey($columnName)) {
                $jobsSheet.Columns.Item([int]$columnIndex[$columnName]).NumberFormat = "yyyy-mm-dd"
            }
        }
        foreach ($columnName in @("source_count", "days_since_published", "match_score", "role_score", "employer_fit", "location_fit", "seniority_fit", "contract_fit", "days_since_first_seen", "days_since_last_seen", "feedback_adjustment")) {
            if ($columnIndex.ContainsKey($columnName)) {
                $jobsSheet.Columns.Item([int]$columnIndex[$columnName]).NumberFormat = "0"
            }
        }

        Set-JobTrackerDataValidation -Workbook $workbook -Excel $excel -Sheet $jobsSheet -ColumnIndex $columnIndex -LastDataRow $lastDataRow
        Set-ReviewPriorityFormulas -Sheet $jobsSheet -ColumnIndex $columnIndex -LastDataRow $lastDataRow

        if ($rowCount -gt 0) {
            $dataRange = $jobsSheet.Range($jobsSheet.Cells.Item(2, 1), $jobsSheet.Cells.Item($lastDataRow, $lastColumn))
            $dataRange.VerticalAlignment = -4160
            $dataRange.RowHeight = 30
            $dataRange.Interior.Color = Get-ExcelColor 255 255 255
            Set-StatusRowConditionalFormatting -Range $dataRange -ColumnIndex $columnIndex
            Set-StatusCellConditionalFormatting -Sheet $jobsSheet -ColumnIndex $columnIndex -LastDataRow $lastDataRow
            Set-ReviewPriorityConditionalFormatting -Sheet $jobsSheet -ColumnIndex $columnIndex -LastDataRow $lastDataRow
            Set-IgnoredNotesReminderFormatting -Sheet $jobsSheet -ColumnIndex $columnIndex -LastDataRow $lastDataRow
        }

        $colors = Get-JobTrackerWorkbookColors -DarkTextColor $darkTextColor
        for ($excelRow = 2; $excelRow -le ($rowCount + 1); $excelRow++) {
            $matchLevel = Get-RowValue -Row $Rows[$excelRow - 2] -Name "match_level"
            $seen = Get-RowValue -Row $Rows[$excelRow - 2] -Name "seen_in_current_crawl"

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
        $employerTypeSummary = @(
            "annonceur {0}" -f @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "employer_type") -eq "annonceur" }).Count
            "agency {0}" -f @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "employer_type") -eq "agency" }).Count
            "consulting {0}" -f @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "employer_type") -eq "consulting" }).Count
            "ESN {0}" -f @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "employer_type") -eq "esn" }).Count
            "unknown {0}" -f @($Rows | Where-Object { [string]::IsNullOrWhiteSpace((Get-RowValue -Row $_ -Name "employer_type")) -or (Get-RowValue -Row $_ -Name "employer_type") -eq "unknown" }).Count
        ) -join " | "
        $fitDemotionSummary = @(
            "employer {0}" -f @($Rows | Where-Object { (Get-IntegerRowValue -Row $_ -Name "employer_fit") -lt 0 }).Count
            "location {0}" -f @($Rows | Where-Object { (Get-IntegerRowValue -Row $_ -Name "location_fit") -lt 0 }).Count
            "seniority {0}" -f @($Rows | Where-Object { (Get-IntegerRowValue -Row $_ -Name "seniority_fit") -lt 0 }).Count
            "contract {0}" -f @($Rows | Where-Object { (Get-IntegerRowValue -Row $_ -Name "contract_fit") -lt 0 }).Count
        ) -join " | "
        $sourceSummary = @(
            "France Travail {0}" -f @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "platform") -match "France Travail" }).Count
            "Adzuna {0}" -f @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "platform") -match "Adzuna" }).Count
            "APEC {0}" -f @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "platform") -match "APEC" }).Count
            "HelloWork {0}" -f @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "platform") -match "HelloWork" }).Count
            "WTTJ {0}" -f @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "platform") -match "Welcome to the Jungle" }).Count
            "LinkedIn {0}" -f @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "platform") -match "LinkedIn" }).Count
        ) -join " | "
        $summaryPairs = @(
            @("Generated", $RunStamp),
            @("Crawl mode", $CrawlMode),
            @("Retention rule", "Keep non-application jobs only when Published is on or after $CutoffDate."),
            @("Rows in workbook", [string](@($Rows).Count)),
            @("Seen in this crawl", [string]$currentVisibleCount),
            @("New this run", [string]$newVisibleCount),
            @("Application rows kept", [string]$applicationVisibleCount),
            @("Match levels", ("High {0} | Medium {1} | Review {2}" -f $highVisibleCount, $mediumVisibleCount, $reviewVisibleCount)),
            @("Sources", $sourceSummary),
            @("Employer types", $employerTypeSummary),
            @("Fit demotions", $fitDemotionSummary),
            @("Total matched before contract filter", (Get-SummaryValue -Summary $Summary -Name "TotalMatched")),
            @("Excluded CDD/apprenticeship/internship/freelance", (Get-SummaryValue -Summary $Summary -Name "ExcludedContractCount")),
            @("Duplicates merged this run", (Get-SummaryValue -Summary $Summary -Name "DuplicateCount")),
            @("Rows removed by retention", (Get-SummaryValue -Summary $Summary -Name "RemovedCount")),
            @("Source diagnostics", (Get-SummaryValue -Summary $Summary -Name "SourceDiagnostics")),
            @("Diagnostic file", (Get-SummaryValue -Summary $Summary -Name "DiagnosticPath")),
            @("Cache pruning", ("{0} file(s), {1} MB removed | {2} MB remaining" -f (Get-SummaryValue -Summary $Summary -Name "CachePrunedFiles"), (Get-SummaryValue -Summary $Summary -Name "CachePrunedMB"), (Get-SummaryValue -Summary $Summary -Name "CacheRemainingMB"))),
            @("Run history", (Get-SummaryValue -Summary $Summary -Name "RunHistoryPath")),
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

        $configRoot = ""
        if (Get-Variable -Name JobCrawlerConfig -Scope Script -ErrorAction SilentlyContinue) {
            $configRoot = [string]$script:JobCrawlerConfig.Root
        }
        $credentialStatuses = @()
        if (Get-Variable -Name JobCrawlerSourcesConfig -Scope Script -ErrorAction SilentlyContinue) {
            $credentialStatuses = @(Get-JobCrawlerCredentialStatuses -SourcesConfig $script:JobCrawlerSourcesConfig)
        }
        $credentialSummary = $(if ($credentialStatuses.Count -gt 0) { ($credentialStatuses | ForEach-Object { "{0}/{1}: {2}" -f $_.Source, $_.Credential, $_.Status }) -join " | " } else { "No credential config" })
        $localOverrideSummary = "none"
        if (Get-Variable -Name JobCrawlerConfig -Scope Script -ErrorAction SilentlyContinue) {
            $localOverrides = @(Get-ConfigProperty -Object $script:JobCrawlerConfig -Name "LocalOverrides" -DefaultValue @())
            if ($localOverrides.Count -gt 0) {
                $localOverrideSummary = ($localOverrides | ForEach-Object { Split-Path -Leaf $_ }) -join " | "
            }
        }
        $sourceDefaultSummary = "No source metadata"
        if (Get-Variable -Name JobCrawlerSourcesConfig -Scope Script -ErrorAction SilentlyContinue) {
            $sourceDefinitions = @(Get-JobCrawlerSourceDefinitions -SourcesConfig $script:JobCrawlerSourcesConfig)
            if ($sourceDefinitions.Count -gt 0) {
                $sourceDefaultSummary = ($sourceDefinitions | ForEach-Object {
                    "{0}: {1}{2}" -f $_.Key,
                        $(if ($_.EnabledByDefault) { "on" } else { "off" }),
                        $(if ($_.RequiresCredential) { " (credentials)" } else { "" })
                }) -join " | "
            }
        }
        $settingsPairs = @(
            @("Crawl mode", $CrawlMode),
            @("Days back", [string]$DaysBack),
            @("Location", $Location),
            @("Tracker", $fullPath),
            @("Config directory", $configRoot),
            @("Local config overrides", $localOverrideSummary),
            @("Source defaults", $sourceDefaultSummary),
            @("Cache directory", $CacheDirectory),
            @("Cache TTL hours", [string]$CacheTtlHours),
            @("Dry run", (Get-SummaryValue -Summary $Summary -Name "DryRun")),
            @("Diagnostics mode", (Get-SummaryValue -Summary $Summary -Name "DiagnosticMode")),
            @("Diagnostic file", (Get-SummaryValue -Summary $Summary -Name "DiagnosticPath")),
            @("Crawl caps", (Get-SummaryValue -Summary $Summary -Name "CrawlCaps")),
            @("Run history", (Get-SummaryValue -Summary $Summary -Name "RunHistoryPath")),
            @("Credential status", $credentialSummary),
            @("LinkedIn queries", [string](@($LinkedInQueries).Count)),
            @("API queries", [string](@($ApiSearchQueries).Count)),
            @("Matching threshold", [string]$MinimumMatchScore)
        )
        Write-KeyValueSheet -Sheet $settingsSheet -Title "Settings" -Pairs $settingsPairs

        $sourceRows = @()
        if (Get-Variable -Name SourceRunStats -Scope Script -ErrorAction SilentlyContinue) {
            $sourceRows = @($script:SourceRunStats.ToArray())
        }
        if ($sourceRows.Count -eq 0) {
            $sourceRows = @([PSCustomObject]@{ Source = "No crawl source"; DurationSeconds = ""; SearchRequests = ""; DetailRequests = ""; CacheHits = ""; Candidates = ""; SelectedDetails = ""; SkippedOld = ""; SkippedContract = ""; SkippedNoMatch = ""; SkippedByCap = ""; Errors = ""; Matches = ""; Notes = "" })
        }
        Write-ObjectListSheet -Sheet $sourceHealthSheet -Title "Source Health" -Rows $sourceRows

        $feedbackRows = @(Get-FeedbackQualityRows -Rows $Rows)
        Write-ObjectListSheet -Sheet $feedbackQualitySheet -Title "Feedback Quality" -Rows $feedbackRows

        $workbook.Worksheets.Item($jobsSheetName).Activate() | Out-Null
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
        Release-ComObject $feedbackQualitySheet
        Release-ComObject $sourceHealthSheet
        Release-ComObject $settingsSheet
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

function Backup-TrackerFile {
    param([string]$Path)

    return Backup-JobTrackerFile -Path $Path -MaxBackups $MaxBackups
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
    $preferredUrlRow = Select-PreferredUrlRow @($CurrentRow, $ExistingRow)
    $primaryUrl = Get-RowValue -Row $preferredUrlRow -Name "job_url_raw"
    if ([string]::IsNullOrWhiteSpace($primaryUrl)) {
        $primaryUrl = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "job_url_raw") -Fallback (Get-RowValue -Row $ExistingRow -Name "job_url_raw")
    }
    $allUrls = @(Get-RowUrlValues @($CurrentRow, $ExistingRow))
    $alternateUrls = @($allUrls | Where-Object { $_ -ne $primaryUrl })
    $jobTitleValue = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "job_title") -Fallback (Get-RowValue -Row $ExistingRow -Name "job_title")
    $companyNameValue = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "company_name") -Fallback (Get-RowValue -Row $ExistingRow -Name "company_name")
    $locationValue = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "location") -Fallback (Get-RowValue -Row $ExistingRow -Name "location")
    $contractTypeValue = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "contract_type") -Fallback (Get-RowValue -Row $ExistingRow -Name "contract_type")
    $matchScoreValue = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "match_score") -Fallback (Get-RowValue -Row $ExistingRow -Name "match_score")
    $matchedKeywordsValue = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "matched_keywords") -Fallback (Get-RowValue -Row $ExistingRow -Name "matched_keywords")
    $employerTypeValue = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "employer_type") -Fallback (Get-RowValue -Row $ExistingRow -Name "employer_type")
    $roleScoreValue = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "role_score") -Fallback (Get-RowValue -Row $ExistingRow -Name "role_score")
    $employerFitValue = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "employer_fit") -Fallback (Get-RowValue -Row $ExistingRow -Name "employer_fit")
    $locationFitValue = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "location_fit") -Fallback (Get-RowValue -Row $ExistingRow -Name "location_fit")
    $seniorityFitValue = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "seniority_fit") -Fallback (Get-RowValue -Row $ExistingRow -Name "seniority_fit")
    $contractFitValue = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "contract_fit") -Fallback (Get-RowValue -Row $ExistingRow -Name "contract_fit")
    $fitNotesValue = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "fit_notes") -Fallback (Get-RowValue -Row $ExistingRow -Name "fit_notes")

    if ([string]::IsNullOrWhiteSpace($employerTypeValue) -or
        [string]::IsNullOrWhiteSpace($roleScoreValue) -or
        [string]::IsNullOrWhiteSpace($employerFitValue) -or
        [string]::IsNullOrWhiteSpace($locationFitValue) -or
        [string]::IsNullOrWhiteSpace($seniorityFitValue) -or
        [string]::IsNullOrWhiteSpace($contractFitValue) -or
        [string]::IsNullOrWhiteSpace($fitNotesValue)) {
        $roleBaseScore = 0
        if (-not [int]::TryParse($roleScoreValue, [ref]$roleBaseScore)) {
            [void][int]::TryParse($matchScoreValue, [ref]$roleBaseScore)
        }
        $fit = Get-JobFitDimensions `
            -RoleScore $roleBaseScore `
            -Title $jobTitleValue `
            -CompanyName $companyNameValue `
            -JobLocation $locationValue `
            -ContractType $contractTypeValue `
            -Text (Join-CleanTextParts @($matchedKeywordsValue, (Get-RowValue -Row $ExistingRow -Name "notes")))

        if ([string]::IsNullOrWhiteSpace($employerTypeValue)) { $employerTypeValue = [string]$fit.EmployerType }
        if ([string]::IsNullOrWhiteSpace($roleScoreValue)) { $roleScoreValue = [string]$fit.RoleScore }
        if ([string]::IsNullOrWhiteSpace($employerFitValue)) { $employerFitValue = [string]$fit.EmployerFit }
        if ([string]::IsNullOrWhiteSpace($locationFitValue)) { $locationFitValue = [string]$fit.LocationFit }
        if ([string]::IsNullOrWhiteSpace($seniorityFitValue)) { $seniorityFitValue = [string]$fit.SeniorityFit }
        if ([string]::IsNullOrWhiteSpace($contractFitValue)) { $contractFitValue = [string]$fit.ContractFit }
        if ([string]::IsNullOrWhiteSpace($fitNotesValue)) { $fitNotesValue = [string]$fit.Notes }
    }

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
        job_title             = $jobTitleValue
        company_name          = $companyNameValue
        employer_type         = $employerTypeValue
        location              = $locationValue
        contract_type         = $contractTypeValue
        match_score           = $matchScoreValue
        match_level           = $matchLevel
        matched_keywords      = $matchedKeywordsValue
        role_score            = $roleScoreValue
        employer_fit          = $employerFitValue
        location_fit          = $locationFitValue
        seniority_fit         = $seniorityFitValue
        contract_fit          = $contractFitValue
        fit_notes             = $fitNotesValue
        job_url               = ConvertTo-ExcelHyperlinkFormula -Url $primaryUrl -Label "Open"
        job_url_raw           = $primaryUrl
        alternate_urls        = ($alternateUrls -join "; ")
        platform              = Join-UniqueTextValues -Values @((Get-RowValue -Row $CurrentRow -Name "platform"), (Get-RowValue -Row $ExistingRow -Name "platform"))
        source_count          = [string](Get-SourceCountFromRows @($CurrentRow, $ExistingRow))
        published_date        = $publishedDate
        days_since_published  = Get-DaysSince $publishedDate
        notes                 = Get-RowValue -Row $ExistingRow -Name "notes"
    }
}

function ConvertTo-TrackerRecordFromExisting {
    param($ExistingRow)

    return ConvertTo-TrackerRecord -CurrentRow $ExistingRow -ExistingRow $ExistingRow -SeenInCurrentCrawl:$false
}

