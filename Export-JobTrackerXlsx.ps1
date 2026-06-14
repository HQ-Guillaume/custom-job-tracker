[CmdletBinding()]
param(
    [string]$TrackerPath = ""
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "JobTracker.Common.ps1")

if ([string]::IsNullOrWhiteSpace($TrackerPath)) {
    $TrackerPath = Join-Path $PSScriptRoot "output\jobs_tracker.xlsx"
}

if (-not (Test-Path $TrackerPath)) {
    throw "Tracker workbook not found: $TrackerPath"
}

$ColumnLabels = Get-JobTrackerColumnLabels

function Get-HeaderMap {
    param($Sheet, [int]$ColumnCount)

    $headers = @{}
    for ($column = 1; $column -le $ColumnCount; $column++) {
        $header = [string]$Sheet.Cells.Item(1, $column).Text
        if (-not [string]::IsNullOrWhiteSpace($header)) {
            $canonicalName = ConvertTo-CanonicalColumnName $header.Trim()
            $headers[$canonicalName] = $column
            $Sheet.Cells.Item(1, $column).Value2 = Get-ColumnLabel $canonicalName
        }
    }

    return $headers
}

$excel = $null
$workbook = $null
$sheet = $null
$usedRange = $null
$tableRange = $null
$dataRange = $null

try {
    $fullPath = (Resolve-Path $TrackerPath).Path
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($fullPath)
    try {
        $sheet = $workbook.Worksheets.Item("Jobs")
    }
    catch {
        $sheet = $workbook.Worksheets.Item(1)
    }

    $usedRange = $sheet.UsedRange
    $rowCount = [int]$usedRange.Rows.Count
    $columnCount = [int]$usedRange.Columns.Count
    $headers = Get-HeaderMap -Sheet $sheet -ColumnCount $columnCount

    $darkTextColor = Get-ExcelColor 40 47 52
    $mutedTextColor = Get-ExcelColor 100 116 139
    $sheet.Cells.Font.Name = "Segoe UI"
    $sheet.Cells.Font.Size = 10
    $sheet.Cells.Font.Color = $darkTextColor
    $sheet.Rows.Item(1).Font.Bold = $true
    $sheet.Rows.Item(1).Font.Color = Get-ExcelColor 255 255 255
    $sheet.Rows.Item(1).Interior.Color = Get-ExcelColor 38 50 56
    $sheet.Rows.Item(1).VerticalAlignment = -4108
    $sheet.Rows.Item(1).RowHeight = 24

    if ($sheet.ListObjects.Count -eq 0 -and $rowCount -ge 1 -and $columnCount -ge 1) {
        $tableRange = $sheet.Range($sheet.Cells.Item(1, 1), $sheet.Cells.Item([Math]::Max(2, $rowCount), $columnCount))
        try {
            $table = $sheet.ListObjects.Add(1, $tableRange, $null, 1)
            $table.Name = "JobsTracker"
            $table.TableStyle = "TableStyleLight9"
            $table.ShowTableStyleRowStripes = $false
            $table.ShowTableStyleColumnStripes = $false
            Release-ComObject $table
        }
        catch {
            $sheet.Range($sheet.Cells.Item(1, 1), $sheet.Cells.Item(1, $columnCount)).AutoFilter() | Out-Null
        }
    }
    elseif ($sheet.ListObjects.Count -gt 0) {
        for ($tableIndex = 1; $tableIndex -le [int]$sheet.ListObjects.Count; $tableIndex++) {
            $existingTable = $sheet.ListObjects.Item($tableIndex)
            $existingTable.TableStyle = "TableStyleLight9"
            $existingTable.ShowTableStyleRowStripes = $false
            $existingTable.ShowTableStyleColumnStripes = $false
            Release-ComObject $existingTable
        }
    }

    foreach ($columnName in @("duplicate_reason", "job_title", "matched_keywords", "job_url_raw", "notes")) {
        if ($headers.ContainsKey($columnName)) {
            $sheet.Columns.Item([int]$headers[$columnName]).WrapText = $true
        }
    }
    $sheet.Columns.AutoFit() | Out-Null
    $columnSizing = Get-JobTrackerColumnSizing
    foreach ($entry in $columnSizing.GetEnumerator()) {
        if ($headers.ContainsKey($entry.Key)) {
            $column = $sheet.Columns.Item([int]$headers[$entry.Key])
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
    Set-JobTrackerColumnVisibility -Sheet $sheet -ColumnIndex $headers
    foreach ($columnName in @("review_priority", "status", "contract_type", "platform", "source_count", "published_date", "days_since_published", "job_url", "applied_date", "match_level", "match_score", "seen_in_current_crawl", "first_seen_date", "last_seen_date", "is_new")) {
        if ($headers.ContainsKey($columnName)) {
            $sheet.Columns.Item([int]$headers[$columnName]).HorizontalAlignment = -4108
        }
    }
    foreach ($columnName in @("published_date", "applied_date", "first_seen_date", "last_seen_date")) {
        if ($headers.ContainsKey($columnName)) {
            $sheet.Columns.Item([int]$headers[$columnName]).NumberFormat = "yyyy-mm-dd"
        }
    }
    foreach ($columnName in @("source_count", "days_since_published", "match_score", "days_since_first_seen", "days_since_last_seen", "feedback_adjustment")) {
        if ($headers.ContainsKey($columnName)) {
            $sheet.Columns.Item([int]$headers[$columnName]).NumberFormat = "0"
        }
    }

    Set-JobTrackerDataValidation -Workbook $workbook -Excel $excel -Sheet $sheet -ColumnIndex $headers -LastDataRow $rowCount

    if ($rowCount -gt 1) {
        $dataRange = $sheet.Range($sheet.Cells.Item(2, 1), $sheet.Cells.Item($rowCount, $columnCount))
        $dataRange.VerticalAlignment = -4160
        $dataRange.RowHeight = 30
        $dataRange.Interior.Color = Get-ExcelColor 255 255 255
        Set-StatusRowConditionalFormatting -Range $dataRange -ColumnIndex $headers
        Set-IgnoredNotesReminderFormatting -Sheet $sheet -ColumnIndex $headers -LastDataRow $rowCount
    }

    $colors = Get-JobTrackerWorkbookColors -DarkTextColor $darkTextColor
    for ($row = 2; $row -le $rowCount; $row++) {
        if ($headers.ContainsKey("job_url") -and $headers.ContainsKey("job_url_raw")) {
            $url = [string]$sheet.Cells.Item($row, [int]$headers["job_url_raw"]).Text
            $cell = $sheet.Cells.Item($row, [int]$headers["job_url"])
            if ($url -match "^https?://") {
                $escapedUrl = $url.Replace('"', '""')
                $cell.Formula = '=HYPERLINK("{0}","Open")' -f $escapedUrl
            }
        }

        if ($headers.ContainsKey("status")) {
            $status = ([string]$sheet.Cells.Item($row, [int]$headers["status"]).Text).ToLowerInvariant()
            $statusCell = $sheet.Cells.Item($row, [int]$headers["status"])
            Clear-CellFill $statusCell
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

        if ($headers.ContainsKey("match_level")) {
            $match = [string]$sheet.Cells.Item($row, [int]$headers["match_level"]).Text
            $matchCell = $sheet.Cells.Item($row, [int]$headers["match_level"])
            Clear-CellFill $matchCell
            $matchCell.Font.Bold = $true
            switch ($match) {
                "High" { $matchCell.Font.Color = $colors.GreenText }
                "Medium" { $matchCell.Font.Color = $colors.AmberText }
                "Review" { $matchCell.Font.Color = $colors.GrayText; $matchCell.Font.Bold = $false }
                default { $matchCell.Font.Color = $colors.DarkText }
            }
        }

        if ($headers.ContainsKey("review_priority")) {
            $priority = [string]$sheet.Cells.Item($row, [int]$headers["review_priority"]).Text
            $priorityCell = $sheet.Cells.Item($row, [int]$headers["review_priority"])
            Clear-CellFill $priorityCell
            $priorityCell.Font.Bold = $true
            switch ($priority) {
                "Application" { $priorityCell.Font.Color = $colors.GreenText }
                "New High" { $priorityCell.Font.Color = $colors.AmberText }
                "Ignored" { $priorityCell.Font.Color = $colors.GrayText; $priorityCell.Font.Bold = $false }
                default { $priorityCell.Font.Color = $colors.DarkText; $priorityCell.Font.Bold = $false }
            }
        }

        if ($headers.ContainsKey("seen_in_current_crawl")) {
            $seenNow = [string]$sheet.Cells.Item($row, [int]$headers["seen_in_current_crawl"]).Text
            if ($seenNow -eq "no") {
                $sheet.Cells.Item($row, [int]$headers["seen_in_current_crawl"]).Font.Color = $mutedTextColor
            }
        }
    }

    $sheet.Activate() | Out-Null
    $excel.ActiveWindow.SplitRow = 1
    $excel.ActiveWindow.SplitColumn = 2
    $excel.ActiveWindow.FreezePanes = $true
    $excel.ActiveWindow.DisplayGridlines = $false
    $workbook.Save() | Out-Null
}
finally {
    if ($null -ne $workbook) {
        $workbook.Close($true) | Out-Null
    }
    if ($null -ne $excel) {
        $excel.Quit() | Out-Null
    }
    Release-ComObject $tableRange
    Release-ComObject $dataRange
    Release-ComObject $usedRange
    Release-ComObject $sheet
    Release-ComObject $workbook
    Release-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host ("Formatted workbook: {0}" -f (Resolve-Path $TrackerPath).Path)
