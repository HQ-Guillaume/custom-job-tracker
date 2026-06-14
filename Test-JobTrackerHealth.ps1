[CmdletBinding()]
param(
    [string]$TrackerPath = "",
    [switch]$WarnOnly
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "JobTracker.Common.ps1")

if ([string]::IsNullOrWhiteSpace($TrackerPath)) {
    $TrackerPath = Join-Path $PSScriptRoot "output\jobs_tracker.xlsx"
}

$errors = New-Object System.Collections.Generic.List[string]
$warnings = New-Object System.Collections.Generic.List[string]

function Add-HealthError {
    param([string]$Message)
    $errors.Add($Message) | Out-Null
}

function Add-HealthWarning {
    param([string]$Message)
    $warnings.Add($Message) | Out-Null
}

function Get-HeaderMap {
    param($Sheet, [int]$ColumnCount)

    $headers = @{}
    for ($column = 1; $column -le $ColumnCount; $column++) {
        $header = [string]$Sheet.Cells.Item(1, $column).Text
        if (-not [string]::IsNullOrWhiteSpace($header)) {
            $headers[(ConvertTo-CanonicalColumnName $header.Trim())] = $column
        }
    }

    return $headers
}

if (-not (Test-Path -LiteralPath $TrackerPath)) {
    throw "Tracker workbook not found: $TrackerPath"
}
if ([IO.Path]::GetExtension($TrackerPath).ToLowerInvariant() -ne ".xlsx") {
    throw "Tracker health check supports only XLSX files."
}

$excel = $null
$workbook = $null
$jobsSheet = $null
$summarySheet = $null
$usedRange = $null
$dataRange = $null

try {
    $fullPath = (Resolve-Path -LiteralPath $TrackerPath).Path
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($fullPath, 0, $true)

    try {
        $jobsSheet = $workbook.Worksheets.Item("Jobs")
    }
    catch {
        Add-HealthError "Missing Jobs sheet."
    }
    try {
        $summarySheet = $workbook.Worksheets.Item("Summary")
    }
    catch {
        Add-HealthWarning "Missing Summary sheet."
    }

    if ($null -ne $jobsSheet) {
        $usedRange = $jobsSheet.UsedRange
        $rowCount = [int]$usedRange.Rows.Count
        $columnCount = [int]$usedRange.Columns.Count
        $headers = Get-HeaderMap -Sheet $jobsSheet -ColumnCount $columnCount
        $expectedColumns = Get-JobTrackerMasterColumns

        foreach ($columnName in $expectedColumns) {
            if (-not $headers.ContainsKey($columnName)) {
                Add-HealthError ("Missing workbook column: {0}" -f (Get-ColumnLabel $columnName))
            }
        }

        if ($headers.ContainsKey("job_title")) {
            $jobTitleColumn = [int]$headers["job_title"]
            $rowsWithTitle = 0
            for ($row = 2; $row -le $rowCount; $row++) {
                if (-not [string]::IsNullOrWhiteSpace([string]$jobsSheet.Cells.Item($row, $jobTitleColumn).Text)) {
                    $rowsWithTitle++
                }
            }
            if ($rowsWithTitle -eq 0) {
                Add-HealthWarning "No job rows with a Job title were found."
            }
        }
        else {
            $rowsWithTitle = 0
        }

        $visibleColumns = New-Object System.Collections.Generic.List[string]
        $hiddenColumns = New-Object System.Collections.Generic.List[string]
        foreach ($columnName in $expectedColumns) {
            if (-not $headers.ContainsKey($columnName)) {
                continue
            }

            if ([bool]$jobsSheet.Columns.Item([int]$headers[$columnName]).Hidden) {
                $hiddenColumns.Add($columnName) | Out-Null
            }
            else {
                $visibleColumns.Add($columnName) | Out-Null
            }
        }

        $expectedVisible = Get-JobTrackerDailyReviewColumns
        $expectedHidden = Get-JobTrackerHiddenWorkbookColumns
        foreach ($columnName in $expectedVisible) {
            if ($hiddenColumns.Contains($columnName)) {
                Add-HealthWarning ("Daily review column is hidden: {0}" -f (Get-ColumnLabel $columnName))
            }
        }
        foreach ($columnName in $expectedHidden) {
            if ($visibleColumns.Contains($columnName)) {
                Add-HealthWarning ("Backend column is visible: {0}" -f (Get-ColumnLabel $columnName))
            }
        }

        if ($headers.ContainsKey("job_id")) {
            $jobIdColumn = [int]$headers["job_id"]
            $seenJobIds = @{}
            for ($row = 2; $row -le $rowCount; $row++) {
                $jobId = [string]$jobsSheet.Cells.Item($row, $jobIdColumn).Text
                if ([string]::IsNullOrWhiteSpace($jobId)) {
                    continue
                }
                if ($seenJobIds.ContainsKey($jobId)) {
                    Add-HealthWarning ("Duplicate Job ID found: {0}" -f $jobId)
                }
                else {
                    $seenJobIds[$jobId] = $true
                }
            }
        }

        if ($headers.ContainsKey("status")) {
            $statusColumn = [int]$headers["status"]
            $validStatuses = Get-JobTrackerStatusOptions
            for ($row = 2; $row -le $rowCount; $row++) {
                $status = ([string]$jobsSheet.Cells.Item($row, $statusColumn).Text).ToLowerInvariant()
                if (-not [string]::IsNullOrWhiteSpace($status) -and $status -notin $validStatuses) {
                    Add-HealthWarning ("Unexpected status on row {0}: {1}" -f $row, $status)
                }
            }
        }

        if ($headers.ContainsKey("status") -and $headers.ContainsKey("notes")) {
            $statusColumn = [int]$headers["status"]
            $notesColumn = [int]$headers["notes"]
            $validIgnoreReasons = Get-JobTrackerIgnoreReasonKeys
            $ignoredWithoutNotes = 0
            for ($row = 2; $row -le $rowCount; $row++) {
                $status = ([string]$jobsSheet.Cells.Item($row, $statusColumn).Text).ToLowerInvariant()
                $notes = [string]$jobsSheet.Cells.Item($row, $notesColumn).Text
                if ($status -eq "ignored" -and [string]::IsNullOrWhiteSpace($notes)) {
                    $ignoredWithoutNotes++
                    continue
                }

                $reason = Get-IgnoreReasonFromNotes $notes
                if (-not [string]::IsNullOrWhiteSpace($reason) -and $reason -notin $validIgnoreReasons) {
                    Add-HealthWarning ("Unexpected ignore_reason on row {0}: {1}" -f $row, $reason)
                }
            }

            if ($ignoredWithoutNotes -gt 0) {
                Add-HealthWarning ("Ignored rows without Apply notes: {0}" -f $ignoredWithoutNotes)
            }

            try {
                $notesValidationType = [int]$jobsSheet.Cells.Item(2, $notesColumn).Validation.Type
                if ($notesValidationType -ne 3) {
                    Add-HealthWarning "Apply notes column does not have the expected dropdown validation."
                }
            }
            catch {
                Add-HealthWarning "Apply notes column validation could not be inspected."
            }
        }

        if ($headers.ContainsKey("job_url") -and $headers.ContainsKey("job_title")) {
            $linkColumn = [int]$headers["job_url"]
            $titleColumn = [int]$headers["job_title"]
            $missingLinks = 0
            for ($row = 2; $row -le $rowCount; $row++) {
                $title = [string]$jobsSheet.Cells.Item($row, $titleColumn).Text
                if ([string]::IsNullOrWhiteSpace($title)) {
                    continue
                }

                $formula = [string]$jobsSheet.Cells.Item($row, $linkColumn).Formula
                if ($formula -notmatch '^=HYPERLINK\(') {
                    $missingLinks++
                }
            }
            if ($missingLinks -gt 0) {
                Add-HealthWarning ("Rows without clickable Link formulas: {0}" -f $missingLinks)
            }
        }

        if ($rowCount -gt 1 -and $columnCount -gt 0) {
            $dataRange = $jobsSheet.Range($jobsSheet.Cells.Item(2, 1), $jobsSheet.Cells.Item($rowCount, $columnCount))
            $formatRuleCount = [int]$dataRange.FormatConditions.Count
            if ($formatRuleCount -lt 7) {
                Add-HealthWarning ("Expected 7 status conditional-formatting rules; found {0}." -f $formatRuleCount)
            }
        }

        Write-Host ("Tracker: {0}" -f $fullPath)
        Write-Host ("Rows with job title: {0}" -f $rowsWithTitle)
        Write-Host ("Visible columns: {0}" -f (($visibleColumns | ForEach-Object { Get-ColumnLabel $_ }) -join " | "))
        Write-Host ("Hidden backend columns: {0}" -f (($hiddenColumns | ForEach-Object { Get-ColumnLabel $_ }) -join " | "))
    }
}
finally {
    if ($null -ne $workbook) {
        $workbook.Close($false) | Out-Null
    }
    if ($null -ne $excel) {
        $excel.Quit() | Out-Null
    }

    Release-ComObject $dataRange
    Release-ComObject $usedRange
    Release-ComObject $summarySheet
    Release-ComObject $jobsSheet
    Release-ComObject $workbook
    Release-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if ($warnings.Count -gt 0) {
    Write-Host ""
    Write-Host "Warnings:"
    foreach ($warning in $warnings) {
        Write-Host ("- {0}" -f $warning)
    }
}

if ($errors.Count -gt 0) {
    Write-Host ""
    Write-Host "Errors:"
    foreach ($item in $errors) {
        Write-Host ("- {0}" -f $item)
    }
    if (-not $WarnOnly) {
        exit 1
    }
}

if ($errors.Count -eq 0) {
    Write-Host ""
    Write-Host "Workbook health check passed."
}
