[CmdletBinding()]
param(
    [string]$TrackerPath = "",
    [string]$JobId = "",
    [string]$Url = "",
    [string]$TitleContains = "",
    [Parameter(Mandatory = $true)]
    [ValidateSet("new", "interesting", "ignored", "applied", "interview", "offer", "rejected", "withdrawn")]
    [string]$Status,
    [string]$AppliedDate = "",
    [string]$Notes = "",
    [switch]$AppendNotes,
    [int]$MaxBackups = 5
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "JobTracker.Common.ps1")

if ([string]::IsNullOrWhiteSpace($TrackerPath)) {
    $TrackerPath = Join-Path $PSScriptRoot "output\jobs_tracker.xlsx"
}

if (-not (Test-Path $TrackerPath)) {
    throw "Tracker file not found: $TrackerPath"
}
if ([IO.Path]::GetExtension($TrackerPath).ToLowerInvariant() -ne ".xlsx") {
    throw "This project uses only the XLSX tracker file. Use output\jobs_tracker.xlsx for -TrackerPath."
}

if ([string]::IsNullOrWhiteSpace($JobId) -and [string]::IsNullOrWhiteSpace($Url) -and [string]::IsNullOrWhiteSpace($TitleContains)) {
    throw "Provide -JobId, -Url, or -TitleContains."
}

$ColumnLabels = Get-JobTrackerColumnLabels

function Backup-TrackerFile {
    param([string]$Path)

    $backupDirectory = Join-Path (Split-Path -Parent $Path) "backups"
    if (-not (Test-Path $backupDirectory)) {
        New-Item -ItemType Directory -Force -Path $backupDirectory | Out-Null
    }

    $stamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $baseName = [IO.Path]::GetFileNameWithoutExtension($Path)
    $extension = [IO.Path]::GetExtension($Path)
    $backupPath = Join-Path $backupDirectory ("{0}_before_status_{1}{2}" -f $baseName, $stamp, $extension)
    Copy-Item -LiteralPath $Path -Destination $backupPath -Force

    if ($MaxBackups -gt 0) {
        $oldBackups = @(Get-ChildItem -LiteralPath $backupDirectory -File -Filter ("{0}_*{1}" -f $baseName, $extension) | Sort-Object LastWriteTime -Descending)
        if ($oldBackups.Count -gt $MaxBackups) {
            $oldBackups | Select-Object -Skip $MaxBackups | ForEach-Object {
                Remove-Item -LiteralPath $_.FullName -Force
            }
        }
    }

    return $backupPath
}

function Get-HeaderMap {
    param(
        $Sheet,
        [int]$ColumnCount
    )

    $headers = @{}
    for ($column = 1; $column -le $ColumnCount; $column++) {
        $header = [string]$Sheet.Cells.Item(1, $column).Text
        if (-not [string]::IsNullOrWhiteSpace($header)) {
            $headers[(ConvertTo-CanonicalColumnName $header.Trim())] = $column
        }
    }

    return $headers
}

function Get-CellText {
    param(
        $Sheet,
        [hashtable]$Headers,
        [int]$Row,
        [string]$Name
    )

    if (-not $Headers.ContainsKey($Name)) {
        return ""
    }

    return [string]$Sheet.Cells.Item($Row, [int]$Headers[$Name]).Text
}

function Get-StatusFontColor {
    param([string]$Value)

    switch ($Value.ToLowerInvariant()) {
        "interesting" { return (Get-ExcelColor 146 64 14) }
        "applied" { return (Get-ExcelColor 34 113 72) }
        "interview" { return (Get-ExcelColor 37 99 235) }
        "offer" { return (Get-ExcelColor 34 113 72) }
        "ignored" { return (Get-ExcelColor 100 116 139) }
        "rejected" { return (Get-ExcelColor 185 28 28) }
        "withdrawn" { return (Get-ExcelColor 100 116 139) }
        default { return (Get-ExcelColor 40 47 52) }
    }
}

function Test-IsStrongStatus {
    param([string]$Value)

    return $Value -in @("interesting", "applied", "interview", "offer")
}

function Find-XlsxMatches {
    param(
        $Sheet,
        [hashtable]$Headers,
        [int]$RowCount
    )

    $matches = New-Object System.Collections.Generic.List[object]
    for ($row = 2; $row -le $RowCount; $row++) {
        $jobIdValue = Get-CellText -Sheet $Sheet -Headers $Headers -Row $row -Name "job_id"
        $urlValue = Get-CellText -Sheet $Sheet -Headers $Headers -Row $row -Name "job_url_raw"
        $titleValue = Get-CellText -Sheet $Sheet -Headers $Headers -Row $row -Name "job_title"

        $isMatch = $false
        if (-not [string]::IsNullOrWhiteSpace($JobId) -and $jobIdValue -eq $JobId) {
            $isMatch = $true
        }
        if (-not [string]::IsNullOrWhiteSpace($Url) -and $urlValue -eq $Url) {
            $isMatch = $true
        }
        if (-not [string]::IsNullOrWhiteSpace($TitleContains) -and $titleValue -like "*$TitleContains*") {
            $isMatch = $true
        }

        if ($isMatch) {
            $matches.Add([PSCustomObject]@{
                Row = $row
                job_id = $jobIdValue
                job_title = $titleValue
                company_name = Get-CellText -Sheet $Sheet -Headers $Headers -Row $row -Name "company_name"
                location = Get-CellText -Sheet $Sheet -Headers $Headers -Row $row -Name "location"
                status = Get-CellText -Sheet $Sheet -Headers $Headers -Row $row -Name "status"
                notes = Get-CellText -Sheet $Sheet -Headers $Headers -Row $row -Name "notes"
            }) | Out-Null
        }
    }

    return @($matches.ToArray())
}

function Update-XlsxTracker {
    param([string]$Path)

    $fullPath = (Resolve-Path $Path).Path
    $excel = $null
    $workbook = $null
    $sheet = $null
    $usedRange = $null

    try {
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
        $headers = Get-HeaderMap -Sheet $sheet -ColumnCount ([int]$usedRange.Columns.Count)
        $matches = @(Find-XlsxMatches -Sheet $sheet -Headers $headers -RowCount ([int]$usedRange.Rows.Count))
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

    if ($matches.Count -eq 0) {
        throw "No matching job found."
    }
    if ($matches.Count -gt 1) {
        $matches | Select-Object job_id, job_title, company_name, location, status | Format-Table -AutoSize
        throw "More than one job matched. Rerun with -JobId for an exact update."
    }

    $targetRow = [int]$matches[0].Row

    if ($Status -eq "ignored" -and [string]::IsNullOrWhiteSpace($Notes) -and [string]::IsNullOrWhiteSpace([string]$matches[0].notes)) {
        $examples = (Get-JobTrackerIgnoreReasonOptions | Select-Object -First 5) -join " | "
        throw "Ignored jobs need Apply notes so future scoring can learn why. Example notes: $examples"
    }

    if ($Status -eq "ignored" -and -not [string]::IsNullOrWhiteSpace($Notes) -and [string]::IsNullOrWhiteSpace((Get-IgnoreReasonFromNotes $Notes))) {
        Write-Host "Tip: start ignored notes with ignore_reason=... so future crawls can use the feedback precisely."
    }

    $backupPath = Backup-TrackerFile $Path

    $excel = $null
    $workbook = $null
    $sheet = $null
    $usedRange = $null

    try {
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
        $headers = Get-HeaderMap -Sheet $sheet -ColumnCount ([int]$usedRange.Columns.Count)

        if (-not $headers.ContainsKey("status")) {
            throw "Tracker is missing the status column."
        }
        $statusCell = $sheet.Cells.Item($targetRow, [int]$headers["status"])
        $statusCell.Value2 = $Status
        Clear-CellFill $statusCell
        $statusCell.Font.Color = Get-StatusFontColor $Status
        $statusCell.Font.Bold = Test-IsStrongStatus $Status

        if ($headers.ContainsKey("applied_date")) {
            if (-not [string]::IsNullOrWhiteSpace($AppliedDate)) {
                $sheet.Cells.Item($targetRow, [int]$headers["applied_date"]).Value2 = $AppliedDate
            }
            elseif ($Status -eq "applied" -and [string]::IsNullOrWhiteSpace((Get-CellText -Sheet $sheet -Headers $headers -Row $targetRow -Name "applied_date"))) {
                $sheet.Cells.Item($targetRow, [int]$headers["applied_date"]).Value2 = Get-Date -Format "yyyy-MM-dd"
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($Notes) -and $headers.ContainsKey("notes")) {
            $notesCell = $sheet.Cells.Item($targetRow, [int]$headers["notes"])
            $existingNotes = [string]$notesCell.Text
            if ($AppendNotes -and -not [string]::IsNullOrWhiteSpace($existingNotes)) {
                $notesCell.Value2 = "{0} | {1}" -f $existingNotes, $Notes
            }
            else {
                $notesCell.Value2 = $Notes
            }
        }

        $workbook.Save() | Out-Null
    }
    finally {
        if ($null -ne $workbook) {
            $workbook.Close($true) | Out-Null
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

    Write-Host ("Updated {0} -> {1}" -f $matches[0].job_id, $Status)
    Write-Host ("Backup: {0}" -f $backupPath)
}

Update-XlsxTracker -Path $TrackerPath
