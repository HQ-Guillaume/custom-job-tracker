$script:JobTrackerColumnLabels = [ordered]@{
    review_priority       = "Priority"
    status                = "Status"
    job_title             = "Job title"
    company_name          = "Company"
    location              = "City / region"
    contract_type         = "Contract"
    platform              = "Sources"
    source_count          = "Source count"
    published_date        = "Published"
    days_since_published  = "Age"
    match_level           = "Match"
    match_score           = "Score"
    job_url               = "Link"
    applied_date          = "Applied date"
    notes                 = "Apply notes"
    matched_keywords      = "Why it matched"
    seen_in_current_crawl = "Seen now"
    first_seen_date       = "First seen"
    last_seen_date        = "Last seen"
    is_new                = "New?"
    duplicate_reason      = "Duplicate / retention note"
    feedback_adjustment   = "History adjustment"
    job_id                = "Job ID"
    job_url_raw           = "Raw URL"
    alternate_urls        = "Other URLs"
    seen_before           = "Seen before"
    days_since_first_seen = "Days since first seen"
    days_since_last_seen  = "Days since last seen"
}

$script:JobTrackerMasterColumns = @(
    "review_priority",
    "status",
    "job_title",
    "company_name",
    "location",
    "contract_type",
    "platform",
    "source_count",
    "published_date",
    "days_since_published",
    "job_url",
    "applied_date",
    "notes",
    "match_level",
    "match_score",
    "matched_keywords",
    "seen_in_current_crawl",
    "first_seen_date",
    "last_seen_date",
    "is_new",
    "duplicate_reason",
    "feedback_adjustment",
    "job_id",
    "job_url_raw",
    "alternate_urls",
    "seen_before",
    "days_since_first_seen",
    "days_since_last_seen"
)

$script:JobTrackerDailyReviewColumns = @(
    "review_priority",
    "status",
    "job_title",
    "company_name",
    "location",
    "contract_type",
    "platform",
    "published_date",
    "days_since_published",
    "job_url",
    "applied_date",
    "notes",
    "match_level",
    "matched_keywords"
)

$script:JobTrackerHiddenWorkbookColumns = @(
    "source_count",
    "match_score",
    "seen_in_current_crawl",
    "first_seen_date",
    "last_seen_date",
    "is_new",
    "duplicate_reason",
    "feedback_adjustment",
    "job_id",
    "job_url_raw",
    "alternate_urls",
    "seen_before",
    "days_since_first_seen",
    "days_since_last_seen"
)

function Get-JobTrackerColumnLabels {
    $labels = @{}
    foreach ($key in $script:JobTrackerColumnLabels.Keys) {
        $labels[$key] = $script:JobTrackerColumnLabels[$key]
    }

    return $labels
}

function Get-JobTrackerMasterColumns {
    return @($script:JobTrackerMasterColumns)
}

function Get-JobTrackerDailyReviewColumns {
    return @($script:JobTrackerDailyReviewColumns)
}

function Get-JobTrackerHiddenWorkbookColumns {
    return @($script:JobTrackerHiddenWorkbookColumns)
}

function Get-JobTrackerStatusOptions {
    return @("new", "interesting", "ignored", "applied", "interview", "offer", "rejected", "withdrawn")
}

function Get-JobTrackerIgnoreReasonOptions {
    return @(
        "ignore_reason=not_analytics_enough; detail=",
        "ignore_reason=too_seo_sea_marketing; detail=",
        "ignore_reason=too_data_analyst; detail=",
        "ignore_reason=too_data_engineering; detail=",
        "ignore_reason=too_bi_reporting; detail=",
        "ignore_reason=too_crm_emailing; detail=",
        "ignore_reason=too_content_social; detail=",
        "ignore_reason=too_product_analytics; detail=",
        "ignore_reason=too_managerial; detail=",
        "ignore_reason=wrong_seniority; detail=",
        "ignore_reason=wrong_location; detail=",
        "ignore_reason=wrong_remote_policy; detail=",
        "ignore_reason=wrong_contract; detail=",
        "ignore_reason=language_issue; detail=",
        "ignore_reason=salary_issue; detail=",
        "ignore_reason=company_not_interested; detail=",
        "ignore_reason=industry_not_interested; detail=",
        "ignore_reason=duplicate; detail=",
        "ignore_reason=low_quality_posting; detail=",
        "ignore_reason=other; detail="
    )
}

function Get-JobTrackerIgnoreReasonKeys {
    $keys = New-Object System.Collections.Generic.List[string]
    foreach ($option in (Get-JobTrackerIgnoreReasonOptions)) {
        $match = [regex]::Match($option, "(?i)\bignore_reason\s*=\s*(?<reason>[a-z0-9_]+)")
        if ($match.Success) {
            $keys.Add($match.Groups["reason"].Value.ToLowerInvariant()) | Out-Null
        }
    }

    return @($keys.ToArray())
}

function ConvertTo-IgnoreReasonKey {
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ""
    }

    return (([string]$Value).Trim().ToLowerInvariant() -replace "[^a-z0-9]+", "_").Trim("_")
}

function Get-IgnoreReasonFromNotes {
    param([AllowNull()][string]$Notes)

    if ([string]::IsNullOrWhiteSpace($Notes)) {
        return ""
    }

    $noteText = [string]$Notes
    $tagMatch = [regex]::Match($noteText, "(?i)\bignore_reason\s*=\s*(?<reason>[a-z0-9_ -]+)")
    if ($tagMatch.Success) {
        return ConvertTo-IgnoreReasonKey $tagMatch.Groups["reason"].Value
    }

    $normalized = ConvertTo-IgnoreReasonKey $noteText
    $knownReasons = Get-JobTrackerIgnoreReasonKeys
    foreach ($reason in $knownReasons) {
        if ($normalized -match [regex]::Escape($reason)) {
            return $reason
        }
    }

    if ($normalized -match "seo|sea|paid|marketing|acquisition|growth") {
        return "too_seo_sea_marketing"
    }
    if ($normalized -match "not_analytics|not_analytic|no_analytics|not_relevant") {
        return "not_analytics_enough"
    }
    if ($normalized -match "data_analyst|python|sql|warehouse") {
        return "too_data_analyst"
    }
    if ($normalized -match "data_engineer|analytics_engineer|dbt|snowflake|airflow|etl") {
        return "too_data_engineering"
    }
    if ($normalized -match "bi|reporting|dashboard") {
        return "too_bi_reporting"
    }

    return ""
}

function Get-JobTrackerColumnSizing {
    return @{
        review_priority       = @{ Min = 11; Max = 16 }
        status                = @{ Min = 10; Max = 14 }
        job_title             = @{ Min = 30; Max = 54 }
        company_name          = @{ Min = 16; Max = 30 }
        location              = @{ Min = 14; Max = 28 }
        contract_type         = @{ Min = 10; Max = 18 }
        platform              = @{ Min = 12; Max = 28 }
        source_count          = @{ Min = 8; Max = 12 }
        published_date        = @{ Min = 12; Max = 14 }
        days_since_published  = @{ Min = 7; Max = 8 }
        job_url               = @{ Min = 8; Max = 9 }
        applied_date          = @{ Min = 12; Max = 14 }
        notes                 = @{ Min = 26; Max = 56 }
        match_level           = @{ Min = 9; Max = 12 }
        match_score           = @{ Min = 7; Max = 8 }
        matched_keywords      = @{ Min = 26; Max = 60 }
        seen_in_current_crawl = @{ Min = 8; Max = 10 }
        first_seen_date       = @{ Min = 12; Max = 14 }
        last_seen_date        = @{ Min = 12; Max = 14 }
        is_new                = @{ Min = 7; Max = 8 }
        duplicate_reason      = @{ Min = 18; Max = 42 }
        feedback_adjustment   = @{ Min = 10; Max = 14 }
        job_id                = @{ Min = 14; Max = 18 }
        job_url_raw           = @{ Min = 24; Max = 48 }
        alternate_urls        = @{ Min = 24; Max = 48 }
        seen_before           = @{ Min = 9; Max = 12 }
        days_since_first_seen = @{ Min = 14; Max = 18 }
        days_since_last_seen  = @{ Min = 14; Max = 18 }
    }
}

function Get-JobTrackerWorkbookColors {
    param([int]$DarkTextColor = (Get-ExcelColor 40 47 52))

    return @{
        GreenText = Get-ExcelColor 34 113 72
        BlueText  = Get-ExcelColor 37 99 235
        AmberText = Get-ExcelColor 146 64 14
        RedText   = Get-ExcelColor 185 28 28
        GrayText  = Get-ExcelColor 100 116 139
        DarkText  = $DarkTextColor
    }
}

function ConvertTo-ColumnLookupKey {
    param([AllowNull()][string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return ""
    }

    return (([string]$Name).ToLowerInvariant() -replace "[^a-z0-9]+", "")
}

function Get-JobTrackerColumnAliases {
    param([string[]]$Columns = $script:JobTrackerMasterColumns)

    $aliases = @{}
    foreach ($columnName in $Columns) {
        $aliases[(ConvertTo-ColumnLookupKey $columnName)] = $columnName
        if ($script:JobTrackerColumnLabels.Contains($columnName)) {
            $aliases[(ConvertTo-ColumnLookupKey $script:JobTrackerColumnLabels[$columnName])] = $columnName
        }
    }
    $aliases[(ConvertTo-ColumnLookupKey "Platform")] = "platform"

    return $aliases
}

function ConvertTo-CanonicalColumnName {
    param([AllowNull()][string]$Name)

    $key = ConvertTo-ColumnLookupKey $Name
    $aliases = Get-JobTrackerColumnAliases
    if ($aliases.ContainsKey($key)) {
        return [string]$aliases[$key]
    }

    return $Name
}

function Get-ColumnLabel {
    param([string]$ColumnName)

    if ($script:JobTrackerColumnLabels.Contains($ColumnName)) {
        return [string]$script:JobTrackerColumnLabels[$ColumnName]
    }

    return $ColumnName
}

function Release-ComObject {
    param([AllowNull()]$Object)

    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
    }
}

function Get-ExcelColor {
    param([int]$Red, [int]$Green, [int]$Blue)

    return $Red + ($Green * 256) + ($Blue * 65536)
}

function ConvertTo-ExcelColumnName {
    param([int]$ColumnNumber)

    $name = ""
    while ($ColumnNumber -gt 0) {
        $modulo = ($ColumnNumber - 1) % 26
        $name = ([string][char](65 + $modulo)) + $name
        $ColumnNumber = [Math]::Floor(($ColumnNumber - $modulo) / 26)
    }

    return $name
}

function Clear-CellFill {
    param($Cell)

    try {
        $Cell.Interior.Pattern = -4142
    }
    catch {
    }
}

function Set-StatusRowConditionalFormatting {
    param(
        [AllowNull()]$Range,
        [hashtable]$ColumnIndex
    )

    if ($null -eq $Range -or -not $ColumnIndex.ContainsKey("status")) {
        return
    }

    $statusColumn = ConvertTo-ExcelColumnName ([int]$ColumnIndex["status"])
    $rules = @(
        @{ Status = "interesting"; Color = Get-ExcelColor 255 249 235 },
        @{ Status = "applied"; Color = Get-ExcelColor 240 253 244 },
        @{ Status = "interview"; Color = Get-ExcelColor 239 246 255 },
        @{ Status = "offer"; Color = Get-ExcelColor 236 253 245 },
        @{ Status = "ignored"; Color = Get-ExcelColor 248 250 252 },
        @{ Status = "rejected"; Color = Get-ExcelColor 254 242 242 },
        @{ Status = "withdrawn"; Color = Get-ExcelColor 245 245 245 }
    )

    try {
        $Range.FormatConditions.Delete()
        foreach ($rule in $rules) {
            $formula = '=${0}2="{1}"' -f $statusColumn, $rule.Status
            $condition = $Range.FormatConditions.Add(2, 0, $formula)
            $condition.Interior.Color = $rule.Color
            Release-ComObject $condition
        }
    }
    catch {
    }
}

function Set-IgnoredNotesReminderFormatting {
    param(
        $Sheet,
        [hashtable]$ColumnIndex,
        [int]$LastDataRow
    )

    if ($null -eq $Sheet -or -not $ColumnIndex.ContainsKey("status") -or -not $ColumnIndex.ContainsKey("notes")) {
        return
    }

    $lastRow = [Math]::Max(2, $LastDataRow)
    $statusColumn = ConvertTo-ExcelColumnName ([int]$ColumnIndex["status"])
    $notesColumn = ConvertTo-ExcelColumnName ([int]$ColumnIndex["notes"])

    try {
        $notesRange = $Sheet.Range($Sheet.Cells.Item(2, [int]$ColumnIndex["notes"]), $Sheet.Cells.Item($lastRow, [int]$ColumnIndex["notes"]))
        $formula = '=AND(LOWER(${0}2)="ignored",LEN(TRIM(${1}2))=0)' -f $statusColumn, $notesColumn
        $condition = $notesRange.FormatConditions.Add(2, 0, $formula)
        $condition.Interior.Color = Get-ExcelColor 255 251 235
        $condition.Font.Color = Get-ExcelColor 146 64 14
        $condition.Font.Bold = $true
        Release-ComObject $condition
        Release-ComObject $notesRange
    }
    catch {
    }
}

function Set-JobTrackerValidationLists {
    param($Workbook)

    $sheetName = "_validation"
    $listSheet = $null

    try {
        $listSheet = $Workbook.Worksheets.Item($sheetName)
    }
    catch {
        $listSheet = $Workbook.Worksheets.Add([System.Type]::Missing, $Workbook.Worksheets.Item([int]$Workbook.Worksheets.Count))
        $listSheet.Name = $sheetName
    }

    $listSheet.Cells.Clear() | Out-Null
    $statusOptions = @(Get-JobTrackerStatusOptions)
    $ignoreOptions = @(Get-JobTrackerIgnoreReasonOptions)

    for ($index = 0; $index -lt $statusOptions.Count; $index++) {
        $listSheet.Cells.Item($index + 1, 1).Value2 = [string]$statusOptions[$index]
    }
    for ($index = 0; $index -lt $ignoreOptions.Count; $index++) {
        $listSheet.Cells.Item($index + 1, 2).Value2 = [string]$ignoreOptions[$index]
    }

    foreach ($name in @("JobTrackerStatusOptions", "JobTrackerApplyNoteTemplates")) {
        try {
            $Workbook.Names.Item($name).Delete()
        }
        catch {
        }
    }

    $escapedSheetName = $sheetName.Replace("'", "''")
    $Workbook.Names.Add("JobTrackerStatusOptions", ("='{0}'!`$A`$1:`$A`${1}" -f $escapedSheetName, $statusOptions.Count)) | Out-Null
    $Workbook.Names.Add("JobTrackerApplyNoteTemplates", ("='{0}'!`$B`$1:`$B`${1}" -f $escapedSheetName, $ignoreOptions.Count)) | Out-Null
    $listSheet.Visible = 0

    Release-ComObject $listSheet

    return @{
        StatusFormula = "=JobTrackerStatusOptions"
        NotesFormula  = "=JobTrackerApplyNoteTemplates"
    }
}

function Set-JobTrackerDataValidation {
    param(
        $Workbook,
        $Excel,
        $Sheet,
        [hashtable]$ColumnIndex,
        [int]$LastDataRow
    )

    if ($null -eq $Workbook -or $null -eq $Sheet) {
        return
    }

    $validationEndRow = [Math]::Max($LastDataRow + 200, 500)
    $validationLists = Set-JobTrackerValidationLists -Workbook $Workbook

    if ($ColumnIndex.ContainsKey("status")) {
        try {
            $statusRange = $Sheet.Range($Sheet.Cells.Item(2, [int]$ColumnIndex["status"]), $Sheet.Cells.Item($validationEndRow, [int]$ColumnIndex["status"]))
            $statusRange.Validation.Delete()
            $statusRange.Validation.Add(3, 1, 1, $validationLists.StatusFormula)
            $statusRange.Validation.InCellDropdown = $true
            $statusRange.Validation.IgnoreBlank = $true
            $statusRange.Validation.ShowInput = $true
            $statusRange.Validation.InputTitle = "Application status"
            $statusRange.Validation.InputMessage = "Choose a status. If you choose ignored, fill Apply notes with an ignore_reason template."
            Release-ComObject $statusRange
        }
        catch {
        }
    }

    if ($ColumnIndex.ContainsKey("notes")) {
        try {
            $notesRange = $Sheet.Range($Sheet.Cells.Item(2, [int]$ColumnIndex["notes"]), $Sheet.Cells.Item($validationEndRow, [int]$ColumnIndex["notes"]))
            $notesRange.Validation.Delete()
            $notesRange.Validation.Add(3, 2, 1, $validationLists.NotesFormula)
            $notesRange.Validation.InCellDropdown = $true
            $notesRange.Validation.IgnoreBlank = $true
            $notesRange.Validation.ShowInput = $true
            $notesRange.Validation.ShowError = $true
            $notesRange.Validation.InputTitle = "Apply notes"
            $notesRange.Validation.InputMessage = "For ignored jobs, select an ignore_reason template, then add a short detail if helpful."
            $notesRange.Validation.ErrorTitle = "Use an ignore_reason when possible"
            $notesRange.Validation.ErrorMessage = "Free notes are allowed, but ignored jobs work best when notes start with ignore_reason=..."
            Release-ComObject $notesRange
        }
        catch {
        }
    }

    if ($ColumnIndex.ContainsKey("applied_date")) {
        try {
            $appliedDateRange = $Sheet.Range($Sheet.Cells.Item(2, [int]$ColumnIndex["applied_date"]), $Sheet.Cells.Item($validationEndRow, [int]$ColumnIndex["applied_date"]))
            $appliedDateRange.Validation.Delete()
            $appliedDateRange.Validation.Add(4, 1, 1, "2020-01-01", "2035-12-31")
            $appliedDateRange.Validation.InputTitle = "Applied date"
            $appliedDateRange.Validation.InputMessage = "Use yyyy-mm-dd when you mark a job as applied."
            Release-ComObject $appliedDateRange
        }
        catch {
        }
    }
}

function Set-JobTrackerColumnVisibility {
    param(
        $Sheet,
        [hashtable]$ColumnIndex
    )

    foreach ($columnName in (Get-JobTrackerDailyReviewColumns)) {
        if ($ColumnIndex.ContainsKey($columnName)) {
            $Sheet.Columns.Item([int]$ColumnIndex[$columnName]).Hidden = $false
        }
    }

    foreach ($columnName in (Get-JobTrackerHiddenWorkbookColumns)) {
        if ($ColumnIndex.ContainsKey($columnName)) {
            $Sheet.Columns.Item([int]$ColumnIndex[$columnName]).Hidden = $true
        }
    }
}
