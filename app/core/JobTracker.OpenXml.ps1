function Initialize-OpenXmlZipAssemblies {
    Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
}

function ConvertTo-OpenXmlEscapedText {
    param([AllowNull()][string]$Value)

    if ($null -eq $Value) {
        return ""
    }

    return ([string]$Value).
        Replace("&", "&amp;").
        Replace("<", "&lt;").
        Replace(">", "&gt;").
        Replace('"', "&quot;").
        Replace("'", "&apos;")
}

function ConvertTo-OpenXmlSheetName {
    param(
        [AllowNull()][string]$Name,
        [string]$DefaultValue
    )

    $sheetName = [string]$Name
    if ([string]::IsNullOrWhiteSpace($sheetName)) {
        $sheetName = $DefaultValue
    }
    $sheetName = [regex]::Replace($sheetName.Trim(), "[][\\/*?:]", " ")
    if ($sheetName.Length -gt 31) {
        $sheetName = $sheetName.Substring(0, 31)
    }
    if ([string]::IsNullOrWhiteSpace($sheetName)) {
        return $DefaultValue
    }

    return $sheetName
}

function Get-OpenXmlScriptValue {
    param(
        [string]$Name,
        [AllowNull()]$DefaultValue = ""
    )

    $variable = Get-Variable -Name $Name -Scope Script -ErrorAction SilentlyContinue
    if ($null -eq $variable) {
        return $DefaultValue
    }

    return $variable.Value
}

function Write-OpenXmlUtf8File {
    param(
        [string]$Path,
        [string]$Content
    )

    $directory = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    $encoding = New-Object System.Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($Path, $Content, $encoding)
}

function Compress-OpenXmlDirectory {
    param(
        [string]$SourceDirectory,
        [string]$DestinationPath
    )

    Initialize-OpenXmlZipAssemblies
    if (Test-Path -LiteralPath $DestinationPath) {
        Remove-Item -LiteralPath $DestinationPath -Force
    }

    $zip = [System.IO.Compression.ZipFile]::Open($DestinationPath, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $sourceRoot = (Resolve-Path -LiteralPath $SourceDirectory).Path.TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
        foreach ($file in @(Get-ChildItem -LiteralPath $SourceDirectory -File -Recurse)) {
            $relativePath = $file.FullName.Substring($sourceRoot.Length).TrimStart([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
            $entryName = $relativePath.Replace("\", "/")
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $file.FullName, $entryName, [System.IO.Compression.CompressionLevel]::Optimal) | Out-Null
        }
    }
    finally {
        $zip.Dispose()
    }
}

function Get-OpenXmlCellReference {
    param(
        [int]$Column,
        [int]$Row
    )

    return "{0}{1}" -f (ConvertTo-ExcelColumnName $Column), $Row
}

function Test-OpenXmlNumericColumn {
    param([string]$ColumnName)

    return $ColumnName -in @(
        "source_count",
        "days_since_published",
        "match_score",
        "role_score",
        "employer_fit",
        "location_fit",
        "seniority_fit",
        "contract_fit",
        "days_since_first_seen",
        "days_since_last_seen",
        "feedback_adjustment"
    )
}

function Get-OpenXmlReviewPriorityFormula {
    param(
        [int]$Row,
        [hashtable]$ColumnIndex
    )

    if (-not $ColumnIndex.ContainsKey("status") -or
        -not $ColumnIndex.ContainsKey("match_level") -or
        -not $ColumnIndex.ContainsKey("is_new")) {
        return ""
    }

    $statusRef = '${0}{1}' -f (ConvertTo-ExcelColumnName ([int]$ColumnIndex["status"])), $Row
    $matchRef = '${0}{1}' -f (ConvertTo-ExcelColumnName ([int]$ColumnIndex["match_level"])), $Row
    $isNewRef = '${0}{1}' -f (ConvertTo-ExcelColumnName ([int]$ColumnIndex["is_new"])), $Row
    return 'IF(OR({0}="applied",{0}="interview",{0}="offer",{0}="rejected",{0}="withdrawn"),"Application",IF({0}="ignored","Ignored",IF(AND({1}="yes",{2}="High"),"New High",IF({1}="yes","New",IF({2}="High","High",{2})))))' -f $statusRef, $isNewRef, $matchRef
}

function New-OpenXmlCellXml {
    param(
        [int]$Row,
        [int]$Column,
        [AllowNull()][string]$Value,
        [int]$StyleId = 0,
        [AllowNull()][string]$Formula = "",
        [switch]$Numeric,
        [switch]$Hyperlink
    )

    $reference = Get-OpenXmlCellReference -Column $Column -Row $Row
    $styleAttribute = $(if ($StyleId -gt 0) { ' s="{0}"' -f $StyleId } else { "" })
    $cleanValue = Repair-DisplayText ([string]$Value)

    if (-not [string]::IsNullOrWhiteSpace($Formula)) {
        $escapedFormula = ConvertTo-OpenXmlEscapedText $Formula
        $escapedValue = ConvertTo-OpenXmlEscapedText $cleanValue
        return '<c r="{0}" t="str"{1}><f>{2}</f><v>{3}</v></c>' -f $reference, $styleAttribute, $escapedFormula, $escapedValue
    }

    if ($Numeric) {
        $number = 0
        if ([int]::TryParse($cleanValue, [ref]$number)) {
            return '<c r="{0}"{1}><v>{2}</v></c>' -f $reference, $styleAttribute, $number
        }
    }

    if ([string]::IsNullOrEmpty($cleanValue)) {
        return '<c r="{0}"{1}/>' -f $reference, $styleAttribute
    }

    $escapedText = ConvertTo-OpenXmlEscapedText $cleanValue
    $spaceAttribute = $(if ($cleanValue -match "^\s|\s$") { ' xml:space="preserve"' } else { "" })
    return '<c r="{0}" t="inlineStr"{1}><is><t{2}>{3}</t></is></c>' -f $reference, $styleAttribute, $spaceAttribute, $escapedText
}

function Get-OpenXmlColumnXml {
    param(
        [string[]]$Columns,
        [hashtable]$ColumnIndex
    )

    $sizing = Get-JobTrackerColumnSizing
    $hiddenColumns = @(Get-JobTrackerHiddenWorkbookColumns)
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($columnName in $Columns) {
        $columnNumber = [int]$ColumnIndex[$columnName]
        $width = 14
        if ($sizing.ContainsKey($columnName)) {
            $width = [double]$sizing[$columnName].Min
            $max = [double]$sizing[$columnName].Max
            if ($width -gt $max) {
                $width = $max
            }
        }
        $hidden = $(if ($columnName -in $hiddenColumns) { ' hidden="1"' } else { "" })
        $parts.Add(('<col min="{0}" max="{0}" width="{1}" customWidth="1"{2}/>' -f $columnNumber, $width, $hidden)) | Out-Null
    }

    return "<cols>{0}</cols>" -f (($parts.ToArray()) -join "")
}

function Get-OpenXmlStatusConditionalFormattingXml {
    param(
        [int]$LastColumn,
        [int]$LastDataRow,
        [hashtable]$ColumnIndex
    )

    if (-not $ColumnIndex.ContainsKey("status")) {
        return ""
    }

    $statusColumn = ConvertTo-ExcelColumnName ([int]$ColumnIndex["status"])
    $lastRow = [Math]::Max($LastDataRow + 200, 500)
    $statusRange = "{0}2:{0}{1}" -f $statusColumn, $lastRow
    $priorityRange = $(if ($ColumnIndex.ContainsKey("review_priority")) { "{0}2:{0}{1}" -f (ConvertTo-ExcelColumnName ([int]$ColumnIndex["review_priority"])), $lastRow } else { "" })

    $statuses = @("interesting", "applied", "interview", "offer", "ignored", "rejected", "withdrawn")
    $builder = New-Object System.Text.StringBuilder
    $priority = 1

    [void]$builder.Append('<conditionalFormatting sqref="')
    [void]$builder.Append($statusRange)
    [void]$builder.Append('">')
    for ($index = 0; $index -lt $statuses.Count; $index++) {
        [void]$builder.Append(('<cfRule type="expression" dxfId="{0}" priority="{1}"><formula>${2}2=&quot;{3}&quot;</formula></cfRule>' -f (7 + $index), $priority, $statusColumn, $statuses[$index]))
        $priority++
    }
    [void]$builder.Append('</conditionalFormatting>')

    if (-not [string]::IsNullOrWhiteSpace($priorityRange)) {
        [void]$builder.Append('<conditionalFormatting sqref="')
        [void]$builder.Append($priorityRange)
        [void]$builder.Append('">')
        for ($index = 0; $index -lt $statuses.Count; $index++) {
            [void]$builder.Append(('<cfRule type="expression" dxfId="{0}" priority="{1}"><formula>${2}2=&quot;{3}&quot;</formula></cfRule>' -f (7 + $index), $priority, $statusColumn, $statuses[$index]))
            $priority++
        }
        if ($ColumnIndex.ContainsKey("is_new") -and $ColumnIndex.ContainsKey("match_level")) {
            $isNewColumn = ConvertTo-ExcelColumnName ([int]$ColumnIndex["is_new"])
            $matchColumn = ConvertTo-ExcelColumnName ([int]$ColumnIndex["match_level"])
            [void]$builder.Append(('<cfRule type="expression" dxfId="14" priority="{0}"><formula>AND(${1}2=&quot;yes&quot;,${2}2=&quot;High&quot;,NOT(OR(${3}2=&quot;applied&quot;,${3}2=&quot;interview&quot;,${3}2=&quot;offer&quot;,${3}2=&quot;rejected&quot;,${3}2=&quot;withdrawn&quot;,${3}2=&quot;ignored&quot;)))</formula></cfRule>' -f $priority, $isNewColumn, $matchColumn, $statusColumn))
        }
        [void]$builder.Append('</conditionalFormatting>')
    }

    if ($ColumnIndex.ContainsKey("notes")) {
        $notesColumn = ConvertTo-ExcelColumnName ([int]$ColumnIndex["notes"])
        $notesRange = "{0}2:{0}{1}" -f $notesColumn, $lastRow
        [void]$builder.Append(('<conditionalFormatting sqref="{0}"><cfRule type="expression" dxfId="15" priority="99"><formula>AND(LOWER(${1}2)=&quot;ignored&quot;,LEN(TRIM(${2}2))=0)</formula></cfRule></conditionalFormatting>' -f $notesRange, $statusColumn, $notesColumn))
    }

    return $builder.ToString()
}

function Get-OpenXmlDataValidationsXml {
    param(
        [int]$LastDataRow,
        [hashtable]$ColumnIndex
    )

    $validationEndRow = [Math]::Max($LastDataRow + 200, 500)
    $items = New-Object System.Collections.Generic.List[string]
    if ($ColumnIndex.ContainsKey("status")) {
        $column = ConvertTo-ExcelColumnName ([int]$ColumnIndex["status"])
        $items.Add(('<dataValidation type="list" allowBlank="1" showInputMessage="1" sqref="{0}2:{0}{1}"><formula1>JobTrackerStatusOptions</formula1></dataValidation>' -f $column, $validationEndRow)) | Out-Null
    }
    if ($ColumnIndex.ContainsKey("notes")) {
        $column = ConvertTo-ExcelColumnName ([int]$ColumnIndex["notes"])
        $items.Add(('<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="{0}2:{0}{1}"><formula1>JobTrackerApplyNoteTemplates</formula1></dataValidation>' -f $column, $validationEndRow)) | Out-Null
    }
    if ($ColumnIndex.ContainsKey("applied_date")) {
        $column = ConvertTo-ExcelColumnName ([int]$ColumnIndex["applied_date"])
        $items.Add(('<dataValidation type="date" operator="between" allowBlank="1" showInputMessage="1" sqref="{0}2:{0}{1}"><formula1>DATE(2020,1,1)</formula1><formula2>DATE(2035,12,31)</formula2></dataValidation>' -f $column, $validationEndRow)) | Out-Null
    }
    if ($items.Count -eq 0) {
        return ""
    }

    return '<dataValidations count="{0}">{1}</dataValidations>' -f $items.Count, (($items.ToArray()) -join "")
}

function New-OpenXmlJobsSheet {
    param(
        [object[]]$Rows,
        [string[]]$Columns
    )

    $columnIndex = @{}
    for ($index = 0; $index -lt $Columns.Count; $index++) {
        $columnIndex[$Columns[$index]] = $index + 1
    }

    $rowCount = @($Rows).Count
    $lastDataRow = [Math]::Max(1, $rowCount + 1)
    $lastColumn = $Columns.Count
    $lastColumnName = ConvertTo-ExcelColumnName $lastColumn
    $dimensionRef = "A1:{0}{1}" -f $lastColumnName, $lastDataRow
    $sheetData = New-Object System.Text.StringBuilder
    $hyperlinkXml = New-Object System.Collections.Generic.List[string]
    $hyperlinkRels = New-Object System.Collections.Generic.List[string]
    $hyperlinkId = 1

    [void]$sheetData.Append('<row r="1" ht="24" customHeight="1">')
    for ($column = 1; $column -le $lastColumn; $column++) {
        [void]$sheetData.Append((New-OpenXmlCellXml -Row 1 -Column $column -Value (Get-ColumnLabel $Columns[$column - 1]) -StyleId 1))
    }
    [void]$sheetData.Append('</row>')

    for ($rowIndex = 0; $rowIndex -lt $rowCount; $rowIndex++) {
        $rowNumber = $rowIndex + 2
        $row = $Rows[$rowIndex]
        [void]$sheetData.Append(('<row r="{0}" ht="30" customHeight="1">' -f $rowNumber))
        foreach ($columnName in $Columns) {
            $columnNumber = [int]$columnIndex[$columnName]
            $value = Repair-DisplayText (Get-RowValue -Row $row -Name $columnName)
            if ($columnName -eq "review_priority") {
                $formula = Get-OpenXmlReviewPriorityFormula -Row $rowNumber -ColumnIndex $columnIndex
                [void]$sheetData.Append((New-OpenXmlCellXml -Row $rowNumber -Column $columnNumber -Value $value -StyleId 0 -Formula $formula))
            }
            elseif ($columnName -eq "job_url") {
                $url = Get-RowValue -Row $row -Name "job_url_raw"
                if ([string]::IsNullOrWhiteSpace($url) -and $value -match "^https?://") {
                    $url = $value
                }
                if ($url -match "^https?://") {
                    $relationshipId = "rId{0}" -f $hyperlinkId
                    $reference = Get-OpenXmlCellReference -Column $columnNumber -Row $rowNumber
                    $hyperlinkXml.Add(('<hyperlink ref="{0}" r:id="{1}"/>' -f $reference, $relationshipId)) | Out-Null
                    $hyperlinkRels.Add(('<Relationship Id="{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="{1}" TargetMode="External"/>' -f $relationshipId, (ConvertTo-OpenXmlEscapedText $url))) | Out-Null
                    $hyperlinkId++
                    [void]$sheetData.Append((New-OpenXmlCellXml -Row $rowNumber -Column $columnNumber -Value "Open" -StyleId 4 -Hyperlink))
                }
                else {
                    [void]$sheetData.Append((New-OpenXmlCellXml -Row $rowNumber -Column $columnNumber -Value $value))
                }
            }
            elseif (Test-OpenXmlNumericColumn $columnName) {
                [void]$sheetData.Append((New-OpenXmlCellXml -Row $rowNumber -Column $columnNumber -Value $value -StyleId 3 -Numeric))
            }
            elseif ($columnName -in @("job_title", "matched_keywords", "fit_notes", "job_url_raw", "notes", "duplicate_reason", "alternate_urls")) {
                [void]$sheetData.Append((New-OpenXmlCellXml -Row $rowNumber -Column $columnNumber -Value $value -StyleId 5))
            }
            else {
                [void]$sheetData.Append((New-OpenXmlCellXml -Row $rowNumber -Column $columnNumber -Value $value))
            }
        }
        [void]$sheetData.Append('</row>')
    }

    $hyperlinksSection = ""
    if ($hyperlinkXml.Count -gt 0) {
        $hyperlinksSection = "<hyperlinks>{0}</hyperlinks>" -f (($hyperlinkXml.ToArray()) -join "")
    }

    $relationshipsXml = ""
    if ($hyperlinkRels.Count -gt 0) {
        $relationshipsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{0}</Relationships>' -f (($hyperlinkRels.ToArray()) -join "")
    }

    $autoFilterRef = "A1:{0}{1}" -f $lastColumnName, $lastDataRow
    $conditionalXml = Get-OpenXmlStatusConditionalFormattingXml -LastColumn $lastColumn -LastDataRow $lastDataRow -ColumnIndex $columnIndex
    $validationXml = Get-OpenXmlDataValidationsXml -LastDataRow $lastDataRow -ColumnIndex $columnIndex
    $columnXml = Get-OpenXmlColumnXml -Columns $Columns -ColumnIndex $columnIndex

    $sheetXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="$dimensionRef"/>
  <sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="C2" sqref="C2"/></sheetView></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  $columnXml
  <sheetData>$($sheetData.ToString())</sheetData>
  <autoFilter ref="$autoFilterRef"/>
  $conditionalXml
  $validationXml
  $hyperlinksSection
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
"@

    return [PSCustomObject]@{
        Xml = $sheetXml
        RelationshipsXml = $relationshipsXml
    }
}

function New-OpenXmlKeyValueSheetXml {
    param(
        [string]$Title,
        [object[]]$Pairs
    )

    $rows = New-Object System.Text.StringBuilder
    [void]$rows.Append('<row r="1">')
    [void]$rows.Append((New-OpenXmlCellXml -Row 1 -Column 1 -Value $Title -StyleId 6))
    [void]$rows.Append('</row>')
    $rowNumber = 3
    foreach ($pair in @($Pairs)) {
        [void]$rows.Append(('<row r="{0}">' -f $rowNumber))
        [void]$rows.Append((New-OpenXmlCellXml -Row $rowNumber -Column 1 -Value ([string]$pair[0]) -StyleId 7))
        [void]$rows.Append((New-OpenXmlCellXml -Row $rowNumber -Column 2 -Value ([string]$pair[1]) -StyleId 5))
        [void]$rows.Append('</row>')
        $rowNumber++
    }
    $lastRow = [Math]::Max(1, $rowNumber - 1)

    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B$lastRow"/>
  <sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols><col min="1" max="1" width="34" customWidth="1"/><col min="2" max="2" width="90" customWidth="1"/></cols>
  <sheetData>$($rows.ToString())</sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
"@
}

function New-OpenXmlObjectListSheetXml {
    param(
        [string]$Title,
        [object[]]$Rows
    )

    $rowsArray = @($Rows)
    $sheetRows = New-Object System.Text.StringBuilder
    [void]$sheetRows.Append('<row r="1">')
    [void]$sheetRows.Append((New-OpenXmlCellXml -Row 1 -Column 1 -Value $Title -StyleId 6))
    [void]$sheetRows.Append('</row>')

    if ($rowsArray.Count -eq 0) {
        [void]$sheetRows.Append('<row r="3">')
        [void]$sheetRows.Append((New-OpenXmlCellXml -Row 3 -Column 1 -Value "No data"))
        [void]$sheetRows.Append('</row>')
        $lastColumn = 1
        $lastRow = 3
    }
    else {
        $properties = @($rowsArray[0].PSObject.Properties.Name)
        $lastColumn = $properties.Count
        [void]$sheetRows.Append('<row r="3">')
        for ($column = 0; $column -lt $properties.Count; $column++) {
            [void]$sheetRows.Append((New-OpenXmlCellXml -Row 3 -Column ($column + 1) -Value $properties[$column] -StyleId 1))
        }
        [void]$sheetRows.Append('</row>')

        for ($rowIndex = 0; $rowIndex -lt $rowsArray.Count; $rowIndex++) {
            $rowNumber = $rowIndex + 4
            [void]$sheetRows.Append(('<row r="{0}">' -f $rowNumber))
            for ($column = 0; $column -lt $properties.Count; $column++) {
                $value = $rowsArray[$rowIndex].PSObject.Properties[$properties[$column]].Value
                [void]$sheetRows.Append((New-OpenXmlCellXml -Row $rowNumber -Column ($column + 1) -Value ([string]$value) -StyleId 5))
            }
            [void]$sheetRows.Append('</row>')
        }
        $lastRow = $rowsArray.Count + 3
    }

    $lastColumnName = ConvertTo-ExcelColumnName $lastColumn
    $autoFilter = $(if ($rowsArray.Count -gt 0) { '<autoFilter ref="A3:{0}{1}"/>' -f $lastColumnName, $lastRow } else { "" })
    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:$lastColumnName$lastRow"/>
  <sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols><col min="1" max="$lastColumn" width="22" customWidth="1"/></cols>
  <sheetData>$($sheetRows.ToString())</sheetData>
  $autoFilter
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
"@
}

function New-OpenXmlValidationSheetXml {
    $statusOptions = @(Get-JobTrackerStatusOptions)
    $ignoreOptions = @(Get-JobTrackerIgnoreReasonOptions)
    $rowCount = [Math]::Max($statusOptions.Count, $ignoreOptions.Count)
    if ($rowCount -lt 1) {
        $rowCount = 1
    }

    $rows = New-Object System.Text.StringBuilder
    for ($index = 0; $index -lt $rowCount; $index++) {
        $rowNumber = $index + 1
        [void]$rows.Append(('<row r="{0}">' -f $rowNumber))
        if ($index -lt $statusOptions.Count) {
            [void]$rows.Append((New-OpenXmlCellXml -Row $rowNumber -Column 1 -Value ([string]$statusOptions[$index])))
        }
        if ($index -lt $ignoreOptions.Count) {
            [void]$rows.Append((New-OpenXmlCellXml -Row $rowNumber -Column 2 -Value ([string]$ignoreOptions[$index])))
        }
        [void]$rows.Append('</row>')
    }

    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B$rowCount"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>$($rows.ToString())</sheetData>
</worksheet>
"@
}

function Get-OpenXmlSourceQuerySummary {
    $names = @(
        @("LinkedIn", "LinkedInQueries"),
        @("HelloWork", "HelloWorkQueries"),
        @("APEC", "ApecQueries"),
        @("France Travail", "FranceTravailQueries"),
        @("Adzuna", "AdzunaQueries"),
        @("API fallback", "ApiSearchQueries")
    )
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($entry in $names) {
        $value = Get-OpenXmlScriptValue -Name $entry[1] -DefaultValue @()
        $parts.Add(("{0} {1}" -f $entry[0], @($value).Count)) | Out-Null
    }
    if ($parts.Count -eq 0) {
        return "Not loaded"
    }

    return ($parts.ToArray()) -join " | "
}

function Get-OpenXmlSummaryPairs {
    param(
        [object[]]$Rows,
        [string]$FullPath,
        [AllowNull()]$Summary
    )

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

    return @(
        @("Generated", [string](Get-OpenXmlScriptValue -Name "RunStamp" -DefaultValue (Get-Date -Format "yyyy-MM-dd_HH-mm-ss"))),
        @("Profile", (Get-SummaryValue -Summary $Summary -Name "Profile")),
        @("Crawl mode", [string](Get-OpenXmlScriptValue -Name "CrawlMode" -DefaultValue "")),
        @("Retention rule", ("Keep non-application jobs only when Published is on or after {0}." -f (Get-OpenXmlScriptValue -Name "CutoffDate" -DefaultValue ""))),
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
        @("Tracker", $FullPath),
        @("Workbook writer", (Get-SummaryValue -Summary $Summary -Name "WorkbookWriter")),
        @("Manual fields", "Status, Applied date, Apply notes with ignore_reason templates"),
        @("Reminder", "Close this workbook before launching the crawler.")
    )
}

function Get-OpenXmlSettingsPairs {
    param(
        [string]$FullPath,
        [AllowNull()]$Summary
    )

    $configRoot = ""
    $localOverrideSummary = "none"
    $sourceDefaultSummary = "No source metadata"
    $credentialSummary = "No credential config"
    if (Get-Variable -Name JobCrawlerConfig -Scope Script -ErrorAction SilentlyContinue) {
        $configRoot = [string]$script:JobCrawlerConfig.Root
        $localOverrides = @(Get-ConfigProperty -Object $script:JobCrawlerConfig -Name "LocalOverrides" -DefaultValue @())
        if ($localOverrides.Count -gt 0) {
            $localOverrideSummary = ($localOverrides | ForEach-Object { Split-Path -Leaf $_ }) -join " | "
        }
    }
    if (Get-Variable -Name JobCrawlerSourcesConfig -Scope Script -ErrorAction SilentlyContinue) {
        $sourceDefinitions = @(Get-JobCrawlerSourceDefinitions -SourcesConfig $script:JobCrawlerSourcesConfig)
        if ($sourceDefinitions.Count -gt 0) {
            $sourceDefaultSummary = ($sourceDefinitions | ForEach-Object {
                "{0}: {1}{2}" -f $_.Key,
                    $(if ($_.EnabledByDefault) { "on" } else { "off" }),
                    $(if ($_.RequiresCredential) { " (credentials)" } else { "" })
            }) -join " | "
        }
        $credentialStatuses = @(Get-JobCrawlerCredentialStatuses -SourcesConfig $script:JobCrawlerSourcesConfig)
        if ($credentialStatuses.Count -gt 0) {
            $credentialSummary = ($credentialStatuses | ForEach-Object { "{0}/{1}: {2}" -f $_.Source, $_.Credential, $_.Status }) -join " | "
        }
    }

    return @(
        @("Profile", (Get-SummaryValue -Summary $Summary -Name "Profile")),
        @("Crawl mode", [string](Get-OpenXmlScriptValue -Name "CrawlMode" -DefaultValue "")),
        @("Days back", [string](Get-OpenXmlScriptValue -Name "DaysBack" -DefaultValue "")),
        @("Location", [string](Get-OpenXmlScriptValue -Name "Location" -DefaultValue "")),
        @("Tracker", $FullPath),
        @("Config directory", $configRoot),
        @("Local config overrides", $localOverrideSummary),
        @("Source defaults", $sourceDefaultSummary),
        @("Cache directory", [string](Get-OpenXmlScriptValue -Name "CacheDirectory" -DefaultValue "")),
        @("Cache TTL hours", [string](Get-OpenXmlScriptValue -Name "CacheTtlHours" -DefaultValue "")),
        @("Dry run", (Get-SummaryValue -Summary $Summary -Name "DryRun")),
        @("Diagnostics mode", (Get-SummaryValue -Summary $Summary -Name "DiagnosticMode")),
        @("Diagnostic file", (Get-SummaryValue -Summary $Summary -Name "DiagnosticPath")),
        @("Crawl caps", (Get-SummaryValue -Summary $Summary -Name "CrawlCaps")),
        @("Run history", (Get-SummaryValue -Summary $Summary -Name "RunHistoryPath")),
        @("Credential status", $credentialSummary),
        @("Source queries", (Get-OpenXmlSourceQuerySummary)),
        @("Matching threshold", [string](Get-OpenXmlScriptValue -Name "MinimumMatchScore" -DefaultValue "")),
        @("Workbook writer", (Get-SummaryValue -Summary $Summary -Name "WorkbookWriter"))
    )
}

function Get-OpenXmlSourceHealthRows {
    if (Get-Variable -Name SourceRunStats -Scope Script -ErrorAction SilentlyContinue) {
        $sourceRows = @($script:SourceRunStats.ToArray())
        if ($sourceRows.Count -gt 0) {
            return $sourceRows
        }
    }

    return @([PSCustomObject]@{ Source = "No crawl source"; DurationSeconds = ""; SearchRequests = ""; DetailRequests = ""; CacheHits = ""; Candidates = ""; SelectedDetails = ""; SkippedOld = ""; SkippedContract = ""; SkippedNoMatch = ""; SkippedByCap = ""; Errors = ""; Matches = ""; Notes = "" })
}

function New-OpenXmlStylesXml {
    return @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>
  <fonts count="5">
    <font><sz val="10"/><color rgb="FF282F34"/><name val="Segoe UI"/></font>
    <font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Segoe UI"/></font>
    <font><u/><sz val="10"/><color rgb="FF2563EB"/><name val="Segoe UI"/></font>
    <font><b/><sz val="16"/><color rgb="FF263238"/><name val="Segoe UI"/></font>
    <font><b/><sz val="10"/><color rgb="FF282F34"/><name val="Segoe UI"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF263238"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color rgb="FFE2E8F0"/></left><right style="thin"><color rgb="FFE2E8F0"/></right><top style="thin"><color rgb="FFE2E8F0"/></top><bottom style="thin"><color rgb="FFE2E8F0"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="8">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
    <xf numFmtId="1" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
    <xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="4" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="16">
    <dxf/>
    <dxf/>
    <dxf/>
    <dxf/>
    <dxf/>
    <dxf/>
    <dxf/>
    <dxf><font><b/><color rgb="FF92400E"/></font></dxf>
    <dxf><font><b/><color rgb="FF227148"/></font></dxf>
    <dxf><font><b/><color rgb="FF2563EB"/></font></dxf>
    <dxf><font><b/><color rgb="FF227148"/></font></dxf>
    <dxf><font><color rgb="FF64748B"/></font></dxf>
    <dxf><font><color rgb="FFB91C1C"/></font></dxf>
    <dxf><font><color rgb="FF64748B"/></font></dxf>
    <dxf><font><b/><color rgb="FF92400E"/></font></dxf>
    <dxf><font><b/><color rgb="FF92400E"/></font></dxf>
  </dxfs>
  <tableStyles count="0" defaultTableStyle="TableStyleLight9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
'@
}

function Export-TrackerWorkbookWithOpenXml {
    param(
        [object[]]$Rows,
        [string]$Path,
        [AllowNull()]$Summary = $null
    )

    Initialize-OpenXmlZipAssemblies
    $fullPath = [IO.Path]::GetFullPath($Path)
    $directory = Split-Path -Parent $fullPath
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    $tempRoot = Join-Path ([IO.Path]::GetTempPath()) ("custom-job-tracker-xlsx-{0}" -f ([Guid]::NewGuid().ToString("N")))
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
    try {
        $sheetNames = @{
            Jobs            = ConvertTo-OpenXmlSheetName -Name (Get-WorkbookSheetName -Key "jobs" -DefaultValue "Jobs") -DefaultValue "Jobs"
            Summary         = ConvertTo-OpenXmlSheetName -Name (Get-WorkbookSheetName -Key "summary" -DefaultValue "Summary") -DefaultValue "Summary"
            Settings        = ConvertTo-OpenXmlSheetName -Name (Get-WorkbookSheetName -Key "settings" -DefaultValue "Settings") -DefaultValue "Settings"
            SourceHealth    = ConvertTo-OpenXmlSheetName -Name (Get-WorkbookSheetName -Key "source_health" -DefaultValue "Source Health") -DefaultValue "Source Health"
            FeedbackQuality = ConvertTo-OpenXmlSheetName -Name (Get-WorkbookSheetName -Key "feedback_quality" -DefaultValue "Feedback Quality") -DefaultValue "Feedback Quality"
            Validation      = "_validation"
        }
        $rowsArray = @($Rows)
        $columns = @(Get-JobTrackerMasterColumns)
        $summaryForWorkbook = $Summary
        if ($null -eq $summaryForWorkbook) {
            $summaryForWorkbook = @{}
        }
        try {
            $summaryForWorkbook["WorkbookWriter"] = "OpenXML no-Excel writer"
        }
        catch {
        }

        $jobsSheet = New-OpenXmlJobsSheet -Rows $rowsArray -Columns $columns
        $summaryPairs = Get-OpenXmlSummaryPairs -Rows $rowsArray -FullPath $fullPath -Summary $summaryForWorkbook
        $settingsPairs = Get-OpenXmlSettingsPairs -FullPath $fullPath -Summary $summaryForWorkbook
        $sourceRows = @(Get-OpenXmlSourceHealthRows)
        $feedbackRows = @(Get-FeedbackQualityRows -Rows $rowsArray)
        $statusRowCount = @(Get-JobTrackerStatusOptions).Count
        if ($statusRowCount -lt 1) { $statusRowCount = 1 }
        $validationRowCount = [Math]::Max($statusRowCount, @(Get-JobTrackerIgnoreReasonOptions).Count)
        if ($validationRowCount -lt 1) { $validationRowCount = 1 }
        $statusDefinedName = "'_validation'!" + '$A$1:$A$' + $statusRowCount
        $notesDefinedName = "'_validation'!" + '$B$1:$B$' + $validationRowCount

        Write-OpenXmlUtf8File -Path (Join-Path $tempRoot "[Content_Types].xml") -Content @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet4.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet5.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet6.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>
"@

        Write-OpenXmlUtf8File -Path (Join-Path $tempRoot "_rels\.rels") -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
'@

        $workbookXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <workbookPr/>
  <bookViews><workbookView activeTab="0"/></bookViews>
  <sheets>
    <sheet name="$(ConvertTo-OpenXmlEscapedText $sheetNames.Jobs)" sheetId="1" r:id="rId1"/>
    <sheet name="$(ConvertTo-OpenXmlEscapedText $sheetNames.Summary)" sheetId="2" r:id="rId2"/>
    <sheet name="$(ConvertTo-OpenXmlEscapedText $sheetNames.Settings)" sheetId="3" r:id="rId3"/>
    <sheet name="$(ConvertTo-OpenXmlEscapedText $sheetNames.SourceHealth)" sheetId="4" r:id="rId4"/>
    <sheet name="$(ConvertTo-OpenXmlEscapedText $sheetNames.FeedbackQuality)" sheetId="5" r:id="rId5"/>
    <sheet name="_validation" sheetId="6" state="hidden" r:id="rId6"/>
  </sheets>
  <definedNames>
    <definedName name="JobTrackerStatusOptions">$(ConvertTo-OpenXmlEscapedText $statusDefinedName)</definedName>
    <definedName name="JobTrackerApplyNoteTemplates">$(ConvertTo-OpenXmlEscapedText $notesDefinedName)</definedName>
  </definedNames>
  <calcPr calcId="191029" fullCalcOnLoad="1"/>
</workbook>
"@
        Write-OpenXmlUtf8File -Path (Join-Path $tempRoot "xl\workbook.xml") -Content $workbookXml
        Write-OpenXmlUtf8File -Path (Join-Path $tempRoot "xl\_rels\workbook.xml.rels") -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet4.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet5.xml"/>
  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet6.xml"/>
  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
'@
        Write-OpenXmlUtf8File -Path (Join-Path $tempRoot "xl\styles.xml") -Content (New-OpenXmlStylesXml)
        Write-OpenXmlUtf8File -Path (Join-Path $tempRoot "xl\worksheets\sheet1.xml") -Content $jobsSheet.Xml
        if (-not [string]::IsNullOrWhiteSpace($jobsSheet.RelationshipsXml)) {
            Write-OpenXmlUtf8File -Path (Join-Path $tempRoot "xl\worksheets\_rels\sheet1.xml.rels") -Content $jobsSheet.RelationshipsXml
        }
        Write-OpenXmlUtf8File -Path (Join-Path $tempRoot "xl\worksheets\sheet2.xml") -Content (New-OpenXmlKeyValueSheetXml -Title "Job Tracker" -Pairs $summaryPairs)
        Write-OpenXmlUtf8File -Path (Join-Path $tempRoot "xl\worksheets\sheet3.xml") -Content (New-OpenXmlKeyValueSheetXml -Title "Settings" -Pairs $settingsPairs)
        Write-OpenXmlUtf8File -Path (Join-Path $tempRoot "xl\worksheets\sheet4.xml") -Content (New-OpenXmlObjectListSheetXml -Title "Source Health" -Rows $sourceRows)
        Write-OpenXmlUtf8File -Path (Join-Path $tempRoot "xl\worksheets\sheet5.xml") -Content (New-OpenXmlObjectListSheetXml -Title "Feedback Quality" -Rows $feedbackRows)
        Write-OpenXmlUtf8File -Path (Join-Path $tempRoot "xl\worksheets\sheet6.xml") -Content (New-OpenXmlValidationSheetXml)

        $tempFile = Join-Path ([IO.Path]::GetTempPath()) ("custom-job-tracker-xlsx-{0}.xlsx" -f ([Guid]::NewGuid().ToString("N")))
        if (Test-Path -LiteralPath $tempFile) {
            Remove-Item -LiteralPath $tempFile -Force
        }
        Compress-OpenXmlDirectory -SourceDirectory $tempRoot -DestinationPath $tempFile
        Move-Item -LiteralPath $tempFile -Destination $fullPath -Force
        $script:LastWorkbookExportResult = [PSCustomObject]@{ Backend = "OpenXML"; Path = $fullPath; FallbackPath = ""; UsedFallback = $false }
    }
    finally {
        if (Test-Path -LiteralPath $tempRoot) {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-OpenXmlZipEntryText {
    param(
        [System.IO.Compression.ZipArchive]$Archive,
        [string]$Name
    )

    $entry = $null
    foreach ($candidate in @($Name.Replace("\", "/"), $Name.Replace("/", "\"), $Name)) {
        $entry = $Archive.GetEntry($candidate)
        if ($null -ne $entry) {
            break
        }
    }
    if ($null -eq $entry) {
        return ""
    }
    $stream = $entry.Open()
    try {
        $reader = New-Object System.IO.StreamReader($stream, [Text.Encoding]::UTF8)
        try {
            return $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }
}

function Get-OpenXmlCellColumnNumber {
    param([string]$Reference)

    $letters = ([regex]::Match([string]$Reference, "^[A-Z]+")).Value
    $number = 0
    foreach ($char in $letters.ToCharArray()) {
        $number = ($number * 26) + ([int][char]$char - [int][char]'A' + 1)
    }

    return $number
}

function Get-OpenXmlCellText {
    param(
        [xml]$Cell,
        [string[]]$SharedStrings
    )

    $cellNode = $Cell.SelectSingleNode("//*[local-name()='c']")
    if ($null -eq $cellNode) {
        return ""
    }

    $type = [string]$cellNode.GetAttribute("t")
    if ($type -eq "inlineStr") {
        return (($cellNode.SelectNodes(".//*[local-name()='is']//*[local-name()='t']") | ForEach-Object { [string]$_.InnerText }) -join "")
    }
    $valueNode = $cellNode.SelectSingleNode("./*[local-name()='v']")
    $value = $(if ($null -ne $valueNode) { [string]$valueNode.InnerText } else { "" })
    if ($type -eq "s") {
        $index = 0
        if ([int]::TryParse($value, [ref]$index) -and $index -ge 0 -and $index -lt $SharedStrings.Count) {
            return [string]$SharedStrings[$index]
        }
    }
    if ($type -eq "str") {
        return $value
    }

    return $value
}

function Import-TrackerRowsFromOpenXmlXlsx {
    param([string]$Path)

    Initialize-OpenXmlZipAssemblies
    $fullPath = (Resolve-Path -LiteralPath $Path).Path
    $archive = [System.IO.Compression.ZipFile]::OpenRead($fullPath)
    try {
        [xml]$workbook = Get-OpenXmlZipEntryText -Archive $archive -Name "xl/workbook.xml"
        [xml]$workbookRels = Get-OpenXmlZipEntryText -Archive $archive -Name "xl/_rels/workbook.xml.rels"
        $targetSheetName = Get-WorkbookSheetName -Key "jobs" -DefaultValue "Jobs"
        $sheetNode = $null
        foreach ($sheet in @($workbook.SelectNodes("//*[local-name()='sheet']"))) {
            if ([string]$sheet.name -eq $targetSheetName) {
                $sheetNode = $sheet
                break
            }
        }
        if ($null -eq $sheetNode) {
            $sheetNode = @($workbook.SelectNodes("//*[local-name()='sheet']")) | Select-Object -First 1
        }
        if ($null -eq $sheetNode) {
            return @()
        }

        $relationshipId = $sheetNode.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        $target = ""
        foreach ($rel in @($workbookRels.SelectNodes("//*[local-name()='Relationship']"))) {
            if ([string]$rel.Id -eq $relationshipId) {
                $target = [string]$rel.Target
                break
            }
        }
        if ([string]::IsNullOrWhiteSpace($target)) {
            return @()
        }
        if ($target -notmatch "^xl/") {
            $target = "xl/{0}" -f $target.TrimStart("/")
        }
        [xml]$sheetXml = Get-OpenXmlZipEntryText -Archive $archive -Name $target

        $sharedStrings = @()
        $sharedText = Get-OpenXmlZipEntryText -Archive $archive -Name "xl/sharedStrings.xml"
        if (-not [string]::IsNullOrWhiteSpace($sharedText)) {
            [xml]$sharedXml = $sharedText
            $sharedStrings = @($sharedXml.SelectNodes("//*[local-name()='si']") | ForEach-Object {
                ($_.SelectNodes(".//*[local-name()='t']") | ForEach-Object { [string]$_.InnerText }) -join ""
            })
        }

        $rowNodes = @($sheetXml.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']"))
        if ($rowNodes.Count -lt 2) {
            return @()
        }

        $headersByColumn = @{}
        foreach ($cell in @($rowNodes[0].SelectNodes("*[local-name()='c']"))) {
            $cellXml = [xml]$cell.OuterXml
            $columnNumber = Get-OpenXmlCellColumnNumber ([string]$cell.r)
            $header = Get-OpenXmlCellText -Cell $cellXml -SharedStrings $sharedStrings
            if (-not [string]::IsNullOrWhiteSpace($header)) {
                $headersByColumn[$columnNumber] = ConvertTo-CanonicalColumnName $header.Trim()
            }
        }

        $rows = New-Object System.Collections.Generic.List[object]
        foreach ($rowNode in @($rowNodes | Select-Object -Skip 1)) {
            $values = @{}
            $hasValue = $false
            foreach ($cell in @($rowNode.SelectNodes("*[local-name()='c']"))) {
                $cellXml = [xml]$cell.OuterXml
                $columnNumber = Get-OpenXmlCellColumnNumber ([string]$cell.r)
                if (-not $headersByColumn.ContainsKey($columnNumber)) {
                    continue
                }
                $name = [string]$headersByColumn[$columnNumber]
                $value = Repair-DisplayText (Get-OpenXmlCellText -Cell $cellXml -SharedStrings $sharedStrings)
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
        $archive.Dispose()
    }
}

function Export-TrackerHtmlReport {
    param(
        [object[]]$Rows,
        [string]$Path,
        [AllowNull()]$Summary = $null,
        [AllowNull()][string]$Reason = ""
    )

    $htmlPath = [IO.Path]::ChangeExtension([IO.Path]::GetFullPath($Path), ".html")
    $columns = @(Get-JobTrackerDailyReviewColumns)
    $builder = New-Object System.Text.StringBuilder
    [void]$builder.AppendLine("<!doctype html><html><head><meta charset=""utf-8""><title>Job Tracker</title>")
    [void]$builder.AppendLine("<style>body{font-family:Segoe UI,Arial,sans-serif;margin:24px;color:#283034}table{border-collapse:collapse;width:100%;font-size:13px}th{background:#263238;color:white;text-align:left;position:sticky;top:0}th,td{border:1px solid #e2e8f0;padding:6px 8px;vertical-align:top}tr.applied,tr.offer{background:#f0fdf4}tr.interview{background:#eff6ff}tr.interesting{background:#fff9eb}tr.ignored{background:#f8fafc;color:#64748b}tr.rejected{background:#fef2f2}tr.withdrawn{background:#f5f5f5;color:#64748b}.meta{margin-bottom:16px;color:#4d5b72}.warning{background:#fffbeb;border:1px solid #f59e0b;padding:10px;margin-bottom:16px}</style>")
    [void]$builder.AppendLine("</head><body><h1>Job Tracker</h1>")
    [void]$builder.AppendLine(("<div class=""meta"">Generated {0} | Rows {1}</div>" -f (ConvertTo-OpenXmlEscapedText (Get-OpenXmlScriptValue -Name "RunStamp" -DefaultValue (Get-Date -Format "yyyy-MM-dd_HH-mm-ss"))), @($Rows).Count))
    if (-not [string]::IsNullOrWhiteSpace($Reason)) {
        [void]$builder.AppendLine(("<div class=""warning"">XLSX export fallback: {0}</div>" -f (ConvertTo-OpenXmlEscapedText $Reason)))
    }
    [void]$builder.AppendLine("<table><thead><tr>")
    foreach ($column in $columns) {
        [void]$builder.AppendLine(("<th>{0}</th>" -f (ConvertTo-OpenXmlEscapedText (Get-ColumnLabel $column))))
    }
    [void]$builder.AppendLine("</tr></thead><tbody>")
    foreach ($row in @($Rows)) {
        $statusClass = ConvertTo-MatchText (Get-RowValue -Row $row -Name "status")
        [void]$builder.AppendLine(("<tr class=""{0}"">" -f (ConvertTo-OpenXmlEscapedText $statusClass)))
        foreach ($column in $columns) {
            $value = Get-RowValue -Row $row -Name $column
            if ($column -eq "job_url") {
                $url = Get-RowValue -Row $row -Name "job_url_raw"
                if ($url -match "^https?://") {
                    [void]$builder.AppendLine(("<td><a href=""{0}"">Open</a></td>" -f (ConvertTo-OpenXmlEscapedText $url)))
                }
                else {
                    [void]$builder.AppendLine(("<td>{0}</td>" -f (ConvertTo-OpenXmlEscapedText $value)))
                }
            }
            else {
                [void]$builder.AppendLine(("<td>{0}</td>" -f (ConvertTo-OpenXmlEscapedText $value)))
            }
        }
        [void]$builder.AppendLine("</tr>")
    }
    [void]$builder.AppendLine("</tbody></table></body></html>")
    Write-OpenXmlUtf8File -Path $htmlPath -Content $builder.ToString()
    $script:LastWorkbookExportResult = [PSCustomObject]@{ Backend = "HTML"; Path = ""; FallbackPath = $htmlPath; UsedFallback = $true }
    return $htmlPath
}
