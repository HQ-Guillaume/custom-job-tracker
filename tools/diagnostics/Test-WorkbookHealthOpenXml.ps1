[CmdletBinding()]
param(
    [string]$TrackerPath = "",
    [switch]$WarnOnly
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
. (Join-Path $projectRoot "app\core\JobTracker.Common.ps1")

if ([string]::IsNullOrWhiteSpace($TrackerPath)) {
    $TrackerPath = Join-Path $projectRoot "output\jobs_tracker.xlsx"
}
if (-not (Test-Path -LiteralPath $TrackerPath)) {
    throw "Tracker workbook not found: $TrackerPath"
}
if ([IO.Path]::GetExtension($TrackerPath).ToLowerInvariant() -ne ".xlsx") {
    throw "OpenXML workbook health check supports only XLSX files."
}

Add-Type -AssemblyName System.IO.Compression.FileSystem

$warnings = New-Object System.Collections.Generic.List[string]
$errors = New-Object System.Collections.Generic.List[string]

function Add-OpenXmlHealthWarning {
    param([string]$Message)
    $warnings.Add($Message) | Out-Null
}

function Add-OpenXmlHealthError {
    param([string]$Message)
    $errors.Add($Message) | Out-Null
}

function Get-OpenXmlText {
    param([AllowNull()][xml]$Xml, [string]$XPath, [AllowNull()]$NamespaceManager = $null)

    if ($null -eq $Xml) {
        return ""
    }

    $node = if ($null -ne $NamespaceManager) { $Xml.SelectSingleNode($XPath, $NamespaceManager) } else { $Xml.SelectSingleNode($XPath) }
    if ($null -eq $node) {
        return ""
    }

    return [string]$node.InnerText
}

function Get-OpenXmlColumnIndex {
    param([string]$CellReference)

    $letters = ([regex]::Match([string]$CellReference, "^[A-Z]+")).Value
    if ([string]::IsNullOrWhiteSpace($letters)) {
        return 0
    }

    $index = 0
    foreach ($char in $letters.ToCharArray()) {
        $index = ($index * 26) + ([int][char]::ToUpperInvariant($char) - [int][char]'A' + 1)
    }

    return $index
}

function Get-OpenXmlCellText {
    param(
        [AllowNull()]$Cell,
        [string[]]$SharedStrings
    )

    if ($null -eq $Cell) {
        return ""
    }

    $type = [string]$Cell.t
    if ($type -eq "inlineStr") {
        return [string]$Cell.is.t
    }

    $value = [string]$Cell.v
    if ($type -eq "s") {
        $index = 0
        if ([int]::TryParse($value, [ref]$index) -and $index -ge 0 -and $index -lt $SharedStrings.Count) {
            return [string]$SharedStrings[$index]
        }
    }

    return $value
}

function Get-OpenXmlWorkbookModel {
    param([string]$Path)

    $tempRoot = Join-Path ([IO.Path]::GetTempPath()) ("custom-job-tracker-health-{0}" -f ([Guid]::NewGuid().ToString("N")))
    [IO.Compression.ZipFile]::ExtractToDirectory($Path, $tempRoot)

    try {
        [xml]$workbook = Get-Content -LiteralPath (Join-Path $tempRoot "xl\workbook.xml") -Raw -Encoding UTF8
        [xml]$rels = Get-Content -LiteralPath (Join-Path $tempRoot "xl\_rels\workbook.xml.rels") -Raw -Encoding UTF8

        $sharedStrings = @()
        $sharedStringsPath = Join-Path $tempRoot "xl\sharedStrings.xml"
        if (Test-Path -LiteralPath $sharedStringsPath) {
            [xml]$sharedXml = Get-Content -LiteralPath $sharedStringsPath -Raw -Encoding UTF8
            $sharedStrings = @($sharedXml.sst.si | ForEach-Object {
                if ($null -ne $_.t) {
                    [string]$_.t
                }
                else {
                    [string](@($_.r | ForEach-Object { $_.t }) -join "")
                }
            })
        }

        $relationshipTargets = @{}
        foreach ($relationship in @($rels.Relationships.Relationship)) {
            $relationshipTargets[[string]$relationship.Id] = [string]$relationship.Target
        }

        $sheets = New-Object System.Collections.Generic.List[object]
        foreach ($sheet in @($workbook.workbook.sheets.sheet)) {
            $sheetName = [string]$sheet.name
            $relationshipId = [string]$sheet.id
            if ([string]::IsNullOrWhiteSpace($relationshipId)) {
                $relationshipId = [string]$sheet.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
            }

            $target = ""
            if ($relationshipTargets.ContainsKey($relationshipId)) {
                $target = $relationshipTargets[$relationshipId]
            }
            if ($target -notmatch "^xl/") {
                $target = Join-Path "xl" $target
            }
            $targetPath = Join-Path $tempRoot $target

            [xml]$sheetXml = Get-Content -LiteralPath $targetPath -Raw -Encoding UTF8
            $sheets.Add([PSCustomObject]@{
                Name = $sheetName
                Path = $targetPath
                Xml = $sheetXml
            }) | Out-Null
        }

        return [PSCustomObject]@{
            TempRoot = $tempRoot
            Sheets = @($sheets.ToArray())
            SharedStrings = @($sharedStrings)
        }
    }
    catch {
        Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        throw
    }
}

$model = Get-OpenXmlWorkbookModel -Path (Resolve-Path -LiteralPath $TrackerPath).Path
try {
    $sheetNames = @($model.Sheets | ForEach-Object { $_.Name })
    foreach ($requiredSheet in @("Jobs", "Summary")) {
        if ($requiredSheet -notin $sheetNames) {
            Add-OpenXmlHealthError "Missing workbook sheet: $requiredSheet"
        }
    }
    foreach ($optionalSheet in @("Settings", "Source Health", "Feedback Quality")) {
        if ($optionalSheet -notin $sheetNames) {
            Add-OpenXmlHealthWarning "Missing workbook sheet: $optionalSheet"
        }
    }

    $jobsSheet = @($model.Sheets | Where-Object { $_.Name -eq "Jobs" } | Select-Object -First 1)
    if ($jobsSheet.Count -gt 0) {
        $rows = @($jobsSheet[0].Xml.worksheet.sheetData.row)
        if ($rows.Count -eq 0) {
            Add-OpenXmlHealthError "Jobs sheet has no rows."
        }
        else {
            $headerRow = $rows[0]
            $headersByIndex = @{}
            foreach ($cell in @($headerRow.c)) {
                $columnIndex = Get-OpenXmlColumnIndex -CellReference ([string]$cell.r)
                if ($columnIndex -gt 0) {
                    $headersByIndex[$columnIndex] = Get-OpenXmlCellText -Cell $cell -SharedStrings $model.SharedStrings
                }
            }

            $expectedLabels = Get-JobTrackerMasterColumns | ForEach-Object { Get-ColumnLabel $_ }
            foreach ($label in @($expectedLabels)) {
                if ($label -notin @($headersByIndex.Values)) {
                    Add-OpenXmlHealthError "Missing workbook column: $label"
                }
            }

            $linkColumn = ($headersByIndex.GetEnumerator() | Where-Object { $_.Value -eq "Link" } | Select-Object -First 1).Key
            if ($null -ne $linkColumn -and $rows.Count -gt 1) {
                $hyperlinkRefs = @{}
                foreach ($hyperlink in @($jobsSheet[0].Xml.worksheet.hyperlinks.hyperlink)) {
                    $ref = [string]$hyperlink.ref
                    if (-not [string]::IsNullOrWhiteSpace($ref)) {
                        $hyperlinkRefs[$ref] = $true
                    }
                }
                $missingLinkFormulas = 0
                foreach ($row in @($rows | Select-Object -Skip 1)) {
                    $cell = @($row.c | Where-Object { (Get-OpenXmlColumnIndex -CellReference ([string]$_.r)) -eq [int]$linkColumn } | Select-Object -First 1)
                    if ($cell.Count -gt 0) {
                        $cellReference = [string]$cell[0].r
                        $formulaText = ""
                        $formulaNodes = @($cell[0].GetElementsByTagName("f"))
                        if ($formulaNodes.Count -gt 0) {
                            $formulaText = [string]$formulaNodes[0].InnerText
                        }
                        if (-not $hyperlinkRefs.ContainsKey($cellReference) -and $formulaText -notmatch '^HYPERLINK\(') {
                            $missingLinkFormulas++
                        }
                    }
                }
                if ($missingLinkFormulas -gt 0) {
                    Add-OpenXmlHealthWarning "Rows without clickable Link formulas: $missingLinkFormulas"
                }
            }

            Write-Host ("Tracker: {0}" -f (Resolve-Path -LiteralPath $TrackerPath).Path)
            Write-Host ("Sheets: {0}" -f ($sheetNames -join " | "))
            Write-Host ("Jobs rows including header: {0}" -f $rows.Count)
            Write-Host ("Jobs columns: {0}" -f (@($headersByIndex.Values) -join " | "))
        }
    }
}
finally {
    if ($null -ne $model -and -not [string]::IsNullOrWhiteSpace([string]$model.TempRoot)) {
        Remove-Item -LiteralPath $model.TempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

if ($warnings.Count -gt 0) {
    Write-Host ""
    Write-Host "Warnings:"
    foreach ($warning in @($warnings)) {
        Write-Host ("- {0}" -f $warning)
    }
}

if ($errors.Count -gt 0) {
    Write-Host ""
    Write-Host "Errors:"
    foreach ($item in @($errors)) {
        Write-Host ("- {0}" -f $item)
    }
    if (-not $WarnOnly) {
        exit 1
    }
}

if ($errors.Count -eq 0) {
    Write-Host ""
    Write-Host "OpenXML workbook health check passed."
}
