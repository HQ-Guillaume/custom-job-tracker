function ConvertTo-JobPipelineCanonicalRow {
    param([AllowNull()]$Row)

    if ($null -eq $Row) {
        return $null
    }

    if (-not (Get-Command New-OrderedJobRecord -ErrorAction SilentlyContinue)) {
        return $Row
    }

    $columns = @()
    if (Get-Variable -Name MasterColumns -Scope Script -ErrorAction SilentlyContinue) {
        $columns = @($script:MasterColumns)
    }
    if ($columns.Count -eq 0 -and (Get-Command Get-JobTrackerMasterColumns -ErrorAction SilentlyContinue)) {
        $columns = @(Get-JobTrackerMasterColumns)
    }
    if ($columns.Count -eq 0) {
        return $Row
    }

    $values = @{}
    foreach ($column in $columns) {
        $values[$column] = Get-RowValue -Row $Row -Name $column
    }

    return New-OrderedJobRecord $values
}

function New-JobPipelineDecision {
    param(
        [AllowNull()]$Row,
        [string]$Stage,
        [bool]$IsEligible,
        [string]$Reason,
        [string]$Rule,
        [bool]$KeepForever = $false
    )

    return [PSCustomObject]@{
        Row         = $Row
        Stage       = $Stage
        IsEligible  = $IsEligible
        Reason      = $Reason
        Rule        = $Rule
        KeepForever = $KeepForever
    }
}

function Test-JobPipelineExcludedContract {
    param(
        [AllowNull()]$Row,
        [AllowNull()]$CurrentRow = $null,
        [AllowNull()]$ExistingRow = $null
    )

    $rowContract = Get-RowValue -Row $Row -Name "contract_type"
    if (Test-IsExcludedContractType $rowContract) {
        return $true
    }

    $currentContract = Get-RowValue -Row $CurrentRow -Name "contract_type"
    $existingContract = Get-RowValue -Row $ExistingRow -Name "contract_type"
    if ([string]::IsNullOrWhiteSpace($currentContract) -and (Test-IsExcludedContractType $existingContract)) {
        return $true
    }

    return $false
}

function Test-JobPipelineInvalidLocation {
    param([AllowNull()]$Row)

    $invalidExistingCheck = Get-Command Test-IsInvalidExistingTrackerRow -ErrorAction SilentlyContinue
    if ($null -eq $invalidExistingCheck) {
        return $false
    }

    return Test-IsInvalidExistingTrackerRow $Row
}

function Get-JobPipelineEligibility {
    param(
        [AllowNull()]$Row,
        [AllowNull()]$CurrentRow = $null,
        [AllowNull()]$ExistingRow = $null,
        [string]$Stage = "pipeline"
    )

    if ($null -eq $Row) {
        return New-JobPipelineDecision -Row $Row -Stage $Stage -IsEligible:$false -Reason "null_row" -Rule "row_required"
    }

    $status = Get-RowValue -Row $Row -Name "status"
    $keepForever = Test-IsKeepForeverStatus $status
    if ($keepForever) {
        return New-JobPipelineDecision -Row $Row -Stage $Stage -IsEligible:$true -Reason "kept_by_status" -Rule "application_history" -KeepForever:$true
    }

    if (Test-JobPipelineInvalidLocation -Row $Row) {
        return New-JobPipelineDecision -Row $Row -Stage $Stage -IsEligible:$false -Reason "invalid_location" -Rule "location"
    }

    $publishedDate = ConvertTo-DateOrNull (Get-RowValue -Row $Row -Name "published_date")
    if ($null -eq $publishedDate) {
        return New-JobPipelineDecision -Row $Row -Stage $Stage -IsEligible:$false -Reason "missing_published_date" -Rule "published_date"
    }

    if (-not (Test-IsRecentTrackerRow $Row)) {
        return New-JobPipelineDecision -Row $Row -Stage $Stage -IsEligible:$false -Reason "outside_published_window" -Rule "published_date"
    }

    if (Test-JobPipelineExcludedContract -Row $Row -CurrentRow $CurrentRow -ExistingRow $ExistingRow) {
        return New-JobPipelineDecision -Row $Row -Stage $Stage -IsEligible:$false -Reason "excluded_contract" -Rule "contract"
    }

    return New-JobPipelineDecision -Row $Row -Stage $Stage -IsEligible:$true -Reason "eligible" -Rule "eligibility"
}

function Invoke-JobPipelineEligibilityGate {
    param(
        [object[]]$Rows,
        [string]$Stage = "pipeline"
    )

    $keptRows = New-Object System.Collections.Generic.List[object]
    $excludedRows = New-Object System.Collections.Generic.List[object]
    $decisions = New-Object System.Collections.Generic.List[object]
    $countByReason = @{}

    foreach ($row in @($Rows)) {
        $canonicalRow = ConvertTo-JobPipelineCanonicalRow $row
        $decision = Get-JobPipelineEligibility -Row $canonicalRow -Stage $Stage
        $decisions.Add($decision) | Out-Null
        if ($decision.IsEligible) {
            $keptRows.Add($canonicalRow) | Out-Null
        }
        else {
            $excludedRows.Add($canonicalRow) | Out-Null
            if (-not $countByReason.ContainsKey($decision.Reason)) {
                $countByReason[$decision.Reason] = 0
            }
            $countByReason[$decision.Reason] = [int]$countByReason[$decision.Reason] + 1
        }
    }

    return [PSCustomObject]@{
        KeptRows      = @($keptRows.ToArray())
        ExcludedRows  = @($excludedRows.ToArray())
        Decisions     = @($decisions.ToArray())
        CountByReason = $countByReason
        KeptCount     = [int]$keptRows.Count
        ExcludedCount = [int]$excludedRows.Count
    }
}

function Get-JobPipelineReasonCount {
    param(
        [AllowNull()]$GateResult,
        [string[]]$Reason
    )

    if ($null -eq $GateResult -or $null -eq $GateResult.CountByReason) {
        return 0
    }

    $total = 0
    foreach ($reasonName in @($Reason)) {
        if ($GateResult.CountByReason.ContainsKey($reasonName)) {
            $total += [int]$GateResult.CountByReason[$reasonName]
        }
    }

    return $total
}

function Format-JobPipelineReasonSummary {
    param([AllowNull()]$GateResult)

    if ($null -eq $GateResult -or $null -eq $GateResult.CountByReason -or $GateResult.CountByReason.Keys.Count -eq 0) {
        return "none"
    }

    return (@($GateResult.CountByReason.Keys | Sort-Object | ForEach-Object {
        "{0}: {1}" -f $_, $GateResult.CountByReason[$_]
    }) -join " | ")
}

function Test-JobPipelineInvariants {
    param(
        [object[]]$Rows,
        [string]$Stage = "pre_export",
        [switch]$ThrowOnIssue
    )

    $issues = New-Object System.Collections.Generic.List[object]
    foreach ($row in @($Rows)) {
        $decision = Get-JobPipelineEligibility -Row $row -Stage $Stage
        if ($decision.IsEligible) {
            continue
        }

        $issues.Add([PSCustomObject]@{
            Stage         = $Stage
            Reason        = $decision.Reason
            Rule          = $decision.Rule
            JobTitle      = Get-RowValue -Row $row -Name "job_title"
            CompanyName   = Get-RowValue -Row $row -Name "company_name"
            Platform      = Get-RowValue -Row $row -Name "platform"
            PublishedDate = Get-RowValue -Row $row -Name "published_date"
            ContractType  = Get-RowValue -Row $row -Name "contract_type"
            Status        = Get-RowValue -Row $row -Name "status"
            JobId         = Get-RowValue -Row $row -Name "job_id"
        }) | Out-Null
    }

    $result = [PSCustomObject]@{
        IsValid = $issues.Count -eq 0
        Issues  = @($issues.ToArray())
    }

    if ($ThrowOnIssue -and -not $result.IsValid) {
        $preview = @($result.Issues | Select-Object -First 5 | ForEach-Object {
            "{0} / {1} / {2} / {3}" -f $_.Reason, $_.JobTitle, $_.CompanyName, $_.Platform
        }) -join "; "
        throw ("Pipeline invariant check failed at {0}: {1} issue(s). {2}" -f $Stage, @($result.Issues).Count, $preview)
    }

    return $result
}
