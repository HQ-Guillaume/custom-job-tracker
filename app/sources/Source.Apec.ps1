function Get-ApecContractType {
    param([AllowNull()]$Job)

    $rawType = [string](Get-ObjectPropertyValue -Object $Job -Names @("typeContrat", "idNomTypeContrat"))
    switch ($rawType) {
        "101888" { return "CDI" }
        "101887" { return "CDD" }
        "597171" { return "Internship" }
        "20053" { return "Apprenticeship" }
        "101930" { return "Interim" }
        "101889" { return "Interim" }
    }

    $contractText = Join-CleanTextParts @(
        (Get-ObjectPropertyValue -Object $Job -Names @("intitule", "title"))
        (Get-ObjectPropertyValue -Object $Job -Names @("texteOffre", "description"))
        $rawType
    )
    return Get-ContractTypeFromText -Text $contractText
}

function Get-ApecJobUrl {
    param([AllowNull()]$Job)

    $numeroOffre = [string](Get-ObjectPropertyValue -Object $Job -Names @("numeroOffre", "NumeroOffre"))
    if ([string]::IsNullOrWhiteSpace($numeroOffre)) {
        $id = [string](Get-ObjectPropertyValue -Object $Job -Names @("id", "Id"))
        if (-not [string]::IsNullOrWhiteSpace($id)) {
            $numeroOffre = "{0}W" -f $id
        }
    }

    if ([string]::IsNullOrWhiteSpace($numeroOffre)) {
        return ""
    }

    $template = [string](Get-ConfigPathValue -Object $script:JobCrawlerSourcesConfig -Path "endpoints.apec_detail" -DefaultValue "https://www.apec.fr/candidat/recherche-emploi.html/emploi/detail-offre/{id}")
    return $template.Replace("{id}", [Uri]::EscapeDataString($numeroOffre.Trim()))
}

function New-ApecSearchBody {
    param(
        [string]$Query,
        [int]$Page,
        [int]$PageSize,
        [string]$SortType = "SCORE"
    )

    return [ordered]@{
        lieux                   = @()
        fonctions               = @()
        statutPoste             = @()
        typesContrat            = @()
        typesConvention         = @("143684", "143685", "143686", "143687", "143706")
        niveauxExperience       = @()
        idsEtablissement        = @()
        secteursActivite        = @()
        typesTeletravail        = @()
        idNomZonesDeplacement   = @()
        positionNumbersExcluded = @()
        typeClient              = "CADRE"
        sorts                   = @(@{ type = $SortType; direction = "DESCENDING" })
        pagination              = @{ range = $PageSize; startIndex = ($Page * $PageSize) }
        activeFiltre            = $true
        pointGeolocDeReference  = @{ distance = 0 }
        motsCles                = $Query
    }
}

function Get-ApecJobs {
    Set-RunWindowTitle "Custom Job Tracker - APEC"
    Write-RunStatus "Collecting APEC jobs from the public search endpoint..."
    Write-RunStatus ("APEC plan: {0} query/queries, up to {1} page(s) each, no detail-page crawl." -f $ApecQueries.Count, $MaxApecPages)

    $stats = Start-SourceStats "APEC"
    $results = New-Object System.Collections.Generic.List[object]
    $headers = @{
        "Accept"  = "application/json, text/plain, */*"
        "Origin"  = [string](Get-ConfigPathValue -Object $script:JobCrawlerSourcesConfig -Path "endpoints.apec_origin" -DefaultValue "https://www.apec.fr")
        "Referer" = [string](Get-ConfigPathValue -Object $script:JobCrawlerSourcesConfig -Path "endpoints.apec_referer" -DefaultValue "https://www.apec.fr/candidat/recherche-emploi.html/emploi")
    }
    $searchUrl = [string](Get-ConfigPathValue -Object $script:JobCrawlerSourcesConfig -Path "endpoints.apec_search" -DefaultValue "https://www.apec.fr/cms/webservices/rechercheOffre")
    $pageSize = 20
    $queryIndex = 0

    foreach ($query in $ApecQueries) {
        $queryIndex++
        Write-RunStatus ("APEC query {0}/{1}: {2}" -f $queryIndex, $ApecQueries.Count, $query)

        for ($page = 0; $page -lt $MaxApecPages; $page++) {
            $body = New-ApecSearchBody -Query $query -Page $page -PageSize $pageSize -SortType "SCORE"
            try {
                Add-SourceMetric -Stats $stats -Name "SearchRequests"
                $response = Invoke-JsonPostRequest -Url $searchUrl -Body $body -Headers $headers -TimeoutSec 30
            }
            catch {
                Add-SourceMetric -Stats $stats -Name "Errors"
                Write-Warning ("APEC search failed for '{0}' page {1}: {2}" -f $query, ($page + 1), $_.Exception.Message)
                break
            }

            $jobArray = @()
            if ($null -ne $response -and @($response.PSObject.Properties.Name) -contains "resultats") {
                $jobArray = @($response.resultats)
            }
            elseif ($null -ne $response -and @($response.PSObject.Properties.Name) -contains "results") {
                $jobArray = @($response.results)
            }

            if ($jobArray.Count -eq 0) {
                break
            }

            foreach ($job in $jobArray) {
                Add-SourceMetric -Stats $stats -Name "Candidates"
                $publishedAt = ConvertTo-DateTimeOffsetOrNull (Get-ObjectPropertyValue -Object $job -Names @("datePublication", "dateValidation", "published_at"))
                if (-not (Test-IsRecent $publishedAt)) {
                    Add-SourceMetric -Stats $stats -Name "SkippedOld"
                    continue
                }

                $title = Repair-DisplayText ([string](Get-ObjectPropertyValue -Object $job -Names @("intitule", "title")))
                $companyName = Repair-DisplayText ([string](Get-ObjectPropertyValue -Object $job -Names @("nomCommercial", "company", "companyName")))
                $jobLocation = Repair-DisplayText ([string](Get-ObjectPropertyValue -Object $job -Names @("lieuTexte", "location")))
                $contractType = Get-ApecContractType $job
                $jobUrl = Get-ApecJobUrl $job
                $sourceText = Join-CleanTextParts @(
                    $title,
                    $companyName,
                    $jobLocation,
                    $contractType,
                    (Get-ObjectPropertyValue -Object $job -Names @("texteOffre", "description", "intituleSurbrillance"))
                )
                if (Test-ShouldSkipEarlyByContract -ContractType $contractType -Text $sourceText -Reliable) {
                    Add-SourceMetric -Stats $stats -Name "SkippedContract"
                    continue
                }

                $match = Get-JobMatch -Title $title -Text $sourceText
                if (-not $match.IsMatch) {
                    Add-SourceMetric -Stats $stats -Name "SkippedNoMatch"
                    continue
                }

                $result = New-JobResult -Title $title -CompanyName $companyName -JobLocation $jobLocation -ContractType $contractType -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url $jobUrl -Platform "APEC" -PublishedAt $publishedAt -SourceText $sourceText
                if ($null -ne $result) {
                    $results.Add($result) | Out-Null
                    Add-SourceMetric -Stats $stats -Name "Matches"
                }
            }

            Write-CountProgress -Activity ("APEC query {0}/{1}" -f $queryIndex, $ApecQueries.Count) -Current ($page + 1) -Total $MaxApecPages -Found $results.Count -Every 1
            if ($jobArray.Count -lt $pageSize) {
                break
            }
            Start-Sleep -Milliseconds $ApecDelayMilliseconds
        }
    }

    Write-RunStatus ("APEC complete: {0} matching jobs." -f $results.Count)
    Complete-SourceStats $stats
    return $results.ToArray()
}

