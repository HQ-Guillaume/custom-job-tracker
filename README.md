# Analytics Job Tracker

Rolling crawler and lightweight application tracker for Web/Digital Analytics jobs.

## Main File

Use this single workbook:

```text
output\jobs_tracker.xlsx
```

The crawler creates and updates the same workbook each time. The `output` folder is ignored by Git so your personal tracker, notes, and backups stay local.

It keeps:

- matching jobs whose `Published` date is within the last 7 days
- older jobs only when their `status` is application-related: `applied`, `interview`, `offer`, `rejected`, or `withdrawn`

CDD, apprenticeship, and internship jobs are excluded from new crawl results.

## Workbook

`Jobs` is the main sheet. It is designed for daily review: filtered table, frozen decision columns, clickable job links, compact row height, hidden technical columns, a status dropdown, gentle full-row status colors, and restrained text cues for priority/match.

Main review columns:

- `Priority`
- `Status`: `new`, `interesting`, `ignored`, `applied`, `interview`, `offer`, `rejected`, `withdrawn`
- `Job title`
- `Company`
- `Employer type`: `annonceur`, `agency`, `consulting`, `esn`, or `unknown`
- `City / region`
- `Contract`
- `Sources`
- `Published`
- `Age`
- `Link`
- `Applied date`
- `Apply notes`
- `Match`
- `Why it matched`

Manual fields are `Status`, `Applied date`, and `Apply notes`. `Apply notes` has dropdown templates for ignored-job feedback. Backend fields such as `Score`, `Role score`, `Employer fit`, `Location fit`, `Seniority fit`, `Contract fit`, `Fit notes`, `Seen now`, `First seen`, `Last seen`, `New?`, `Duplicate / retention note`, `Job ID`, `Raw URL`, `Other URLs`, and `Source count` are hidden but kept for ranking, deduplication, status updates, and retention.

`Summary` contains the latest crawl report, including the published-date retention rule, match levels, employer-type distribution, fit demotions, and backup path.

Close `jobs_tracker.xlsx` before launching the crawler so Excel does not lock the file.

## Launch

Double-click:

```text
Run-AnalyticsJobCrawler.cmd
```

or run:

```powershell
cd "C:\Users\Guillaume\Documents\Codex\job-crawler"
powershell -ExecutionPolicy Bypass -File .\Find-AnalyticsJobs.ps1
```

## Status Updates

You can edit the workbook directly, or use:

```powershell
powershell -ExecutionPolicy Bypass -File .\Update-JobStatus.ps1 -JobId "JOB_ID" -Status applied -Notes "Applied on LinkedIn"
```

When status is set to `applied`, `applied_date` is filled with today by default unless you provide one:

```powershell
powershell -ExecutionPolicy Bypass -File .\Update-JobStatus.ps1 -JobId "JOB_ID" -Status applied -AppliedDate 2026-06-13
```

When you set `Status` to `ignored`, fill `Apply notes` with one of the `ignore_reason=...` templates. Blank notes on ignored rows are highlighted in the workbook.

Useful ignored reasons:

- `ignore_reason=not_analytics_enough; detail=`
- `ignore_reason=too_seo_sea_marketing; detail=`
- `ignore_reason=too_data_analyst; detail=`
- `ignore_reason=too_data_engineering; detail=`
- `ignore_reason=too_bi_reporting; detail=`
- `ignore_reason=too_crm_emailing; detail=`
- `ignore_reason=too_content_social; detail=`
- `ignore_reason=agency_consulting_esn; detail=`
- `ignore_reason=wrong_seniority; detail=`
- `ignore_reason=wrong_location; detail=`
- `ignore_reason=company_not_interested; detail=`
- `ignore_reason=duplicate; detail=`
- `ignore_reason=other; detail=`

## Outputs

The crawler writes only the main workbook:

- `output\jobs_tracker.xlsx`: source of truth
- `output\backups\*.xlsx`: recent automatic workbook backups before tracker/status updates

These files are local personal data and are not committed to the GitHub repository.

To re-apply workbook formatting without crawling:

```powershell
powershell -ExecutionPolicy Bypass -File .\Export-JobTrackerXlsx.ps1
```

This requires desktop Microsoft Excel.

To check workbook health without crawling:

```powershell
powershell -ExecutionPolicy Bypass -File .\Test-JobTrackerHealth.ps1
```

The health check opens the workbook read-only and verifies the expected sheets, columns, hidden backend fields, clickable links, status values, duplicate job IDs, and status row formatting.

## Matching And Ranking

Main role-matching signals:

- title signals: Web Analyst, Digital Analyst, Digital Analytics Consultant, Tracking, Web Analytics, digital performance, CRO, plus Data Analyst when another relevant signal is present
- description/tool signals: Google Tag Manager, GTM, Google Analytics, GA4, Piano Analytics, ContentSquare, dataLayer, tagging plan, consent mode, A/B testing, dashboards, KPIs
- ranking: `High` >= 80, `Medium` >= 50, `Review` >= 35
- jobs with only description/tool matches and no analytics-related title are kept but capped at `Review`

The final `Match` uses several dimensions:

- `Role score`: web/digital analytics relevance from title and description
- `Employer fit`: annonceur is favored; agency, consulting, and ESN are demoted but not excluded
- `Location fit`: Paris/Ile-de-France/France/remote signals are favored; foreign locations are demoted
- `Seniority fit`: internship/junior/managerial roles are demoted
- `Contract fit`: CDI/permanent/full-time is favored; freelance is slightly demoted; CDD/apprenticeship/internship are excluded before export

The tracker also uses your history:

- similar jobs to `applied`, `interview`, `offer`, or `interesting` can receive a small score boost
- similar jobs to `ignored` can receive a score penalty
- ignored jobs with structured `ignore_reason=...` notes teach the crawler more precisely: SEO/SEA rejects affect marketing roles, data-engineering rejects affect dbt/Snowflake/pipeline roles, and `duplicate` does not reduce relevance
- agency/cabinet/ESN feedback is treated as an employer-type preference: strong Web/Digital Analytics roles are kept, but annonceur roles are favored for review

You can tune fit weights and location patterns in:

```text
config\preferences.json
```

Keep the weights moderate if you want to avoid missing relevant jobs. The role score should remain the strongest signal; preference scores are mainly for ordering review priority.

## Deduplication

Jobs are merged by normalized company family, role title, and location family, not by URL alone. This helps catch:

- the same job reposted with a different published date
- the same offer appearing on LinkedIn and Welcome to the Jungle
- titles with platform noise such as `CDI`, `H/F`, city suffixes, or company names inside the title

When duplicates are merged, the visible `Sources` column can contain multiple platforms. Extra URLs are kept in hidden workbook fields.

## Welcome To The Jungle

The script supports the official WelcomeKit API when a token is available:

```powershell
$env:WK_API_KEY = "your_api_key"
powershell -ExecutionPolicy Bypass -File .\Find-AnalyticsJobs.ps1
```

To persist the token for future manual runs:

```powershell
[Environment]::SetEnvironmentVariable("WK_API_KEY", "your_api_key", "User")
```

Without `WK_API_KEY`, it uses a public WTTJ sitemap fallback.

## LinkedIn

LinkedIn is queried through public guest job endpoints only. The script does not log in, bypass CAPTCHA, or use a private account.

## Manual Use

The crawler is manual-only. Nothing in this project is scheduled to run at Windows startup or at a fixed time.

## Maintenance

- Close `jobs_tracker.xlsx` before crawling, formatting, or updating status.
- Keep `output\jobs_tracker.xlsx` as the only working tracker file.
- Keep recent files in `output\backups` only for rollback; old backups are pruned automatically.
- Run `Test-JobTrackerHealth.ps1` after larger changes or if the workbook looks odd.
- Run `Test-ScoringRules.ps1` after changing matching, feedback, or preference rules. It does not require Excel.
- Shared workbook schema and styling helpers live in `JobTracker.Common.ps1`.

## Adjust Defaults

```powershell
powershell -ExecutionPolicy Bypass -File .\Find-AnalyticsJobs.ps1 -DaysBack 7 -Location "France"
```
