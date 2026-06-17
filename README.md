# Analytics Job Tracker

Rolling crawler and lightweight application tracker for Web/Digital Analytics jobs.

This repository is designed as a clean public release. Personal tracker data, credentials, local overrides, caches, backups, diagnostics, and generated workbooks stay outside Git.

## Requirements

- Windows
- Windows PowerShell 5.1
- Desktop Microsoft Excel for workbook creation/formatting
- Internet access for crawling public job sources

## Main File

Use this single workbook:

```text
output\jobs_tracker.xlsx
```

The crawler creates and updates the same workbook each time. The `output` folder is ignored by Git so your personal tracker, notes, and backups stay local.

It keeps:

- matching jobs whose `Published` date is within the last 7 days
- older jobs only when their `status` is application-related: `applied`, `interview`, `offer`, `rejected`, or `withdrawn`

CDD, apprenticeship, internship, and freelance jobs are excluded from new crawl results.

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

`Summary` contains the latest crawl report, including the crawl mode, published-date retention rule, match levels, employer-type distribution, source diagnostics, fit demotions, and backup path.

Additional sheets:

- `Settings`: active mode, caps, config path, cache path, credentials detected/missing, query counts
- `Source Health`: per-platform duration, requests, cache hits, skipped counts, errors, matches
- `Feedback Quality`: ignored rows missing structured notes, application rows missing dates, and other feedback hygiene checks

Close `jobs_tracker.xlsx` before launching the crawler so Excel does not lock the file.

## First Run

1. Download or clone the repository.
2. Double-click `Run-AnalyticsJobCrawler-GUI.vbs`.
3. Keep the default public sources enabled, or add credentials for optional API sources.
4. Click `Create tracker` to create an empty workbook, or click `Run crawl` to create and populate it.
5. Use `output\jobs_tracker.xlsx` as your private tracker.

The public release does not include a real tracker workbook. Each user creates their own local workbook.
Source checkboxes and crawl modes are read from `config\sources.json` and `config\crawl_modes.json`, so public defaults and private local overrides stay in one place.

## Launch

Recommended: double-click the no-console WinForms launcher:

```text
Run-AnalyticsJobCrawler-GUI.vbs
```

Fallback if Windows blocks `.vbs` files: double-click the command launcher:

```text
Run-AnalyticsJobCrawler-GUI.cmd
```

The GUI lets you choose Fast/Default/Deep mode, enable or disable sources, run dry-run/diagnostic crawls, see live progress logs, open the tracker, and check whether credentials are configured.

Command-line fallback: double-click:

```text
Run-AnalyticsJobCrawler.cmd
```

The launcher asks for a crawl mode:

- `Fast`: lower page/detail caps for quick checks
- `Default`: balanced mode for normal manual runs
- `Deep`: wider crawl for more coverage when runtime matters less

You can also pass the mode directly:

```text
Run-AnalyticsJobCrawler.cmd Fast
Run-AnalyticsJobCrawler.cmd Default
Run-AnalyticsJobCrawler.cmd Deep
```

Custom modes added to `config\crawl_modes.json` can also be passed directly to the command launcher.

or run:

```powershell
cd "path\to\analytics-job-tracker"
powershell -ExecutionPolicy Bypass -File .\Find-AnalyticsJobs.ps1
```

Useful crawl modes:

```powershell
powershell -ExecutionPolicy Bypass -File .\Find-AnalyticsJobs.ps1 -CrawlMode Fast
powershell -ExecutionPolicy Bypass -File .\Find-AnalyticsJobs.ps1 -CrawlMode Default
powershell -ExecutionPolicy Bypass -File .\Find-AnalyticsJobs.ps1 -CrawlMode Deep
```

- `Fast`: lower page/detail caps for quick daily checks
- `Default`: balanced mode with generous LinkedIn detail cap
- `Deep`: wider crawl for more coverage when runtime matters less

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

The crawler keeps these local output files:

- `output\jobs_tracker.xlsx`: source of truth
- `output\backups\*.xlsx`: recent automatic workbook backups before tracker/status updates
- `output\cache\*`: local detail-page cache for slow public sources
- `output\run_history.jsonl`: compact local run history for troubleshooting crawl duration and source health

These files are local personal data and are not committed to the GitHub repository.

To re-apply workbook formatting without crawling:

```powershell
powershell -ExecutionPolicy Bypass -File .\Export-JobTrackerXlsx.ps1
```

This requires desktop Microsoft Excel.

To check workbook health without crawling:

```powershell
powershell -ExecutionPolicy Bypass -File .\dev-tools\Test-JobTrackerHealth.ps1
```

The health check opens the workbook read-only and verifies the expected sheets, columns, hidden backend fields, clickable links, status values, duplicate job IDs, and status row formatting.

To compare a test crawl workbook against the current master:

```powershell
powershell -ExecutionPolicy Bypass -File .\dev-tools\Compare-JobTrackerWorkbooks.ps1 -CandidatePath .\output\jobs_tracker_test.xlsx
```

## Public Release Privacy

The repository and release assets should not contain:

- `output\jobs_tracker.xlsx`, backups, cache, or diagnostics
- CVs, resumes, screenshots with credentials, or personal application notes
- `.env`, `.key`, `.secret`, `config\local*.json`, or `config\local\*`
- absolute machine-specific user paths
- real API credentials

Before publishing a release, run:

```powershell
powershell -ExecutionPolicy Bypass -File .\dev-tools\Test-ReleaseSafety.ps1
```

To create a clean zip from committed public files only:

```powershell
powershell -ExecutionPolicy Bypass -File .\dev-tools\New-PublicReleasePackage.ps1 -Version v1.0.0
```

## Matching And Ranking

Main role-matching signals:

- title signals: Web Analyst, Digital Analyst, Digital Analytics Consultant, Tracking, Web Analytics, digital performance, CRO, plus Data Analyst when another relevant signal is present
- description/tool signals: Google Tag Manager, GTM, Google Analytics, GA4, Piano Analytics, ContentSquare, Tag Commander, Commanders Act, Tealium, dataLayer, tagging plan, server-side tracking, consent mode, RGPD/GDPR, A/B testing, dashboards, KPIs
- ranking: `High` >= 80, `Medium` >= 50, `Review` >= 35
- jobs with only description/tool matches and no analytics-related title are kept but capped at `Review`

The final `Match` uses several dimensions:

- `Role score`: web/digital analytics relevance from title and description
- `Employer fit`: annonceur is favored; agency, consulting, and ESN are demoted but not excluded
- `Location fit`: Paris/Ile-de-France/France/remote signals are favored; foreign locations are demoted
- `Seniority fit`: internship/junior/managerial roles are demoted
- `Contract fit`: CDI/permanent/full-time is favored; CDD/apprenticeship/internship/freelance are excluded before export

The tracker also uses your history:

- at the beginning of every manual crawl, the programme reads the saved `jobs_tracker.xlsx` and builds a fresh feedback profile from your `Status` and `Apply notes`; this is recalculated from the workbook each run, so it does not accumulate duplicate learning over time
- similar jobs to `applied`, `interview`, `offer`, or `interesting` can receive a small score boost
- similar jobs to `ignored` can receive a score penalty
- ignored jobs with structured `ignore_reason=...` notes teach the crawler more precisely: SEO/SEA rejects affect marketing roles, data-engineering rejects affect dbt/Snowflake/pipeline roles, and `duplicate` does not reduce relevance
- agency/cabinet/ESN feedback is treated as an employer-type preference: strong Web/Digital Analytics roles are kept, but annonceur roles are favored for review

You can tune fit weights and location patterns in:

```text
config\preferences.json
```

More tunable values are in:

```text
config\runtime.json        # default days, location, tracker path, cache, delays
config\crawl_modes.json    # Fast / Default / Deep source caps
config\sources.json        # source order, endpoints, query pools, credential environment variable names
config\matching_rules.json # matching thresholds, positive signals, negative signals, feedback-learning rules
config\workbook.json       # status dropdowns, ignored-reason templates, sheet names
```

For private machine-specific changes, create ignored local override files instead of editing public defaults:

```text
config\local.runtime.json
config\local.sources.json
config\local.preferences.json
config\local.matching_rules.json
config\local.workbook.json
config\local.crawl_modes.json
```

You can also use a `config\local\` folder with files such as `config\local\sources.json`. Public config is loaded first, then local overrides are merged on top.

Validate config without crawling:

```powershell
powershell -ExecutionPolicy Bypass -File .\dev-tools\Test-JobCrawlerConfig.ps1
```

Keep the weights moderate if you want to avoid missing relevant jobs. The role score should remain the strongest signal; preference scores are mainly for ordering review priority.

## Deduplication

Jobs are merged by normalized company family, role title, and location family, not by URL alone. This helps catch:

- the same job reposted with a different published date
- the same offer appearing on LinkedIn, Welcome to the Jungle, APEC, HelloWork, France Travail, or Adzuna
- titles with platform noise such as `CDI`, `H/F`, city suffixes, or company names inside the title

When duplicates are merged, the visible `Sources` column can contain multiple platforms. Extra URLs are kept in hidden workbook fields. `Source count` means unique platforms, not the number of raw URLs, so one opportunity found on LinkedIn, APEC, France Travail, and Adzuna is counted as four sources but remains one row.

## Platforms And Credentials

Public default sources that do not require credentials:

- APEC
- HelloWork
- Welcome to the Jungle public sitemap fallback
- LinkedIn public guest endpoints

Optional sources that require user-provided credentials:

- France Travail API: `FRANCE_TRAVAIL_CLIENT_ID`, `FRANCE_TRAVAIL_CLIENT_SECRET`, optional `FRANCE_TRAVAIL_SCOPE`
- Adzuna API: `ADZUNA_APP_ID`, `ADZUNA_APP_KEY`
- WelcomeKit official API: `WK_API_KEY`

Credentialed sources are disabled by default in the public config. Enable them from the GUI source checkboxes, with command-line switches such as `-EnableFranceTravail`, `-EnableAdzuna`, or `-EnableWelcomeKit`, or through a local config override.

## Welcome To The Jungle

The script supports the official WelcomeKit API when a token is available:

```powershell
$env:WK_API_KEY = "your_api_key"
powershell -ExecutionPolicy Bypass -File .\Find-AnalyticsJobs.ps1 -EnableWelcomeKit
```

To persist the token for future manual runs:

```powershell
[Environment]::SetEnvironmentVariable("WK_API_KEY", "your_api_key", "User")
```

Without `WK_API_KEY`, it uses a public WTTJ sitemap fallback.

## Credential Storage

The JSON config files store credential variable names only. The real credential values are stored outside the project in Windows User environment variables, so they are not committed to Git and are not written into the tracker workbook.

You can set or update credentials from the GUI with `Set credential`, or with PowerShell:

```powershell
[Environment]::SetEnvironmentVariable("WK_API_KEY", "your_api_key", "User")
[Environment]::SetEnvironmentVariable("FRANCE_TRAVAIL_CLIENT_ID", "your_client_id", "User")
[Environment]::SetEnvironmentVariable("FRANCE_TRAVAIL_CLIENT_SECRET", "your_client_secret", "User")
[Environment]::SetEnvironmentVariable("ADZUNA_APP_ID", "your_app_id", "User")
[Environment]::SetEnvironmentVariable("ADZUNA_APP_KEY", "your_app_key", "User")
```

The variable names themselves are configured in:

```text
config\sources.json
```

## France Travail

France Travail is supported through the official API Offres d'emploi. It is disabled by default in the public release and skipped unless credentials are configured and the source is enabled:

```powershell
[Environment]::SetEnvironmentVariable("FRANCE_TRAVAIL_CLIENT_ID", "your_client_id", "User")
[Environment]::SetEnvironmentVariable("FRANCE_TRAVAIL_CLIENT_SECRET", "your_client_secret", "User")
powershell -ExecutionPolicy Bypass -File .\Find-AnalyticsJobs.ps1 -EnableFranceTravail
```

Optional scope override:

```powershell
[Environment]::SetEnvironmentVariable("FRANCE_TRAVAIL_SCOPE", "api_offresdemploiv2 o2dsoffre", "User")
```

The crawler searches the same web/digital analytics query pool, asks the API for jobs published in the last 7 days, then maps France Travail fields into the same tracker columns: title, company, city/region, contract, URL, published date, match score, and source. If France Travail only returns a board/origin name instead of a real employer, that generic origin is not used as a company dedupe key.

## Adzuna

Adzuna is supported through the official jobs API. It is disabled by default in the public release and skipped unless credentials are configured and the source is enabled:

```powershell
[Environment]::SetEnvironmentVariable("ADZUNA_APP_ID", "your_app_id", "User")
[Environment]::SetEnvironmentVariable("ADZUNA_APP_KEY", "your_app_key", "User")
powershell -ExecutionPolicy Bypass -File .\Find-AnalyticsJobs.ps1 -EnableAdzuna
```

Adzuna has tighter public API limits, so the crawler uses a small default page count and pauses between calls. It uses `max_days_old` to keep the same 7-day crawl window.

## APEC

APEC is crawled through its public job-search JSON endpoint. It is enabled by default and does not need credentials.

APEC is relatively fast because the search response already contains the title, company, city/region, contract type, published date, and a description snippet. The crawler uses relevance-sorted search results, applies the same 7-day `Published` filter afterwards, reads only a small number of result pages by default, and does not open every APEC detail page.

## HelloWork

HelloWork is crawled from public search and job pages. It is enabled by default and does not need credentials.

To keep the run time reasonable, HelloWork uses a two-step crawl:

- collect unique candidate URLs from search result pages
- fetch only the best-scoring detail pages, capped by `-MaxHelloWorkDetails`
- skip visibly excluded contracts from search cards before opening details

The default cap is intentionally conservative. Increase it only when you want a wider HelloWork pass and accept a longer run.

## LinkedIn

LinkedIn is queried through public guest job endpoints only. The script does not log in, bypass CAPTCHA, or use a private account.

LinkedIn uses a recall-protective candidate queue: it first collects unique search cards, ranks them broadly, and then opens at most `-MaxLinkedInDetails` detail pages. Use `Deep` mode if you want a wider LinkedIn pass.

## Manual Use

The crawler is manual-only. Nothing in this project is scheduled to run at Windows startup or at a fixed time.

## Maintenance

- Close `jobs_tracker.xlsx` before crawling, formatting, or updating status.
- Keep `output\jobs_tracker.xlsx` as the only working tracker file.
- Keep recent files in `output\backups` only for rollback; old backups are pruned automatically.
- Runtime modules live in `app\`.
- Source-specific crawlers live in `app\sources\`.
- Development tools and parser fixtures live in `dev-tools\`.
- Run `dev-tools\Test-JobTrackerHealth.ps1` after larger changes or if the workbook looks odd.
- Run `dev-tools\Test-ScoringRules.ps1` after changing matching, feedback, or preference rules. It does not require Excel.
- Run `dev-tools\Test-ParserFixtures.ps1` after changing APEC, HelloWork, LinkedIn, or dedupe parsing. It does not require Excel or network access.
- Run `dev-tools\Test-Integration.ps1` after changing source orchestration, deduplication, cache pruning, or run history.
- Run `dev-tools\Test-ReleaseSafety.ps1` before publishing a public release.
- Shared workbook schema and styling helpers live in `JobTracker.Common.ps1`.

## Adjust Defaults

```powershell
powershell -ExecutionPolicy Bypass -File .\Find-AnalyticsJobs.ps1 -DaysBack 7 -Location "France"
```

You can disable individual sources for diagnostics:

```powershell
powershell -ExecutionPolicy Bypass -File .\Find-AnalyticsJobs.ps1 -SkipFranceTravail -SkipAdzuna -SkipApec -SkipHelloWork -SkipWttj -SkipLinkedIn
```

For Welcome to the Jungle specifically, `-SkipWttj` disables both the public fallback and WelcomeKit. Use `-DisableWttjPublicFallback` or `-DisableWelcomeKit` only when you want to disable one WTTJ path and keep the other available.

Useful speed knobs:

```powershell
powershell -ExecutionPolicy Bypass -File .\Find-AnalyticsJobs.ps1 -CrawlMode Fast -MaxLinkedInDetails 50
```

Useful maintenance modes:

```powershell
powershell -ExecutionPolicy Bypass -File .\Find-AnalyticsJobs.ps1 -DryRun
powershell -ExecutionPolicy Bypass -File .\Find-AnalyticsJobs.ps1 -DiagnosticMode
powershell -ExecutionPolicy Bypass -File .\Find-AnalyticsJobs.ps1 -ValidateConfig
```

`-DryRun` crawls and merges in memory without writing the workbook. `-DiagnosticMode` writes `output\diagnostics\crawl_diagnostics_*.csv` with matched pre-filter rows and contract-exclusion status.

To bypass the local detail-page cache for a fresh diagnostic run:

```powershell
powershell -ExecutionPolicy Bypass -File .\Find-AnalyticsJobs.ps1 -DisableCache
```
