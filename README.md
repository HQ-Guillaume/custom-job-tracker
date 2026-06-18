# Custom Job Tracker

Rolling crawler and lightweight application tracker for customizable job-search profiles.

This repository is designed as a clean public release. Personal tracker data, credentials, local overrides, caches, backups, diagnostics, and generated workbooks stay outside Git.

The public default profile is `Digital Analytics`, and users can create additional profiles from the GUI without editing JSON.

## Requirements

- Windows for the GUI launchers
- Windows PowerShell 5.1 or PowerShell 7+
- Desktop Microsoft Excel is optional; when it is missing, the built-in OpenXML writer still creates `jobs_tracker.xlsx`
- Internet access for crawling public job sources

macOS/Linux compatibility is partial today: the current GUI uses WinForms and the `.cmd`/`.vbs` launchers are Windows-only, but the core PowerShell crawler and no-Excel XLSX writer are designed to be portable in PowerShell 7. In a GitHub source checkout, run `tools\tests\Test-EnvironmentCompatibility.ps1` to see exactly what is supported on a machine.

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
2. Double-click `Run-CustomJobTracker-GUI.vbs`.
3. Keep the default `Digital Analytics` profile, or create a custom profile from the GUI.
4. Keep the default public sources enabled, or add credentials for optional API sources.
5. Click `Create tracker` to create an empty workbook, or click `Run crawl` to create and populate it.
6. Use `output\jobs_tracker.xlsx` as your private tracker.

The public release does not include a real tracker workbook. Each user creates their own local workbook.
Profiles, source checkboxes, and crawl modes are read from `config\profiles\`, `config\sources.json`, and `config\crawl_modes.json`, so public defaults and private local overrides stay in one place.

## Project Layout

The root folder is kept for daily-use entry points only:

- `Run-CustomJobTracker-GUI.vbs`: recommended no-console launcher
- `Run-CustomJobTracker-GUI.cmd`: visible fallback launcher for the GUI
- `Run-CustomJobTracker.cmd`: command-window fallback crawler
- `README.md`: usage and maintenance guide

Internal files are grouped by purpose:

- `app\cli\`: runnable PowerShell scripts and GUI helper modules used by the launchers
- `app\core\`: shared crawler, matching, deduplication, Excel, profile, and config modules
- `app\sources\`: platform-specific crawlers
- `config\`: public defaults, public profiles, and ignored local overrides
- `tools\`: source-repository tests, diagnostics, and release helpers
- `output\`: ignored local tracker, cache, diagnostics, and backups

The compact release zip excludes `tools\` and `.github\`; they stay in the GitHub source repository for development and maintenance.

## Launch

Recommended: double-click the no-console WinForms launcher:

```text
Run-CustomJobTracker-GUI.vbs
```

Fallback if Windows blocks `.vbs` files: double-click the command launcher:

```text
Run-CustomJobTracker-GUI.cmd
```

The GUI lets you choose a profile, create/edit/duplicate profiles without editing JSON, choose Fast/Default/Deep mode, enable or disable sources, see live progress logs, open the tracker, clean old managed cache/log files, force a fresh fetch when needed, and check whether credentials are configured.

Command-line fallback: double-click:

```text
Run-CustomJobTracker.cmd
```

The launcher asks for a crawl mode:

- `Fast`: lower page/detail caps for quick checks
- `Default`: balanced mode for normal manual runs
- `Deep`: wider crawl for more coverage when runtime matters less

You can also pass the mode directly:

```text
Run-CustomJobTracker.cmd Fast
Run-CustomJobTracker.cmd Default
Run-CustomJobTracker.cmd Deep
```

Custom modes added to `config\crawl_modes.json` can also be passed directly to the command launcher.

or run:

```powershell
cd "path\to\custom-job-tracker"
powershell -ExecutionPolicy Bypass -File .\app\cli\Find-AnalyticsJobs.ps1
```

Useful crawl modes:

```powershell
powershell -ExecutionPolicy Bypass -File .\app\cli\Find-AnalyticsJobs.ps1 -CrawlMode Fast
powershell -ExecutionPolicy Bypass -File .\app\cli\Find-AnalyticsJobs.ps1 -CrawlMode Default
powershell -ExecutionPolicy Bypass -File .\app\cli\Find-AnalyticsJobs.ps1 -CrawlMode Deep
```

- `Fast`: lower page/detail caps for quick daily checks
- `Default`: balanced mode with generous LinkedIn detail cap
- `Deep`: wider crawl for more coverage when runtime matters less

To run a specific profile from the command line:

```powershell
powershell -ExecutionPolicy Bypass -File .\app\cli\Find-AnalyticsJobs.ps1 -Profile digital_analytics -CrawlMode Default
```

The GUI writes custom profiles to ignored local files under `config\local\profiles\`.

## Status Updates

You can edit the workbook directly, or use:

```powershell
powershell -ExecutionPolicy Bypass -File .\app\cli\Update-JobStatus.ps1 -JobId "JOB_ID" -Status applied -Notes "Applied on LinkedIn"
```

When status is set to `applied`, `applied_date` is filled with today by default unless you provide one:

```powershell
powershell -ExecutionPolicy Bypass -File .\app\cli\Update-JobStatus.ps1 -JobId "JOB_ID" -Status applied -AppliedDate 2026-06-13
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
powershell -ExecutionPolicy Bypass -File .\app\cli\Export-JobTrackerXlsx.ps1
```

In `auto` mode this uses desktop Excel when available, otherwise it uses the built-in OpenXML writer.

To check workbook health without crawling:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\tests\Test-JobTrackerHealth.ps1
```

The health check is available in a GitHub source checkout. It opens the workbook read-only and verifies the expected sheets, columns, hidden backend fields, clickable links, status values, duplicate job IDs, and status row formatting.
If desktop Excel is not available, it falls back to the no-Excel OpenXML workbook health checker:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\diagnostics\Test-WorkbookHealthOpenXml.ps1
```

To compare a test crawl workbook against the current master:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\diagnostics\Compare-JobTrackerWorkbooks.ps1 -CandidatePath .\output\jobs_tracker_test.xlsx
```

## Public Release Privacy

The repository and release assets should never contain personal or secret data:

- `output\jobs_tracker.xlsx`, backups, cache, or diagnostics
- CVs, resumes, screenshots with credentials, or personal application notes
- `.env`, `.key`, `.secret`, `config\local*.json`, or `config\local\*`
- absolute machine-specific user paths
- real API credentials

The compact release zip also excludes source-only maintenance folders such as `tools\` and `.github\`.

Before publishing a release, run:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\release\Test-ReleaseSafety.ps1
```

To create a clean zip from committed public files only:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\release\New-PublicReleasePackage.ps1 -Version v1.0.0
```

## Profiles

A profile defines the job-search intent: target titles, important skills, exclusion keywords, search queries, target/excluded locations, excluded contracts, employer preference, matching signals, and ignore-note templates.

The public default profile is:

```text
config\profiles\digital_analytics.json
```

Use the GUI buttons in `Crawl setup`:

- `New`: create a profile from form fields
- `Edit`: customize the selected profile locally
- `Duplicate`: copy an existing profile and adjust it for another search
- `Set default`: remember the selected profile for future launches

Custom profiles are saved in:

```text
config\local\profiles\PROFILE_ID.json
```

The default-profile choice is saved in:

```text
config\local.runtime.json
```

Both locations are ignored by Git. They are personal local settings and are not included in public releases.

Profile form fields are intentionally plain text. Put one title, skill, keyword, location, or query per line; the crawler saves a compact `profile_builder` file and generates matching rules, source queries, and fit preferences at load time.

Advanced profiles can override source query pools explicitly under `sources.queries`. The crawler first looks for the source-specific list, then falls back to `api` when a source list is empty:

```json
"sources": {
  "queries": {
    "linkedin": [],
    "hellowork": [],
    "apec": [],
    "france_travail": [],
    "adzuna": [],
    "api": []
  }
}
```

## Matching And Ranking

Main role-matching signals:

- title signals: generated from the active profile's target titles
- description/tool signals: generated from the active profile's important skills
- global safety signals: broadly irrelevant roles such as SEO/SEA-only, data engineering, HR/recruiting, and excluded contracts are demoted or removed
- ranking: `High` >= 80, `Medium` >= 50, `Review` >= 35
- jobs with only description/tool matches and no profile-related title are kept but capped at `Review`

The final `Match` uses several dimensions:

- `Role score`: profile relevance from title and description
- `Employer fit`: annonceur is favored; agency, consulting, and ESN are demoted but not excluded
- `Location fit`: Paris/Ile-de-France/France/remote signals are favored; foreign locations are demoted
- `Seniority fit`: internship/junior/managerial roles are demoted
- `Contract fit`: CDI/permanent/full-time is favored; CDD/apprenticeship/internship/freelance are excluded before export

The tracker also uses your history:

- at the beginning of every manual crawl, the programme reads the saved `jobs_tracker.xlsx` and builds a fresh feedback profile from your `Status` and `Apply notes`; this is recalculated from the workbook each run, so it does not accumulate duplicate learning over time
- similar jobs to `applied`, `interview`, `offer`, or `interesting` can receive a small score boost
- similar jobs to `ignored` can receive a score penalty
- ignored jobs with structured `ignore_reason=...` notes teach the crawler more precisely: SEO/SEA rejects affect marketing roles, data-engineering rejects affect dbt/Snowflake/pipeline roles, and `duplicate` does not reduce relevance
- agency/cabinet/ESN feedback is treated as an employer-type preference: strong profile matches are kept, but annonceur roles can be favored for review

For the default Digital Analytics search, tune profile titles, skills, query pools, ignored-reason templates, fit weights, and location patterns in:

```text
config\profiles\digital_analytics.json
```

More tunable values are in:

```text
config\runtime.json        # default days, location, tracker path, cache, delays
config\crawl_modes.json    # Fast / Default / Deep source caps
config\sources.json        # source order, endpoints, source defaults, credential environment variable names
config\profiles\*.json     # compact profile builders and any explicit profile overrides
config\matching_rules.json # global matching thresholds
config\workbook.json       # workbook backend, status dropdowns, and sheet names
```

For private machine-specific changes, create ignored local override files instead of editing public defaults:

```text
config\local.runtime.json
config\local.sources.json
config\local.preferences.json
config\local.matching_rules.json
config\local.workbook.json
config\local.crawl_modes.json
config\local\profiles\your_profile.json
```

There is no public `config\preferences.json` because the default preferences are generated from the active profile and built-in defaults. A private `config\local.preferences.json` is still supported if you need a machine-specific preference override.

You can also use a `config\local\` folder with files such as `config\local\sources.json`. Public config is loaded first, then local overrides are merged on top.

Validate config without crawling:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\tests\Test-JobCrawlerConfig.ps1
```

Keep the weights moderate if you want to avoid missing relevant jobs. The role score should remain the strongest signal; preference scores are mainly for ordering review priority.

## Pipeline

Source crawlers may skip obvious old or excluded-contract jobs early to save time, but those early skips are only an optimization. The authoritative workflow is centralized in `app\core\JobTracker.Pipeline.ps1`:

```text
fetch source results
normalize to tracker rows
enrich details
apply post-enrichment eligibility
apply feedback scoring
deduplicate and merge with tracker history
apply final invariant validation
export workbook
```

The pipeline gate owns hard rules such as published-date retention, excluded contracts, and invalid WTTJ locations. Merge and export call the same gate, so a late detail fetch or an existing tracker row cannot reintroduce a non-application CDD/freelance/internship/apprenticeship job or an old non-application job. Application-history statuses are intentionally preserved outside the 7-day window.

## Deduplication

Jobs are merged with a hierarchy rather than one fragile key:

```text
1. same canonical URL
2. same platform/source job id
3. same company alias + strong normalized title
4. same company alias + same role family + compatible location
```

Company names are normalized automatically, so labels such as `NEXTON`, `Nexton Consulting`, `L'Olivier Assurance`, and `Olivier` can still meet the same-company condition when the title/location evidence also supports a duplicate. Explicit company alias groups can also be added in `config\matching_rules.json` under `deduplication.company_aliases`, or privately in `config\local.matching_rules.json`.

Jobs are still not merged by company alone. The alias layer only helps when URL, source id, title similarity, role family, or location compatibility supports the duplicate. This helps catch:

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

Credentialed sources are disabled by default in the public config. Enable them from the GUI source checkboxes, with `-EnableSource <source_key>`, or through a local config override. The legacy per-source switches such as `-EnableFranceTravail`, `-EnableAdzuna`, and `-EnableWelcomeKit` are still supported for compatibility.

## Welcome To The Jungle

The script supports the official WelcomeKit API when a token is available:

```powershell
$env:WK_API_KEY = "your_api_key"
powershell -ExecutionPolicy Bypass -File .\app\cli\Find-AnalyticsJobs.ps1 -EnableSource welcome_kit
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
powershell -ExecutionPolicy Bypass -File .\app\cli\Find-AnalyticsJobs.ps1 -EnableSource france_travail
```

Optional scope override:

```powershell
[Environment]::SetEnvironmentVariable("FRANCE_TRAVAIL_SCOPE", "api_offresdemploiv2 o2dsoffre", "User")
```

The crawler searches the active profile query pool, asks the API for jobs published in the last 7 days, then maps France Travail fields into the same tracker columns: title, company, city/region, contract, URL, published date, match score, and source. If France Travail only returns a board/origin name instead of a real employer, that generic origin is not used as a company dedupe key.

## Adzuna

Adzuna is supported through the official jobs API. It is disabled by default in the public release and skipped unless credentials are configured and the source is enabled:

```powershell
[Environment]::SetEnvironmentVariable("ADZUNA_APP_ID", "your_app_id", "User")
[Environment]::SetEnvironmentVariable("ADZUNA_APP_KEY", "your_app_key", "User")
powershell -ExecutionPolicy Bypass -File .\app\cli\Find-AnalyticsJobs.ps1 -EnableSource adzuna
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

## Workbook Backend

The default workbook backend is configured in `config\workbook.json`:

```json
"output_backend": "auto"
```

- `auto`: use desktop Excel COM when available; otherwise create `jobs_tracker.xlsx` with the built-in OpenXML writer
- `excel`: require desktop Excel COM and fail clearly if it is missing
- `openxml`: always use the no-Excel XLSX writer

The Excel backend gives the richest Windows rendering. The OpenXML backend keeps the same `.xlsx` tracker format, clickable links, hidden backend columns, dropdown validations, filters, freeze panes, summary/settings sheets, and status-based conditional formatting without requiring desktop Excel. If both XLSX writers fail, the program writes a readable `.html` fallback and reports the XLSX export failure.

## Compatibility

Full GUI support today is Windows:

- the GUI uses Windows Forms
- `.cmd` and `.vbs` launchers are Windows launchers
- desktop Excel is optional and used only when available in `auto` mode

On macOS or Linux, PowerShell 7 can run the core crawler and OpenXML workbook output, but the current GUI and launchers remain Windows-only. A future fully cross-platform app would need a non-WinForms GUI.

In a GitHub source checkout, check a machine with:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\tests\Test-EnvironmentCompatibility.ps1
```

## Maintenance

- Close `jobs_tracker.xlsx` before crawling, formatting, or updating status.
- Keep `output\jobs_tracker.xlsx` as the only working tracker file.
- Keep recent files in `output\backups` only for rollback; old backups are pruned automatically.
- Daily runnable scripts live in `app\cli\`.
- Runtime modules live in `app\core\`.
- Hard workflow gates live in `app\core\JobTracker.Pipeline.ps1`; change this module when a rule must apply to every source and merge/export path.
- Source-specific crawlers live in `app\sources\`. Files named `Source.*.ps1` are auto-discovered; source metadata still belongs in `config\sources.json`.
- Source adapter validation lives in `app\core\JobTracker.SourceAdapter.ps1`.
- Output/cache cleanup helpers live in `app\core\JobTracker.OutputMaintenance.ps1`.
- Development tools and parser fixtures live in `tools\` in the GitHub source repository and are excluded from the compact release zip.
- Run `tools\tests\Run-AllTests.ps1` after larger source changes. Add `-CoreOnly` or `-SkipGui` for non-GUI checks, and add `-IncludeWorkbookHealth` when Excel is available and you also want to inspect the tracker workbook.
- Run `tools\tests\Test-EnvironmentCompatibility.ps1` after changing launchers, Excel integration, or setup requirements.
- Run `tools\tests\Test-ProfileBuilder.ps1` after changing profile creation, profile expansion, or default profile config.
- Run `tools\tests\Test-JobTrackerHealth.ps1` after larger changes or if the workbook looks odd.
- Run `tools\tests\Test-ScoringRules.ps1` after changing matching, feedback, or preference rules. It does not require Excel.
- Run `tools\tests\Test-ParserFixtures.ps1` after changing APEC, HelloWork, LinkedIn, or dedupe parsing. It does not require Excel or network access.
- Run `tools\tests\Test-PipelineGuards.ps1` after changing retention, contract, location, merge, or final export eligibility.
- Run `tools\tests\Test-SourceAdapters.ps1` after adding or renaming a source adapter.
- Run `tools\tests\Test-Integration.ps1` after changing source orchestration, deduplication, cache pruning, or run history.
- Run `tools\release\Test-ReleaseSafety.ps1` before publishing a public release.
- Shared workbook schema and styling helpers live in `app\core\JobTracker.Common.ps1`.

## Adjust Defaults

```powershell
powershell -ExecutionPolicy Bypass -File .\app\cli\Find-AnalyticsJobs.ps1 -DaysBack 7 -Location "France"
```

You can disable individual sources for diagnostics. Use a comma-separated list when passing several source keys:

```powershell
powershell -ExecutionPolicy Bypass -File .\app\cli\Find-AnalyticsJobs.ps1 -SkipSource france_travail,adzuna,apec,hellowork,wttj_public,welcome_kit,linkedin
```

For Welcome to the Jungle specifically, `wttj_public` controls the public fallback and `welcome_kit` controls the official API. Legacy switches such as `-SkipWttj`, `-DisableWttjPublicFallback`, and `-DisableWelcomeKit` are still supported for compatibility.

Useful speed knobs:

```powershell
powershell -ExecutionPolicy Bypass -File .\app\cli\Find-AnalyticsJobs.ps1 -CrawlMode Fast -MaxLinkedInDetails 50
```

Useful maintenance modes:

```powershell
powershell -ExecutionPolicy Bypass -File .\app\cli\Find-AnalyticsJobs.ps1 -DryRun
powershell -ExecutionPolicy Bypass -File .\app\cli\Find-AnalyticsJobs.ps1 -DiagnosticMode
powershell -ExecutionPolicy Bypass -File .\app\cli\Find-AnalyticsJobs.ps1 -ValidateConfig
```

`-DryRun` crawls and merges in memory without writing the workbook. `-DiagnosticMode` writes `output\diagnostics\crawl_diagnostics_*.csv` with matched pre-filter rows and contract-exclusion status.

To bypass the local detail-page cache for a fresh diagnostic run:

```powershell
powershell -ExecutionPolicy Bypass -File .\app\cli\Find-AnalyticsJobs.ps1 -DisableCache
```

To review or clean managed output without opening the GUI:

```powershell
powershell -ExecutionPolicy Bypass -File .\app\cli\Clear-JobTrackerOutput.ps1
powershell -ExecutionPolicy Bypass -File .\app\cli\Clear-JobTrackerOutput.ps1 -Cache -Logs
```

To review how your status and ignore notes could tune future matching:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\diagnostics\Get-FeedbackTuningSuggestions.ps1
```
