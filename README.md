# Survey Response Analyzer

A static browser application for Excel, CSV, and public Google Sheets surveys. Uploaded files are processed locally. There is no analytics backend, account, database, or paid API. Public Google Sheets are fetched directly from Google without credentials.

## Workspaces

- **Charts:** choose questions, generate chart cards, compare questions, add checklist filters, merge or hide answers, and export images or CSV. The sheet selector belongs to this workspace. Switching sheets preserves each sheet's chart settings for the current browser session.
- **Breakdown Report:** choose the report source and primary sheet within its options. Optionally link another sheet or file, choose questions and a grouping column, and combine multiple filters. Sheet changes keep this tab open. Settings and selections are retained per source/sheet during the session.
- **Data Preview:** inspect source row numbers, search records, and hide columns from charts and reports without modifying the workbook. Chart summary tables support paging; the raw preview shows the first 100 matching rows.

The file status and active source/sheet/row counts appear above the application title. **Clear data** clears every loaded source, chart, link, and report. A replacement file is parsed and validated before the current analysis is cleared; failed imports preserve it.

## Import rules

Use the first row as headers. Completely blank records are ignored, while source row numbers retain gaps. Duplicate, blank, and reserved internal headers receive unique names; source values are preserved. The optional **Data Dictionary** sheet maps sheet name (A), source header (B), and full display question (C), with an optional header row.

CSV supports comma, tab, and semicolon field separators, quoted multiline fields, Unicode BOMs, and malformed-quote errors. Response separators are separate from CSV field separators: select semicolon (default), comma, pipe, line break, or **Single answer / do not split**. Secondary linked questions have their own separator setting. Dates use a consistent ISO representation in analysis.

## Counts and percentages

Answer labels match case-insensitively after whitespace/quote normalization. Repeated choices and choices merged into one label count once per respondent. **Other** represents the union of respondents in the capped categories, with a distinct name when a literal Other answer exists. Choose **All responses** to remove chart caps.

Charts exclude blanks by default. Including blanks changes the percentage base to all included respondents. Binary 0/1 labels change to No/Yes only when enabled. Pie/doughnut and 100% stacked views switch to suitable bar views when choices overlap.

Reports calculate percentages using answered respondents for each question within each group. Each section includes **Answered**, **Missing**, and **Eligible** totals. An unanswered group has a missing percentage (displayed as a dash), rather than an invented 0%. Multiple-choice percentages can sum above 100%; the base is respondents rather than selections. Expected scale labels never suppress unexpected observed answers.

Questions and grouping columns with more than 15 choices are hidden by default; enable the explicit override to include them. Filters support fewer than 500 unique values, value search, select/clear controls, and AND between filter columns. Within each column, a respondent qualifies if any chosen category matches. **Select visible** applies to search results; **Select all** and **Clear all** apply to the entire question list.

Changing report settings marks the previous preview stale and blocks exports until regeneration. Generation runs in a worker over HTTP with progress, timeout, and cancellation; direct file opening uses a yielding fallback.

## Linked surveys

Open **Link a secondary survey** within the report options (or use the Charts shortcut). Match the selected report primary sheet to another sheet or file. Normalized matching tolerates case, spacing, and punctuation; **Exact identifier** preserves exact text matching. Normalization collisions are disclosed. Duplicate secondary keys are ambiguous and excluded; unmatched primary rows are inspectable/downloadable with original row numbers.

Select a secondary question after matching. Multi-choice categories can overlap; coverage distinguishes matched sites and survey response rows. Links are scoped to the matched primary workbook/sheet. Changing the report primary context clears the match. Linked questions are available to Charts when its active sheet matches that primary context.

## Exports

CSV/Excel reports include source, sheet, filter, grouping, generation time, included/excluded rows, and denominator information, plus linked category coverage where applicable. Choose counts, percentages, or both independently of preview mode. Excel percentages are numeric cells formatted as percentages; base totals remain numbers. CSV includes a UTF-8 BOM and protects formula-like text. Excel text is exported as text, never as executable formulas.

## Run and verify

Use Node.js 24 or later:

```sh
npm test
npm run build
python -m http.server 8765
```

Open `http://localhost:8765`. Opening `index.html` directly also works with parsing/report fallbacks, but serving over HTTP enables workers. `npm run build` produces static assets in `dist/` and a Fetch-compatible worker in `dist/server/index.js`. Both root and `/Survey_breakdowm` worker routes are supported, including binary logo assets. GitHub Pages can serve the repository's root from the main branch after changes are merged.

Dependencies are pinned and bundled in `vendor/` so uploads do not require CDN script requests: SheetJS 0.20.3, Chart.js 4.4.7, and chartjs-plugin-datalabels 2.2.0. License notices and checksums are included. A content security policy limits scripts/workers to local assets and permits Google export connections.

## Practical limits

This is descriptive survey analysis, without weighting, statistical tests, or imputation. The application does not recalculate Excel formulas; cached values must exist. Very wide, high-cardinality reports can still consume substantial memory and render slowly, even though calculation runs off the UI thread. Generate only needed questions for large datasets. Session state is not retained across browser reloads; the app warns before leaving when analysis exists. Raw data preview remains capped at 100 matching rows. Private Google Sheets require an authenticated integration and are not supported.

## CI template

`ci/validate.yml` provides GitHub Actions test/build validation. Move it to `.github/workflows/validate.yml` to enable it. The current GitHub authentication token lacks the `workflow` scope, so the implementation PR includes a ready template without changing repository automation permissions.
