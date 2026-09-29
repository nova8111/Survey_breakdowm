# Survey Response Analyzer

A static web app for analyzing Excel, CSV, and public Google Sheets survey-style data directly in the browser.

## What it does

- Upload `.xlsx`, `.xls`, or `.csv` files.
- Load a public Google Sheet that does not require sign-in.
- Choose a sheet when an Excel file has multiple sheets.
- Create multiple chart cards from the same file.
- Let users choose any non-metadata response column for a chart, including empty and open-ended-style columns.
- Choose one or more response questions and generate all selected charts with a single action, without duplicating charts already in the workspace.
- Analyze one column by response count and percentage.
- Split semicolon-delimited multi-select responses into individual dashboard choices without changing raw data.
- Compare one column against another.
- Add one or more checklist filters.
- Hide selected responses or manually combine similar response labels.
- Navigate focused Charts, Question Breakdown Report, and Data Preview workspaces.
- Generate sheet-style question breakdown reports from the active workbook or public Google Sheet.
- Preview the first 100 rows, search the dataset, inspect column quality, and hide columns from chart analysis without changing the source.
- Choose response columns to include with checked lists.
- Optionally break those response columns down by one selected column.
- Filter breakdown reports by one selected column with fewer than 500 unique values.
- View the generated breakdown table in a fixed scrollable panel with percentage heat maps.
- Hide open-ended-style columns from report choices when they have more than 15 unique responses.
- Export chart images, summary tables, and filtered data as CSV.
- Export generated breakdown reports as CSV or Excel.
- Optionally link the active survey to another sheet or survey file using user-selected matching fields.
- Review matched/unmatched rates, inspect duplicates, and view or download unmatched primary records.
- Use any linked-survey question as a chart, filter, or report breakdown, including multi-select questions.

Uploaded files are processed only in your browser. The app does not use a backend, database, sign-in, API key, or paid service.

## Files

- `index.html`
- `style.css`
- `script.js`
- `data-dictionary.js`
- `README.md`

The app uses CDN links for:

- SheetJS
- Chart.js
- Chart.js Data Labels plugin

## Eligible chart selection

The Charts workspace lists every non-metadata response column before generating anything. Select individual questions or **Select all**, adjust the selection as needed, and choose **Generate Selected Charts**. Blank responses are not shown in generated charts. Existing charts are preserved, and repeating generation adds only selected questions that do not already have a chart. Each generated chart defaults to an automatic view, and its header **View** selector lets you switch to horizontal or vertical bars, pie, doughnut, line, or table-only output. For stacked charts, choose a comparison column in the chart settings and then select a stacked comparison view.

## Optional Data Dictionary

An uploaded workbook may include a sheet named **Data Dictionary**. Put the survey sheet name in column A, the original/header name in column B, and the full display question in column C. A header row is optional. The analyzer matches trimmed, case-insensitive sheet and header names, applies mappings independently per survey sheet, and falls back to the original header when a mapping or display question is missing. Raw headers remain the internal data keys, so filters, matching, and calculations are not changed by display text.

## Multi-select responses

For dashboard charts, filters, and breakdown reports, semicolons separate multiple selected choices. Each choice is counted independently and choices are sorted by descending count by default. Percentages use the number of respondents with at least one selection as the denominator, so multi-select percentages may total more than 100%. The source workbook, raw response values, and data preview remain unchanged.

## Question Breakdown Report

The Question Breakdown Report uses the currently active dataset. It can use:

- The uploaded workbook.
- A public Google Sheet link that can be opened without signing in. Loaded public sheets can also be charted and used for breakdown reports.

For best results, include sheets like:

- A raw data sheet with one header row.

Response-column choices appear checked by default so you can uncheck anything you do not want in the output. Columns with no responses or more than 15 unique responses are hidden from report choices so empty and open-ended questions do not create unusable breakdowns. The breakdown dropdown starts with no breakdown selected; when you choose one, that choice appears above the on-screen report. The report filter lets you select a column with fewer than 500 unique values, then uncheck values you want to exclude.

Private Google Sheets are not connected in this version because that would require Google sign-in/API setup.

## Linked survey analysis

After loading a primary survey, open **Link a secondary survey**. The secondary survey can be another sheet in the active workbook or a separate Excel/CSV file. Select one matching field from each survey, review the match diagnostics, and then choose a secondary question for disaggregation.

Matching ignores capitalization, surrounding whitespace, common punctuation, and spacing differences. Primary rows without a unique secondary match are reported and excluded only while linked analysis is active. Duplicate secondary keys are treated as ambiguous and are never selected automatically. Removing the link restores the original single-survey workflow.

Multi-select secondary answers separated by commas, semicolons, pipes, or line breaks are expanded into category memberships. A primary response can therefore appear in more than one linked category. Linked reports include category-level matched-site and survey-response counts, and those rows are included in CSV and Excel report exports.

## Run locally

Open `index.html` in your browser.

## Deploy with GitHub Pages

1. Create a GitHub repository.
2. Upload the project files.
3. Open the repository settings.
4. Open the Pages section.
5. Select deployment from the main branch.
6. Open the generated GitHub Pages URL.

## Notes

- The first row is treated as the column header row.
- Completely empty rows are ignored.
- Blank cells can be grouped as `No Response` when a chart is configured to include blanks.
- Zero and `false` values are preserved as valid responses.
- The original uploaded file is not changed.
- The app is designed for normal Excel and CSV files up to about 50,000 rows.
