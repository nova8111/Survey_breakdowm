# Audit fixes — October 2026

## Implemented

- Keep sheet selection inside Charts; keep report primary/secondary controls inside report options; group source status above the title.
- Preserve chart work per sheet and report choices per source/sheet; recover from empty sheets; validate replacement uploads before clearing current analysis; Clear data removes all loaded sources.
- Share normalization and counting rules; deduplicate per respondent, use union counts for merged/Other categories, avoid missing-label/header collisions, retain unexpected scale answers, and preserve source row numbers.
- Use answered respondent percentages with explicit bases/missingness; avoid misleading overlapping pie/100% stacked charts; make binary conversion opt-in; offer explicit delimiters and All responses.
- Add searchable report filter values, select/clear controls, multiple AND filters, a large-category override, and predictable question selection controls.
- Mark changed/failed reports stale and block exports; calculate reports in a cancellable worker with progress; destroy chart instances before replacement and update individual cards for settings changes.
- Fix linked primary context, exact/normalized matching, ambiguous IDs, separator consistency, matched-category unions, and diagnostics/exports using the correct workbook columns.
- Export source/filter/base metadata, selected count/percentage modes, readable CSV percentages, numeric Excel percentages, preserved integer bases, Unicode CSV, and formula-safe text.
- Improve CSV parsing/encoding errors, native expanded dialogs and focus return, keyboard tabs, mobile layout, accessible chart/table descriptions, and summary paging.
- Bundle pinned dependencies/licenses/checksums; copy worker/vendor/binary assets in builds; constrain script/network sources; split counting, import/export, and report worker modules; add integrated regressions and build verification.

## Validation

Automated tests cover calculations, import/export, linking, actual chart wrapper integration, report filter intersection, Excel roundtrip formatting and literal formula text, plus root/prefixed binary build routes. Browser checks used a synthetic multi-sheet workbook: sheet/tab isolation, chart preservation, mixed-case groups, filters, linked coverage, stale exports, expanded preview Escape/focus, empty-sheet recovery, failed replacement preservation, and mobile overflow. A 50,000-row/15-group report-section calculation took about 505 ms in the local Node benchmark; browser/render timings will vary.

## Remaining enhancements and limits

- Saved/shareable configuration files and per-question separator overrides are future features; current settings are session-scoped, with per-card chart separators and report-wide/secondary separators.
- Raw preview still shows 100 matching rows. Chart summary tables page; very wide reports and high-cardinality matrices may still be expensive.
- Excel formulas require cached values; the app does not calculate formulas. Weighting, statistical significance, and imputation are outside the descriptive reporting scope.
- Live public Google Sheets access was not exercised during this implementation; local upload, parser, worker, and build paths were tested.
- The CI template is included under `ci/` because the authenticated GitHub token lacks workflow permission. It can be enabled by moving it into `.github/workflows/` with an appropriately scoped credential.
- Implementation used the primary Codex agent plus three GPT-5.6 Luna workers with advisory routing; a full historical routing audit was not verified.
