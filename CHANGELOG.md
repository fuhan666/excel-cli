# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- `excel-cli grep` command for recursive search across Excel files, with Markdown table (default) and JSON output.
- Markdown table output for `read rows`, `read records`, `inspect sample`, `read range`, and `grep` commands.
- `--skip-errors` option for `grep` command to skip worksheets that cannot be read instead of returning an error.

### Changed

- `grep` defaults to Markdown output and supports `-f markdown` and `-f json`; text output (`-f text`) is no longer supported.
- `--output-shape jsonl` now rejects `--format text` and `--format markdown`; use the default JSON format or `-f json`.
- Lazy-loaded worksheet read errors now return an error instead of silently skipping the sheet (unless `--skip-errors` is used).

## [1.3.2] - 2026-05-28

### Added

- `:freeze` and `:unfreeze` commands to freeze and unfreeze rows and columns at the current or specified cell.

## [1.3.1] - 2026-04-24

### Added

- TUI `:help` command — opens a scrollable, structured key-reference overlay with all navigation, editing, and command-mode shortcuts.

### Changed

- TUI visual refresh: cleaner title bar with branding and inline row/column stats, single-line status bar, and a consistent dark theme across panels.
- `check` scans are more efficient and report findings with A1-style cell ranges.

## [1.3.0] - 2026-04-22

### Added

- `check <file>` workbook and sheet quality scans with a stable JSON summary, stats, and findings contract.
- Fixed v1.3.0 rule registry covering `blank_headers`, `duplicate_headers`, `blank_rows`, `blank_columns`, `null_ratio`, `duplicate_values`, `type_drift`, and `formula_presence`.
- TUI findings panel support for browsing quality-check results without leaving the terminal.
- Regression coverage for rule positives and negatives, workbook and sheet targeting, threshold filtering, finding order, and exit-code behavior.

## [1.2.0] - 2026-04-21

### Added

- `read rows` column selection with `--select <col1,col2>`.
- `read rows` and `read records` filtering with repeated `--filter field:op:value` conditions combined with AND semantics.
- Filter operators for equality, inequality, numeric comparisons, string contains, regular expressions, null checks, and non-null checks.
- Pagination controls with `--limit` and `--offset`.
- `--non-empty` to drop all-empty rows from read results.
- `read records <file>` for header-keyed records by default.
- `--output-shape rows|records|jsonl` on row and record reads.
- JSON Lines output for stream-friendly record processing.
- Read-only TUI query preview with `:preview` and `:pv`.
- Regression coverage for filtering, output shapes, invalid query errors, no-match results, help text, and query preview behavior.

### Changed

- Enriched read metadata now reports applied filters, selected columns, returned row count, truncation, and output shape.

## [1.1.0] - 2026-04-21

### Added

- `inspect columns <file> --sheet <name>` to profile column headers, generated safe names, duplicate headers, inferred types, non-null ratios, formula ratios, and sample values.
- `inspect tables <file> --sheet <name>` to detect table-like regions with ranges, header rows, dimensions, and confidence scores.
- Regression coverage for structure inspection cases including duplicate headers, blank headers, preamble sections, late headers, multi-table sheets, mixed-type columns, formula columns, and non-ASCII column names.

## [1.0.0] - 2026-04-20

### Added

- New explicit subcommand architecture: `inspect`, `read`, `check`, `ui`
  - `inspect workbook <file>` — list all sheets with metadata
  - `inspect sheet <file> --sheet <name>` / `--sheet-index <n>` — sheet metadata
  - `inspect sample <file> --sheet <name> [--range] [--rows]` — sampled data
  - `read cell <file> --sheet <name> --cell <A1>` — single cell value
  - `read range <file> --sheet <name> --range <A1:F20>` — rectangular range
  - `read rows <file> --sheet <name> [--range] [--header-row auto|N]` — row-oriented data
  - `ui <file>` — explicit interactive TUI entry (replaces bare file path)
  - `check <file> --rule <rule>` — namespace reserved for v1.3.0 quality checks
- Unified success/error JSON envelope with stable schema
  - Success: `schema_version`, `command`, `file`, `target`, `meta`, `data`, `warnings`
  - Error: `schema_version`, `command`, `file`, `error` with `code`, `message`, `details`
- Stable exit code taxonomy: 0 (success), 1 (check findings), 2 (invalid args), 3 (file error), 4 (parse error), 5 (target not found), 6 (invalid query), 7 (internal error)
- `--sheet` and `--sheet-index` are now independent and explicit (no ambiguity with numeric sheet names)
- `--format json|text` with `json` as the default for all headless commands
- Header row auto-detection with `header_candidates` and `recommended_header_row` in `inspect sheet`
- New internal `src/cli/` module for command parsing, dispatch, envelope, and error handling

### Changed

- CLI entry is now subcommand-based; bare file path is rejected with exit code 2
- `json` is the default headless output format (was `text`)
- `stdout` only contains results; `stderr` only contains errors
- All headless outputs use the unified envelope structure

### Removed

- Old positional-file-plus-flags public CLI surface (`--sheets`, `--peek`, `--cell`, `--json-export`, `--direction`, `--header-count`)
- No backward-compatibility layer for old flags

## [0.5.2] - 2026-04-20

### Fixed

- Remove redundant `.max(0)` calls after `saturating_sub` to satisfy clippy's `unnecessary_min_or_max` lint without changing behavior.

## [0.5.1] - 2026-04-17

### Fixed

- Prevent process panic when opening malformed XLSX files with invalid shared-string references. Parser panics from `calamine` are now caught and converted into controlled `anyhow` errors with a non-zero exit code.

## [0.5.0] - 2026-04-17

### Added

- Added AI-friendly headless inspect/query commands:
  - `--sheets` to list all sheets
  - `--sheet <sheet>` to show sheet metadata
  - `--peek <sheet>!<range>` to preview a cell range
  - `--cell <sheet>!<cell>` to read a single cell value
  - `--sheet <sheet> --json-export` to export a single sheet to JSON
- Added `--format <text|json>` flag for inspect commands
- Added A1 notation parsing helpers for range and cell references
- Added `Workbook::resolve_sheet` and `Workbook::get_used_range` utilities

### Changed

- `--json-export` now also supports `--sheet` to export a single sheet instead of the whole workbook
- Inspect commands output only data to stdout and errors to stderr with stable exit codes
- Out-of-bounds peek ranges are clamped to actual sheet bounds
- Empty cells output empty string in text mode and `null` in JSON mode

## [0.4.2] - 2026-03-10

### Fixed

- Prevent command parsing from panicking on non-ASCII sheet names such as `:addsheet test1`

## [0.4.1] - 2026-03-10

### Fixed

- Publish crates.io release with the correct crate version after the `v0.4.0` GitHub release

## [0.4.0] - 2026-03-10

### Added

- Added `:addsheet <name>` to create a new sheet after the current sheet
- Added undo and redo support for sheet creation

### Fixed

- Load all lazy-loaded sheets before saving to avoid writing unloaded sheets as blank
- Count sheet name length by characters so non-ASCII names such as Chinese sheet names are validated correctly

## [0.3.0] - 2025-05-07

### Added

- Support dimming other areas while editing
- Remember the selected cell when switching between sheets
- Added optional lazy loading for xlsx and xlsb files to improve performance with large files (enabled with -l flag)

### Fixed

- Fixed the issue where row numbers are not fully displayed when exceeding 100,000

### Changed

- Edit cell content using vim shortcuts
- Multiple UI improvements
- Replace ratatui_textarea with tui_textarea
- Upgraded calamine to version 0.27.0

## [0.2.0] - 2025-04-27

### Added

- Added restriction to require `-j` or `--json-export` flag when using pipe redirection
- Added undo and redo functionality, supporting operation history for cell editing, row/column deletion, and sheet deletion
- Added support for Command + arrow key navigation on Mac, similar to Ctrl + arrow key functionality

### Fixed

- Fixed the issue where cells outside the maximum range of sheet data cannot be edited
- No longer reports an error when deleting rows or columns outside the sheet's data range

### Changed

- Changed the cell content panel and notification panel to vertical layout
- Export JSON files to the same directory as the original Excel file
- Some minor UI adjustments
- Downgraded Rust edition version to 2021

## [0.1.1] - 2025-04-24

### Changed

- improve installation instructions in README

## [0.1.0] - 2025-04-24

This is the initial release of excel-cli, a lightweight terminal-based Excel viewer with Vim-like navigation.

### Added

- Browse and switch between worksheets in multi-sheet workbooks
- Delete worksheets from multi-sheet workbooks
- Delete column and row functionality
- Edit cell contents with a text editor
- Save changes to Excel files
- Export data to JSON format with customizable header options
- Vim-like commands for navigation and operations:
  - `h`, `j`, `k`, `l` for cell navigation
  - `0`, `^`, `$` for row navigation
  - `gg`, `G` for column navigation
  - `Ctrl+←→↑↓` for jumping to non-empty cells
  - `/`, `?` for searching with `n`, `N` for navigation between matches
  - `:` for command mode with various commands
- Copy, cut, and paste functionality with `y`, `d`, and `p` keys
- Support for pipe operator when exporting to JSON

[Unreleased]: https://github.com/fuhan666/excel-cli/compare/v1.3.2...HEAD
[1.3.2]: https://github.com/fuhan666/excel-cli/releases/tag/v1.3.2
[1.3.1]: https://github.com/fuhan666/excel-cli/releases/tag/v1.3.1
[1.3.0]: https://github.com/fuhan666/excel-cli/releases/tag/v1.3.0
[1.2.0]: https://github.com/fuhan666/excel-cli/releases/tag/v1.2.0
[1.1.0]: https://github.com/fuhan666/excel-cli/releases/tag/v1.1.0
[1.0.0]: https://github.com/fuhan666/excel-cli/releases/tag/v1.0.0
[0.5.2]: https://github.com/fuhan666/excel-cli/releases/tag/v0.5.2
[0.5.1]: https://github.com/fuhan666/excel-cli/releases/tag/v0.5.1
[0.5.0]: https://github.com/fuhan666/excel-cli/releases/tag/v0.5.0
[0.4.2]: https://github.com/fuhan666/excel-cli/releases/tag/v0.4.2
[0.4.1]: https://github.com/fuhan666/excel-cli/releases/tag/v0.4.1
[0.4.0]: https://github.com/fuhan666/excel-cli/releases/tag/v0.4.0
[0.3.0]: https://github.com/fuhan666/excel-cli/releases/tag/v0.3.0
[0.2.0]: https://github.com/fuhan666/excel-cli/releases/tag/v0.2.0
[0.1.1]: https://github.com/fuhan666/excel-cli/releases/tag/v0.1.1
[0.1.0]: https://github.com/fuhan666/excel-cli/releases/tag/v0.1.0
