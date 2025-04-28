# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- Support dimming other areas while editing

### Changed

- Edit cell content using vim shortcuts
- Multiple UI improvements

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

[Unreleased]: https://github.com/fuhan666/excel-cli/compare/v0.2.0...HEAD
[0.2.0]: https://github.com/fuhan666/excel-cli/releases/tag/v0.2.0
[0.1.1]: https://github.com/fuhan666/excel-cli/releases/tag/v0.1.1
[0.1.0]: https://github.com/fuhan666/excel-cli/releases/tag/v0.1.0
