# Sheet-CLI

A terminal-based Excel viewer and editor written in Rust, offering a smooth operating experience.

## Features

- Browse Excel worksheets
- Switch between sheets in multi-sheet workbooks
- Delete worksheets from multi-sheet workbooks
- Navigate cells using hjkl or arrow keys
- Edit cell contents
- Jump to specific cells using Vim-style commands
- View and create formulas
- Save changes to Excel files
- Export data to JSON format

## Installation

Requires Rust and Cargo. Install using the following commands:

```bash
# Compile from source
git clone https://github.com/yourusername/sheet-cli.git
cd sheet-cli
cargo build --release

# Install to system
cargo install --path .
```

## Usage

```bash
# Open an Excel file
sheet-cli path/to/your/file.xlsx
```

## User Interface

The application has a simple and intuitive interface:

- **Title Bar with Sheet Tabs**: Displays the current file name and all available sheets with the current sheet highlighted
- **Spreadsheet**: The main area displaying the Excel data
- **Status Bar**: Shows current cell reference and available commands

## Keyboard Shortcuts

- `h`, `j`, `k`, `l` or arrow keys: Move selection (1 cell)
- `[`: Switch to previous sheet (stops at first sheet)
- `]`: Switch to next sheet (stops at last sheet)
- `0`: Jump to first column in current row
- `^`: Jump to first non-empty column in current row
- `$`: Jump to last column in current row
- `gg`: Jump to first row in current column
- `G`: Jump to last row in current column
- `Ctrl+←`: If current cell is empty, jump to the first non-empty cell to the left; if current cell is not empty, jump to the last non-empty cell to the left
- `Ctrl+→`: If current cell is empty, jump to the first non-empty cell to the right; if current cell is not empty, jump to the last non-empty cell to the right
- `Ctrl+↑`: If current cell is empty, jump to the first non-empty cell above; if current cell is not empty, jump to the last non-empty cell above
- `Ctrl+↓`: If current cell is empty, jump to the first non-empty cell below; if current cell is not empty, jump to the last non-empty cell below
- `i`: Edit current cell
- `y`: Copy current cell content
- `d`: Cut current cell content
- `p`: Paste clipboard content to current cell
- `/`: Start forward search
- `?`: Start backward search
- `n`: Jump to next search result
- `N`: Jump to previous search result
- `:`: Enter command mode (for Vim-style commands)

## Edit Mode

In edit mode:

- `Enter`: Confirm edit
- `Esc`: Cancel edit
- Formulas can be entered by starting with `=`

## Search Mode

Enter search mode by pressing `/` (forward search) or `?` (backward search):

- Type your search query
- `Enter`: Execute search and jump to the first match
- `Esc`: Cancel search
- `n`: Jump to next match (after search is executed)
- `N`: Jump to previous match (after search is executed)
- Search results are highlighted in yellow
- Search uses row-first, column-second order (searches through each row from left to right, then moves to the next row)

## Command Mode

Enter command mode by pressing `:`. Available commands:

### Column Width Commands

- `:cw fit` - Auto-adjust current column width to fit content
- `:cw fit all` - Auto-adjust all column widths to fit content
- `:cw min` - Minimize current column width (max 15 or content width)
- `:cw min all` - Minimize all column widths (max 15 or content width)
- `:cw [number]` - Set current column width to specified value

### JSON Export Commands

- `:ej [h|v] [rows]` - Export data to JSON format
  - `h|v` - Header direction: `h` for horizontal (top rows), `v` for vertical (left columns)
  - `rows` - Number of header rows (for horizontal) or columns (for vertical)

The output filename is automatically generated in the format: `original_filename_sheet_SheetName_YYYYMMDD_HHMMSS.json`

### Vim-style Commands

- `:w` - Save file without exiting
- `:wq` or `:x` - Save and exit
- `:q` - Quit (will warn if there are unsaved changes)
- `:q!` - Force quit without saving
See [File Saving Logic](#file-saving-logic) for details on how files are saved.

- `:y` - Copy current cell content
- `:d` - Cut current cell content
- `:put` or `:pu` - Paste clipboard content to current cell
- `:[cell]` - Jump to cell (e.g., `:A1`, `:B10`). Supports both uppercase and lowercase column letters (`:a1` works the same as `:A1`)

### Sheet Management Commands

- `:sheet [name/number]` - Switch to sheet by name or index (1-based)
- `:delsheet` - Delete the current sheet (prevents deleting the last sheet)

### Other Commands

- `:nohlsearch` or `:noh` - Disable search highlighting
- `:help` - Show available commands

## File Saving Logic

Sheet-CLI uses a non-destructive approach to file saving:

- When you save a file (using `:w`, `:wq`, or `:x`), the application checks if any changes have been made.
- If no changes have been made, no new file is created, and a "No changes to save" message is displayed.
- If changes have been made, a new file is created with a timestamp in the filename, following the format `original_filename_YYYYMMDD_HHMMSS.xlsx`.
- The original file is never modified.

## Technical Stack

- Written in Rust
- Uses ratatui library for terminal UI
- crossterm for terminal input handling
- calamine library for reading Excel files
- rust_xlsxwriter for writing Excel files
- serde_json for JSON serialization

## License

MIT
