# Excel-CLI

A lightweight terminal-based Excel viewer with Vim-like navigation for viewing, editing, and exporting Excel data to JSON format.

English | [中文](README_zh.md)

## Features

- Browse and navigate Excel worksheets with Vim-like hotkeys
- Switch between sheets in multi-sheet workbooks
- Edit cell contents directly in the terminal
- Export data to JSON format
- Delete rows, columns, and sheets
- Search functionality with highlighting
- Command mode for advanced operations

## Installation & Uninstallation

### Installation

#### Option 1: Install from Cargo (Recommended)

The package is published to crates.io and can be installed directly using:

```bash
cargo install excel-cli --locked
```

#### Option 2: Download from GitHub Release

1. Visit the [GitHub Releases](https://github.com/fuhan666/excel-cli/releases)
2. Download the pre-compiled binary for your operating system
3. Place the executable in your system path, or run it directly from the download location

Linux and macOS users may need to add execute permissions first

#### Option 3: Compile from Source

Requires Rust and Cargo. Install using the following commands:

```bash
# Clone the repository
git clone https://github.com/fuhan666/excel-cli.git
cd excel-cli
cargo build --release

# Install to system
cargo install --path . --locked
```

### Uninstallation

```bash
cargo uninstall excel-cli
```

## Usage

```bash
# Open an Excel file in interactive mode
excel-cli path/to/your/file.xlsx

# Export all sheets to JSON and output to stdout (for piping)
excel-cli path/to/your/file.xlsx --json-export

# Export with custom header direction and count
excel-cli path/to/your/file.xlsx -j -d v -r 2

# Pipe JSON output to another command
excel-cli path/to/your/file.xlsx -j > data.json # (example) Save JSON output to a file
```

### Command-line Options

- `--json-export`, `-j`: Export all sheets to JSON and output to stdout (for piping)
- `--direction`, `-d`: Header direction in Excel: 'h' for horizontal (top rows), 'v' for vertical (left columns). Default: 'h'
- `--header-count`, `-r`: Number of header rows (for horizontal) or columns (for vertical) in Excel. Default: 1
- `--lazy-loading`, `-l`: Enable lazy loading for large Excel files (only loads data when needed)

## User Interface

The application has a simple and intuitive interface:

- **Title Bar with Sheet Tabs**: Displays the current file name and all available sheets with the current sheet highlighted
- **Spreadsheet**: The main area displaying the Excel data
- **Content Panel**: Displays the full content of the currently selected cell
- **Notification Panel**: Displays operation feedback and system notifications
- **Status Bar**: Displays operation hints and current input commands

## Keyboard Shortcuts

- `h`, `j`, `k`, `l` or arrow keys: Move between cells (1 cell)
- `[`: Switch to previous sheet (stops at first sheet)
- `]`: Switch to next sheet (stops at last sheet)
- `0`: Jump to first column in current row
- `^`: Jump to first non-empty column in current row
- `$`: Jump to last column in current row
- `gg`: Jump to first row in current column
- `G`: Jump to last row in current column
- `Ctrl+←` (or `Command+←` on Mac): If current cell is empty, jump to the first non-empty cell to the left; if current cell is not empty, jump to the last non-empty cell to the left
- `Ctrl+→` (or `Command+→` on Mac): If current cell is empty, jump to the first non-empty cell to the right; if current cell is not empty, jump to the last non-empty cell to the right
- `Ctrl+↑` (or `Command+↑` on Mac): If current cell is empty, jump to the first non-empty cell above; if current cell is not empty, jump to the last non-empty cell above
- `Ctrl+↓` (or `Command+↓` on Mac): If current cell is empty, jump to the first non-empty cell below; if current cell is not empty, jump to the last non-empty cell below
- `Enter`: Edit current cell
- `y`: Copy current cell content
- `d`: Cut current cell content
- `p`: Paste clipboard content to current cell
- `u`: Undo the last operation (edit, row/column/sheet deletion)
- `Ctrl+r`: Redo the last undone operation
- `/`: Start forward search
- `?`: Start backward search
- `n`: Jump to next search result
- `N`: Jump to previous search result
- `:`: Enter command mode (for Vim-like commands)

## Vim Edit Mode

When editing cell content (press `Enter` to enter edit mode):

- **Mode Switching**:

  - `Esc`: Exit Vim mode and save changes
  - `i`: Enter Insert mode
  - `v`: Enter Visual mode

- **Navigation (in Normal mode)**:

  - `h`, `j`, `k`, `l`: Move cursor left, down, up, right
  - `w`: Move to next word
  - `b`: Move to beginning of word
  - `e`: Move to end of word
  - `$`: Move to end of line
  - `^`: Move to first non-blank character of line
  - `gg`: Move to first line
  - `G`: Move to last line

- **Editing Operations**:

  - `x`: Delete character under cursor
  - `D`: Delete to end of line
  - `C`: Change to end of line
  - `o`: Open new line below and enter Insert mode
  - `O`: Open new line above and enter Insert mode
  - `A`: Append at end of line
  - `I`: Insert at beginning of line

- **Visual Mode Operations**:

  - `y`: Yank (copy) selected text
  - `d`: Delete selected text
  - `c`: Change selected text (delete and enter Insert mode)

- **Operator Commands**:

  - `y{motion}`: Yank text specified by motion
  - `d{motion}`: Delete text specified by motion
  - `c{motion}`: Change text specified by motion

- **Clipboard Operations**:
  - `p`: Paste yanked or deleted text
  - `u`: Undo last change
  - `Ctrl+r`: Redo last undone change

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

- `:ej [h|v] [rows]` - Export current sheet data to JSON format

  - `h|v` - Header direction: `h` for horizontal (top rows), `v` for vertical (left columns)
  - `rows` - Number of header rows (for horizontal) or columns (for vertical)

- `:eja [h|v] [rows]` - Export all sheets to a single JSON file
  - Uses the same parameters as `:ej`
  - Creates a JSON object with sheet names as keys and sheet data as values

The output filename is automatically generated in one of these formats:

- For single sheet: `original_filename_sheet_SheetName_YYYYMMDD_HHMMSS.json`
- For all sheets: `original_filename_all_sheets_YYYYMMDD_HHMMSS.json`

The JSON files are saved in the same directory as the original Excel file.

### Vim-like Commands

- `:w` - Save file without exiting
- `:wq` or `:x` - Save and exit
- `:q` - Quit (will warn if there are unsaved changes)
- `:q!` - Force quit without saving
  See [File Saving Logic](#file-saving-logic) for details on how files are saved.

- `:y` - Copy current cell content
- `:d` - Cut current cell content
- `:put` or `:pu` - Paste clipboard content to current cell
- `:[cell]` - Jump to cell (e.g., `:A1`, `:B10`). Supports both uppercase and lowercase letters (`:a1` works the same as `:A1`)

### Sheet Management Commands

- `:sheet [name/number]` - Switch to sheet by name or index (1-based)
- `:delsheet` - Delete the current sheet

### Row and Column Management Commands

- `:dr` - Delete the current row
- `:dr [row]` - Delete a specific row (e.g., `:dr 5` deletes row 5)
- `:dr [start] [end]` - Delete a range of rows (e.g., `:dr 5 10` deletes rows 5 through 10)
- `:dc` - Delete the current column
- `:dc [col]` - Delete a specific column (e.g., `:dc A` or `:dc a` or `:dc 1` all delete column A)
- `:dc [start] [end]` - Delete a range of columns (e.g., `:dc A C` or `:dc a c` deletes columns A through C)

### Other Commands

- `:nohlsearch` or `:noh` - Disable search highlighting
- `:help` - Show available commands

## File Saving Logic

Excel-CLI uses a non-destructive approach to file saving:

- When you save a file (using `:w`, `:wq`, or `:x`), the application checks if any changes have been made
- If no changes have been made, no new file is created, and a "No changes to save" message is displayed
- If changes have been made, a new file is created with a timestamp in the filename, following the format `original_filename_YYYYMMDD_HHMMSS.xlsx`
- The new file is created without any styling
- The original file is never modified

## Technical Stack

- Written in Rust
- Uses ratatui library for terminal UI
- crossterm for terminal input handling
- calamine library for reading Excel files
- rust_xlsxwriter for writing Excel files
- serde_json for JSON serialization

## License

MIT
