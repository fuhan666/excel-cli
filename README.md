# Sheet-CLI

A terminal-based Excel viewer and editor written in Rust, offering a smooth operating experience.

## Features

- Browse Excel worksheets
- Navigate cells using hjkl or arrow keys
- Edit cell contents
- Jump to specific cells
- View and create formulas
- Save changes back to Excel files
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

## Keyboard Shortcuts

- `h`, `j`, `k`, `l` or arrow keys: Move selection (1 cell)
- `H`, `J`, `K`, `L`: Move selection (5 cells)
- `i`: Edit current cell
- `g`: Go to specific cell (enter cell reference like A1, B2, etc.)
- `y`: Copy current cell content
- `d`: Cut current cell content
- `p`: Paste clipboard content to current cell
- `:`: Enter command mode (for Vim-style commands)

## Edit Mode

In edit mode:

- `Enter`: Confirm edit
- `Esc`: Cancel edit
- Formulas can be entered by starting with `=`

## Goto Mode

In goto mode:

- Enter a cell reference (e.g., A1, B10, Z52)
- `Enter`: Confirm and jump to the cell
- `Esc`: Cancel

## Command Mode

Enter command mode by pressing `:`. Available commands:

### Column Width Commands

- `:cw fit` - Auto-adjust current column width to fit content
- `:cw fit all` - Auto-adjust all column widths to fit content
- `:cw min` - Minimize current column width (max 15 or content width)
- `:cw min all` - Minimize all column widths (max 15 or content width)
- `:cw [number]` - Set current column width to specified value

### JSON Export Commands

- `:export json [filename] [h|v] [rows]` - Export data to JSON format
  - `filename` - The name of the output JSON file
  - `h|v` - Header direction: `h` for horizontal (top rows), `v` for vertical (left columns)
  - `rows` - Number of header rows (for horizontal) or columns (for vertical)
- `:ej [filename] [h|v] [rows]` - Shorthand for export json command

### Vim-style Commands

- `:w` - Save file without exiting
- `:wq` or `:x` - Save and exit
- `:q` - Quit (will warn if there are unsaved changes)
- `:q!` - Force quit without saving
- `:y` - Copy current cell content
- `:d` - Cut current cell content
- `:put` or `:pu` - Paste clipboard content to current cell

### Other Commands

- `:help` - Show available commands

## Technical Stack

- Written in Rust
- Uses ratatui library for terminal UI
- crossterm for terminal input handling
- calamine library for reading Excel files
- simple_excel_writer for writing Excel files
- serde_json for JSON serialization

## License

MIT
