use anyhow::Result;
use std::path::PathBuf;

use crate::excel::Workbook;

pub enum InputMode {
    Normal,
    Editing,
    Goto,
    Confirm,
}

pub struct AppState {
    pub workbook: Workbook,
    pub file_path: PathBuf,
    pub selected_cell: (usize, usize),  // (row, col)
    pub start_row: usize,
    pub start_col: usize,
    pub visible_rows: usize,
    pub visible_cols: usize,
    pub input_mode: InputMode,
    pub input_buffer: String,
    pub status_message: String,
    pub should_quit: bool,
    pub column_widths: Vec<usize>,  // Store width for each column
}

impl AppState {
    pub fn new(workbook: Workbook, file_path: PathBuf) -> Result<Self> {
        // Initialize default column widths
        let max_cols = workbook.get_current_sheet().max_cols;
        let default_width = 15;
        let column_widths = vec![default_width; max_cols + 1];

        Ok(Self {
            workbook,
            file_path,
            selected_cell: (1, 1),  // Excel uses 1-based indexing
            start_row: 1,
            start_col: 1,
            visible_rows: 30,  // Default values, will be adjusted based on window size
            visible_cols: 15,  // Default values, will be adjusted based on window size
            input_mode: InputMode::Normal,
            input_buffer: String::new(),
            status_message: String::new(),
            should_quit: false,
            column_widths,
        })
    }

    pub fn move_cursor(&mut self, delta_row: isize, delta_col: isize) {
        // Calculate new position
        let new_row = (self.selected_cell.0 as isize + delta_row).max(1) as usize;
        let new_col = (self.selected_cell.1 as isize + delta_col).max(1) as usize;

        // Update selected position
        self.selected_cell = (new_row, new_col);

        // Handle scrolling
        self.handle_scrolling();
    }

    fn handle_scrolling(&mut self) {
        // Handle row scrolling
        if self.selected_cell.0 < self.start_row {
            self.start_row = self.selected_cell.0;
        } else if self.selected_cell.0 >= self.start_row + self.visible_rows {
            self.start_row = self.selected_cell.0 - self.visible_rows + 1;
        }

        // Handle column scrolling with Excel-like behavior
        // First, check if the selected column is before the first visible column
        if self.selected_cell.1 < self.start_col {
            // If selected column is before visible area, make it the first visible column
            self.start_col = self.selected_cell.1;
        } else {
            // Calculate which columns are currently visible
            let mut visible_cols = Vec::new();
            let mut width_sum = 0;
            let mut last_fully_visible_col = self.start_col;
            let mut last_partially_visible_col = self.start_col;

            // Get an approximate available width (this will be refined in update_visible_area)
            let approx_available_width = self.visible_cols * 15; // Approximate width

            // Calculate which columns are visible
            for col in self.start_col.. {
                let col_width = self.get_column_width(col);

                // If we can fully fit this column
                if width_sum + col_width <= approx_available_width {
                    visible_cols.push(col);
                    width_sum += col_width;
                    last_fully_visible_col = col;
                    last_partially_visible_col = col;
                }
                // If we can at least partially fit this column (Excel-like behavior)
                else if width_sum < approx_available_width {
                    visible_cols.push(col);
                    last_partially_visible_col = col;
                    break;
                } else {
                    break;
                }
            }

            // If selected column is not visible, adjust start_col
            if self.selected_cell.1 > last_partially_visible_col {
                // When moving right to a column that's not visible,
                // make it the first visible column
                self.start_col = self.selected_cell.1;
            }
            // If selected column is the last partially visible column and not fully visible,
            // adjust to make it more visible
            else if self.selected_cell.1 == last_partially_visible_col &&
                    self.selected_cell.1 > last_fully_visible_col {
                // If we're on a partially visible column, make it the first column
                // This ensures we can see as much of it as possible
                self.start_col = self.selected_cell.1;
            }
        }
    }

    pub fn get_cell_content(&self, row: usize, col: usize) -> String {
        let sheet = self.workbook.get_current_sheet();

        if row < sheet.data.len() && col < sheet.data[0].len() {
            let cell = &sheet.data[row][col];
            if cell.is_formula {
                format!("Formula: {}", cell.value)
            } else {
                cell.value.clone()
            }
        } else {
            String::new()
        }
    }

    pub fn start_editing(&mut self) {
        self.input_mode = InputMode::Editing;
        self.input_buffer = self.get_cell_content(
            self.selected_cell.0,
            self.selected_cell.1
        );
    }

    pub fn start_goto(&mut self) {
        self.input_mode = InputMode::Goto;
        self.input_buffer = String::new();
    }

    pub fn cancel_input(&mut self) {
        self.input_mode = InputMode::Normal;
        self.input_buffer = String::new();
    }

    pub fn confirm_edit(&mut self) -> Result<()> {
        if let InputMode::Editing = self.input_mode {
            self.workbook.set_cell_value(
                self.selected_cell.0,
                self.selected_cell.1,
                self.input_buffer.clone()
            )?;
            self.input_mode = InputMode::Normal;
            self.input_buffer = String::new();
        }
        Ok(())
    }

    pub fn confirm_goto(&mut self) {
        if let InputMode::Goto = self.input_mode {
            // Parse cell reference, e.g. A1, B2, etc.
            if let Some(cell_ref) = parse_cell_reference(&self.input_buffer) {
                self.selected_cell = cell_ref;
                self.handle_scrolling();
            } else {
                self.status_message = "Invalid cell reference".to_string();
            }
            self.input_mode = InputMode::Normal;
            self.input_buffer = String::new();
        }
    }

    pub fn add_char_to_input(&mut self, c: char) {
        self.input_buffer.push(c);
    }

    pub fn delete_char_from_input(&mut self) {
        self.input_buffer.pop();
    }

    pub fn exit(&mut self) {
        if self.workbook.is_modified() {
            self.input_mode = InputMode::Confirm;
            self.status_message = "File modified. Save? (y)es/(n)o/(c)ancel".to_string();
        } else {
            self.should_quit = true;
        }
    }

    pub fn save_and_exit(&mut self) {
        // Try to save the file
        match self.workbook.save() {
            Ok(_) => {
                self.status_message = "File saved".to_string();
                self.should_quit = true;
            },
            Err(e) => {
                self.status_message = format!("Save failed: {}", e);
                self.input_mode = InputMode::Normal;
            }
        }
    }

    pub fn exit_without_saving(&mut self) {
        self.should_quit = true;
    }

    pub fn cancel_exit(&mut self) {
        self.input_mode = InputMode::Normal;
        self.status_message = String::new();
    }

    pub fn auto_adjust_column_width(&mut self, col: Option<usize>) {
        let sheet = self.workbook.get_current_sheet();
        let default_min_width = 5;

        match col {
            // Adjust specific column
            Some(column) => {
                if column < self.column_widths.len() {
                    let width = self.calculate_column_width(column);
                    self.column_widths[column] = width.max(default_min_width);
                    self.status_message = format!("Column {} width adjusted", index_to_col_name(column));
                }
            },
            // Adjust all columns
            None => {
                for col_idx in 1..=sheet.max_cols {
                    let width = self.calculate_column_width(col_idx);
                    self.column_widths[col_idx] = width.max(default_min_width);
                }
                self.status_message = "All column widths adjusted".to_string();
            }
        }
    }

    fn calculate_column_width(&self, col: usize) -> usize {
        let sheet = self.workbook.get_current_sheet();
        let mut max_width = 3; // Minimum width

        // Calculate header width
        let col_name = index_to_col_name(col);
        max_width = max_width.max(col_name.len());

        // Check width of each cell content
        for row in 1..=sheet.max_rows {
            if row < sheet.data.len() && col < sheet.data[row].len() {
                let content = &sheet.data[row][col].value;

                // Calculate display width with better handling for different character types
                let display_width = content.chars().fold(0, |acc, c| {
                    // More accurate width calculation:
                    // - CJK characters: 2 units wide
                    // - ASCII letters/numbers: 1 unit wide
                    // - Other characters may need special handling
                    if c.is_ascii() {
                        // All ASCII characters count as 1 unit wide in the base calculation
                        acc + 1
                    } else {
                        // CJK characters are double-width
                        acc + 2
                    }
                });

                // For English text, add extra width to ensure it fits completely
                let adjusted_width = if content.chars().all(|c| c.is_ascii()) {
                    // For pure English text, add extra space to ensure it fits
                    // The factor 1.3 accounts for proportional font spacing in terminals
                    // and wider characters like 'W', 'M', etc.
                    (display_width as f32 * 1.3).ceil() as usize
                } else {
                    display_width
                };

                max_width = max_width.max(adjusted_width);
            }
        }

        // Add padding to ensure content fits
        // Use different padding for different content types
        if max_width > 20 {
            // For very wide content, less padding is needed proportionally
            max_width + 3
        } else {
            // For narrower content, more padding helps readability
            max_width + 4
        }
    }

    pub fn get_column_width(&self, col: usize) -> usize {
        if col < self.column_widths.len() {
            self.column_widths[col]
        } else {
            15 // Default width
        }
    }
}

// Parse cell reference, e.g. A1, B2, etc.
fn parse_cell_reference(input: &str) -> Option<(usize, usize)> {
    // Simple regex pattern matching
    let mut col_str = String::new();
    let mut row_str = String::new();

    // Separate column and row
    for c in input.chars() {
        if c.is_alphabetic() {
            col_str.push(c.to_ascii_uppercase());
        } else if c.is_numeric() {
            row_str.push(c);
        } else {
            return None;
        }
    }

    if col_str.is_empty() || row_str.is_empty() {
        return None;
    }

    // Convert column name to index
    let col = col_name_to_index(&col_str)?;

    // Convert row to index
    let row = row_str.parse::<usize>().ok()?;

    Some((row, col))
}

// Convert column name to index, e.g. A->1, B->2, AA->27, etc.
fn col_name_to_index(col_name: &str) -> Option<usize> {
    let mut index = 0;

    for c in col_name.chars() {
        if !c.is_ascii_alphabetic() {
            return None;
        }

        index = index * 26 + (c as usize - 'A' as usize + 1);
    }

    Some(index)
}

// Convert column index to column name (e.g. 1->A, 27->AA)
pub fn index_to_col_name(index: usize) -> String {
    let mut name = String::new();
    let mut idx = index;

    while idx > 0 {
        // Convert to 0-based for calculation
        idx -= 1;
        let remainder = idx % 26;
        name.insert(0, (b'A' + remainder as u8) as char);
        idx /= 26;
    }

    if name.is_empty() {
        name = "A".to_string();
    }

    name
}