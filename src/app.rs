use anyhow::Result;
use ratatui_textarea::TextArea;
use std::collections::HashMap;
use std::path::{Path, PathBuf};

use crate::excel::Workbook;
use crate::json_export::{HeaderDirection, export_all_sheets_json, export_json};

pub enum InputMode {
    Normal,
    Editing,
    Command,
    SearchForward,
    SearchBackward,
    Help,
}

pub struct AppState<'a> {
    pub workbook: Workbook,
    pub file_path: PathBuf,
    pub selected_cell: (usize, usize), // (row, col)
    pub start_row: usize,
    pub start_col: usize,
    pub visible_rows: usize,
    pub visible_cols: usize,
    pub input_mode: InputMode,
    pub input_buffer: String,
    pub text_area: TextArea<'a>,
    pub should_quit: bool,
    pub column_widths: Vec<usize>, // Store width for current sheet's columns
    pub sheet_column_widths: HashMap<String, Vec<usize>>, // Store column widths for each sheet
    pub clipboard: Option<String>, // Store copied/cut cell content
    pub g_pressed: bool,           // Track if 'g' was pressed for 'gg' command
    pub search_query: String,      // Current search query
    pub search_results: Vec<(usize, usize)>, // List of cells matching the search query
    pub current_search_idx: Option<usize>, // Index of current search result
    pub search_direction: bool,    // true for forward, false for backward
    pub highlight_enabled: bool,   // Control whether search results are highlighted
    pub info_panel_height: usize,
    pub notification_messages: Vec<String>,
    pub max_notifications: usize,
    pub help_text: String,
    pub help_scroll: usize,
    pub help_visible_lines: usize,
}

impl<'a> AppState<'a> {
    pub fn new(workbook: Workbook, file_path: PathBuf) -> Result<Self> {
        // Initialize default column widths for current sheet
        let max_cols = workbook.get_current_sheet().max_cols;
        let default_width = 15;
        let column_widths = vec![default_width; max_cols + 1];

        // Initialize column widths for all sheets
        let mut sheet_column_widths = HashMap::new();
        let sheet_names = workbook.get_sheet_names();

        for (i, name) in sheet_names.iter().enumerate() {
            if i == workbook.get_current_sheet_index() {
                sheet_column_widths.insert(name.clone(), column_widths.clone());
            } else {
                let sheet_max_cols = if let Some(sheet) = workbook.get_sheet_by_index(i) {
                    sheet.max_cols
                } else {
                    max_cols // Fallback to current sheet's max_cols
                };
                sheet_column_widths.insert(name.clone(), vec![default_width; sheet_max_cols + 1]);
            }
        }

        // Initialize TextArea
        let text_area = TextArea::default();

        Ok(Self {
            workbook,
            file_path,
            selected_cell: (1, 1), // Excel uses 1-based indexing
            start_row: 1,
            start_col: 1,
            visible_rows: 30, // Default values, will be adjusted based on window size
            visible_cols: 15, // Default values, will be adjusted based on window size
            input_mode: InputMode::Normal,
            input_buffer: String::new(),
            text_area,
            should_quit: false,
            column_widths,
            sheet_column_widths,
            clipboard: None,
            g_pressed: false,
            search_query: String::new(),
            search_results: Vec::new(),
            current_search_idx: None,
            search_direction: true,  // Default to forward search
            highlight_enabled: true, // Default to showing highlights
            info_panel_height: 5,
            notification_messages: Vec::new(),
            max_notifications: 5,
            help_text: String::new(),
            help_scroll: 0,
            help_visible_lines: 20,
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
        if self.selected_cell.0 < self.start_row {
            self.start_row = self.selected_cell.0;
        } else if self.selected_cell.0 >= self.start_row + self.visible_rows {
            self.start_row = self.selected_cell.0 - self.visible_rows + 1;
        }

        self.handle_column_scrolling();
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

    pub fn add_notification(&mut self, message: String) {
        self.notification_messages.push(message);

        if self.notification_messages.len() > self.max_notifications {
            self.notification_messages.remove(0);
        }
    }

    pub fn adjust_info_panel_height(&mut self, delta: isize) {
        let new_height = (self.info_panel_height as isize + delta).max(3).min(15) as usize;
        if new_height != self.info_panel_height {
            self.info_panel_height = new_height;
            self.add_notification(format!("Info panel height: {}", self.info_panel_height));
        }
    }

    pub fn start_editing(&mut self) {
        self.input_mode = InputMode::Editing;
        let content = self.get_cell_content(self.selected_cell.0, self.selected_cell.1);
        self.input_buffer = content.clone();

        // Set up TextArea for editing
        self.text_area = TextArea::default();
        self.text_area.insert_str(&content);
    }

    pub fn cancel_input(&mut self) {
        // If in help mode, just close the help window
        if let InputMode::Help = self.input_mode {
            self.input_mode = InputMode::Normal;
            return;
        }

        // Otherwise, cancel the current input
        self.input_mode = InputMode::Normal;
        self.input_buffer = String::new();
        self.text_area = TextArea::default();
    }

    pub fn show_help(&mut self) {
        self.help_scroll = 0;

        self.help_text = format!(
            "FILE OPERATIONS:\n\
             :w          - Save file\n\
             :wq, :x     - Save and quit\n\
             :q          - Quit (will warn if unsaved changes)\n\
             :q!         - Force quit without saving\n\n\
             NAVIGATION:\n\
             :[cell]     - Jump to cell (e.g., :B10)\n\
             hjkl        - Move cursor (left, down, up, right)\n\
             0           - Jump to first column\n\
             ^           - Jump to first non-empty column\n\
             $           - Jump to last column\n\
             gg          - Jump to first row\n\
             G           - Jump to last row\n\
             Ctrl+arrows - Jump to next non-empty cell\n\
             [           - Switch to previous sheet\n\
             ]           - Switch to next sheet\n\
             :sheet [name/number] - Switch to sheet by name or index\n\n\
             EDITING:\n\
             i           - Edit current cell\n\
             :y          - Copy current cell\n\
             :d          - Cut current cell\n\
             :put, :pu   - Paste to current cell\n\n\
             SEARCH:\n\
             /           - Search forward\n\
             ?           - Search backward\n\
             n           - Jump to next search result\n\
             N           - Jump to previous search result\n\
             :nohlsearch, :noh - Disable search highlighting\n\n\
             COLUMN OPERATIONS:\n\
             :cw fit     - Adjust width of current column to fit its content\n\
             :cw fit all - Adjust width of all columns to fit their content\n\
             :cw min     - Set current column width to minimum (5 characters)\n\
             :cw min all - Set all columns width to minimum\n\
             :cw [number] - Set current column width to specific number of characters\n\
             :dc         - Delete current column\n\
             :dc [col]   - Delete specific column (e.g., :dc A or :dc 1)\n\
             :dc [start] [end] - Delete columns from start to end (e.g., :dc A C)\n\n\
             ROW OPERATIONS:\n\
             :dr         - Delete current row\n\
             :dr [row]   - Delete specific row\n\
             :dr [start] [end] - Delete rows from start to end\n\n\
             EXPORT:\n\
             :ej [h|v] [rows]  - Export current sheet to JSON\n\
             :eja [h|v] [rows] - Export all sheets to a single JSON file\n\
                                h=horizontal (default), v=vertical\n\
                                [rows]=number of header rows (default: 1)\n\n\
             SHEET OPERATIONS:\n\
             :delsheet   - Delete the current sheet\n\n\
             UI ADJUSTMENTS:\n\
             +/=         - Increase info panel height\n\
             -           - Decrease info panel height\n\n\
             HELP NAVIGATION:\n\
             j/k, Up/Down  - Scroll help text up/down\n\
             Enter/Esc     - Close help window"
        );

        self.input_mode = InputMode::Help;
    }

    pub fn confirm_edit(&mut self) -> Result<()> {
        if let InputMode::Editing = self.input_mode {
            // Get content from TextArea
            let content = self.text_area.lines().join("\n");

            self.workbook
                .set_cell_value(self.selected_cell.0, self.selected_cell.1, content)?;
            self.input_mode = InputMode::Normal;
            self.input_buffer = String::new();
            self.text_area = TextArea::default();
        }
        Ok(())
    }

    pub fn jump_to_first_row(&mut self) {
        let current_col = self.selected_cell.1;
        self.selected_cell = (1, current_col);
        self.handle_scrolling();
        self.add_notification("Jumped to first row".to_string());
    }

    pub fn jump_to_last_row(&mut self) {
        let sheet = self.workbook.get_current_sheet();
        let current_col = self.selected_cell.1;

        let max_row = sheet.max_rows;

        self.selected_cell = (max_row, current_col);
        self.handle_scrolling();
        self.add_notification("Jumped to last row".to_string());
    }

    pub fn jump_to_first_column(&mut self) {
        let current_row = self.selected_cell.0;
        self.selected_cell = (current_row, 1); // First column is 1 in our system
        self.handle_scrolling();
        self.add_notification("Jumped to first column".to_string());
    }

    pub fn jump_to_first_non_empty_column(&mut self) {
        let sheet = self.workbook.get_current_sheet();
        let current_row = self.selected_cell.0;

        let mut first_non_empty_col = 1; // Default to first column

        if current_row < sheet.data.len() {
            for col in 1..=sheet.max_cols {
                if col < sheet.data[current_row].len()
                    && !sheet.data[current_row][col].value.is_empty()
                {
                    first_non_empty_col = col;
                    break;
                }
            }
        }

        self.selected_cell = (current_row, first_non_empty_col);
        self.handle_scrolling();
        self.add_notification("Jumped to first non-empty column".to_string());
    }

    pub fn jump_to_last_column(&mut self) {
        let sheet = self.workbook.get_current_sheet();
        let current_row = self.selected_cell.0;

        let max_col = sheet.max_cols;

        self.selected_cell = (current_row, max_col);
        self.handle_scrolling();
        self.add_notification("Jumped to last column".to_string());
    }

    pub fn jump_to_prev_non_empty_cell_left(&mut self) {
        let sheet = self.workbook.get_current_sheet();
        let (row, col) = self.selected_cell;

        if col <= 1 {
            return;
        }

        let current_cell_is_empty = row >= sheet.data.len()
            || col >= sheet.data[row].len()
            || sheet.data[row][col].value.is_empty();

        if current_cell_is_empty {
            let mut target_col = 1;
            let mut found_non_empty = false;

            for c in (1..col).rev() {
                if row < sheet.data.len()
                    && c < sheet.data[row].len()
                    && !sheet.data[row][c].value.is_empty()
                {
                    target_col = c;
                    found_non_empty = true;
                    break;
                }
            }

            if !found_non_empty {
                target_col = 1;
            }

            self.selected_cell = (row, target_col);
            self.handle_scrolling();
            self.add_notification("Jumped to first non-empty cell (left)".to_string());
        } else {
            let mut target_col = 1;
            let mut last_non_empty_col = col;
            let mut found_empty_after_non_empty = false;

            for c in (1..col).rev() {
                if row < sheet.data.len() && c < sheet.data[row].len() {
                    if sheet.data[row][c].value.is_empty() {
                        target_col = c + 1;
                        found_empty_after_non_empty = true;
                        break;
                    } else {
                        last_non_empty_col = c;
                    }
                } else {
                    target_col = c + 1;
                    found_empty_after_non_empty = true;
                    break;
                }
            }

            if !found_empty_after_non_empty {
                target_col = last_non_empty_col;
            }

            self.selected_cell = (row, target_col);
            self.handle_scrolling();
            self.add_notification("Jumped to last non-empty cell (left)".to_string());
        }
    }

    pub fn jump_to_prev_non_empty_cell_right(&mut self) {
        let sheet = self.workbook.get_current_sheet();
        let (row, col) = self.selected_cell;
        let max_col = sheet.max_cols;

        if col >= max_col {
            return;
        }

        let current_cell_is_empty = row >= sheet.data.len()
            || col >= sheet.data[row].len()
            || sheet.data[row][col].value.is_empty();

        if current_cell_is_empty {
            let mut target_col = max_col;
            let mut found_non_empty = false;

            for c in (col + 1)..=max_col {
                if row < sheet.data.len()
                    && c < sheet.data[row].len()
                    && !sheet.data[row][c].value.is_empty()
                {
                    target_col = c;
                    found_non_empty = true;
                    break;
                }
            }

            if !found_non_empty {
                target_col = max_col;
            }

            self.selected_cell = (row, target_col);
            self.handle_scrolling();
            self.add_notification("Jumped to first non-empty cell (right)".to_string());
        } else {
            let mut target_col = max_col;
            let mut last_non_empty_col = col;
            let mut found_empty_after_non_empty = false;

            for c in (col + 1)..=max_col {
                if row < sheet.data.len() && c < sheet.data[row].len() {
                    if sheet.data[row][c].value.is_empty() {
                        target_col = c - 1;
                        found_empty_after_non_empty = true;
                        break;
                    } else {
                        last_non_empty_col = c;
                    }
                } else {
                    target_col = c - 1;
                    found_empty_after_non_empty = true;
                    break;
                }
            }

            if !found_empty_after_non_empty {
                target_col = last_non_empty_col;
            }

            self.selected_cell = (row, target_col);
            self.handle_scrolling();
            self.add_notification("Jumped to last non-empty cell (right)".to_string());
        }
    }

    pub fn jump_to_prev_non_empty_cell_up(&mut self) {
        let sheet = self.workbook.get_current_sheet();
        let (row, col) = self.selected_cell;

        if row <= 1 {
            return;
        }

        let current_cell_is_empty = row >= sheet.data.len()
            || col >= sheet.data[row].len()
            || sheet.data[row][col].value.is_empty();

        if current_cell_is_empty {
            let mut target_row = 1;
            let mut found_non_empty = false;

            for r in (1..row).rev() {
                if r < sheet.data.len()
                    && col < sheet.data[r].len()
                    && !sheet.data[r][col].value.is_empty()
                {
                    target_row = r;
                    found_non_empty = true;
                    break;
                }
            }

            if !found_non_empty {
                target_row = 1;
            }

            self.selected_cell = (target_row, col);
            self.handle_scrolling();
            self.add_notification("Jumped to first non-empty cell (up)".to_string());
        } else {
            let mut target_row = 1;
            let mut last_non_empty_row = row;
            let mut found_empty_after_non_empty = false;

            for r in (1..row).rev() {
                if r < sheet.data.len() && col < sheet.data[r].len() {
                    if sheet.data[r][col].value.is_empty() {
                        target_row = r + 1;
                        found_empty_after_non_empty = true;
                        break;
                    } else {
                        last_non_empty_row = r;
                    }
                } else {
                    target_row = r + 1;
                    found_empty_after_non_empty = true;
                    break;
                }
            }

            if !found_empty_after_non_empty {
                target_row = last_non_empty_row;
            }

            self.selected_cell = (target_row, col);
            self.handle_scrolling();
            self.add_notification("Jumped to last non-empty cell (up)".to_string());
        }
    }

    pub fn jump_to_prev_non_empty_cell_down(&mut self) {
        let sheet = self.workbook.get_current_sheet();
        let (row, col) = self.selected_cell;
        let max_row = sheet.max_rows;

        if row >= max_row {
            return;
        }

        let current_cell_is_empty = row >= sheet.data.len()
            || col >= sheet.data[row].len()
            || sheet.data[row][col].value.is_empty();

        if current_cell_is_empty {
            let mut target_row = max_row;
            let mut found_non_empty = false;

            for r in (row + 1)..=max_row {
                if r < sheet.data.len()
                    && col < sheet.data[r].len()
                    && !sheet.data[r][col].value.is_empty()
                {
                    target_row = r;
                    found_non_empty = true;
                    break;
                }
            }

            if !found_non_empty {
                target_row = max_row;
            }

            self.selected_cell = (target_row, col);
            self.handle_scrolling();
            self.add_notification("Jumped to first non-empty cell (down)".to_string());
        } else {
            let mut target_row = max_row;
            let mut last_non_empty_row = row;
            let mut found_empty_after_non_empty = false;

            for r in (row + 1)..=max_row {
                if r < sheet.data.len() && col < sheet.data[r].len() {
                    if sheet.data[r][col].value.is_empty() {
                        target_row = r - 1;
                        found_empty_after_non_empty = true;
                        break;
                    } else {
                        last_non_empty_row = r;
                    }
                } else {
                    target_row = r - 1;
                    found_empty_after_non_empty = true;
                    break;
                }
            }

            if !found_empty_after_non_empty {
                target_row = last_non_empty_row;
            }

            self.selected_cell = (target_row, col);
            self.handle_scrolling();
            self.add_notification("Jumped to last non-empty cell (down)".to_string());
        }
    }

    pub fn start_search_forward(&mut self) {
        self.input_mode = InputMode::SearchForward;
        self.input_buffer = String::new();
        self.text_area = TextArea::default();
        self.add_notification("Search forward mode".to_string());
        self.highlight_enabled = true;
    }

    pub fn add_char_to_input(&mut self, c: char) {
        self.input_buffer.push(c);
    }

    pub fn delete_char_from_input(&mut self) {
        self.input_buffer.pop();
    }

    pub fn start_search_backward(&mut self) {
        self.input_mode = InputMode::SearchBackward;
        self.input_buffer = String::new();
        self.text_area = TextArea::default();
        self.add_notification("Search backward mode".to_string());
        self.highlight_enabled = true;
    }

    pub fn execute_search(&mut self) {
        let query = self.text_area.lines().join("\n");
        self.input_buffer = query.clone();

        if query.is_empty() {
            self.input_mode = InputMode::Normal;
            return;
        }

        // Save the query for n/N commands
        self.search_query = query.clone();

        // Set search direction based on mode
        match self.input_mode {
            InputMode::SearchForward => self.search_direction = true,
            InputMode::SearchBackward => self.search_direction = false,
            _ => {}
        }

        // Perform the search
        self.search_results = self.find_all_matches(&query);

        if self.search_results.is_empty() {
            self.add_notification(format!("Pattern not found: {}", query));
            self.current_search_idx = None;
        } else {
            // Find the appropriate result to jump to based on search direction and current position
            self.jump_to_next_search_result();
            self.add_notification(format!(
                "{} matches found for: {}",
                self.search_results.len(),
                query
            ));
        }

        self.input_mode = InputMode::Normal;
        self.input_buffer = String::new();
        self.text_area = TextArea::default();
    }

    // Find all cells that match the search query
    pub fn find_all_matches(&self, query: &str) -> Vec<(usize, usize)> {
        let sheet = self.workbook.get_current_sheet();
        let query_lower = query.to_lowercase();
        let mut results = Vec::new();

        // Search through all cells in the sheet using row-first, column-second order
        for row in 1..=sheet.max_rows {
            for col in 1..=sheet.max_cols {
                if row < sheet.data.len() && col < sheet.data[row].len() {
                    let cell_content = &sheet.data[row][col].value;
                    if cell_content.to_lowercase().contains(&query_lower) {
                        results.push((row, col));
                    }
                }
            }
        }

        // The search already uses row-first, column-second order, so no need to sort
        results
    }

    // Jump to the next search result based on current position and search direction
    // Uses row-first, column-second order for determining "next" and "previous"
    pub fn jump_to_next_search_result(&mut self) {
        if self.search_results.is_empty() {
            return;
        }

        // Enable highlighting when jumping between search results
        self.highlight_enabled = true;

        let current_pos = self.selected_cell;

        if self.search_direction {
            // Forward search
            // Find the next result after current position using row-first, column-second order
            let next_idx = self.search_results.iter().position(|&pos| {
                // First compare rows, then columns
                pos.0 > current_pos.0 || (pos.0 == current_pos.0 && pos.1 > current_pos.1)
            });

            match next_idx {
                Some(idx) => {
                    self.current_search_idx = Some(idx);
                    self.selected_cell = self.search_results[idx];
                }
                None => {
                    // Wrap around to the first result
                    self.current_search_idx = Some(0);
                    self.selected_cell = self.search_results[0];
                    self.add_notification("Search wrapped to top".to_string());
                }
            }
        } else {
            // Backward search
            // Find the previous result before current position using row-first, column-second order
            let prev_idx = self.search_results.iter().rposition(|&pos| {
                // First compare rows, then columns
                pos.0 < current_pos.0 || (pos.0 == current_pos.0 && pos.1 < current_pos.1)
            });

            match prev_idx {
                Some(idx) => {
                    self.current_search_idx = Some(idx);
                    self.selected_cell = self.search_results[idx];
                }
                None => {
                    // Wrap around to the last result
                    let last_idx = self.search_results.len() - 1;
                    self.current_search_idx = Some(last_idx);
                    self.selected_cell = self.search_results[last_idx];
                    self.add_notification("Search wrapped to bottom".to_string());
                }
            }
        }

        self.handle_scrolling();
    }

    // Jump to the previous search result (opposite of current search direction)
    pub fn jump_to_prev_search_result(&mut self) {
        if self.search_results.is_empty() {
            return;
        }

        // Temporarily flip the search direction
        self.search_direction = !self.search_direction;
        self.jump_to_next_search_result();
        // Restore original search direction
        self.search_direction = !self.search_direction;
    }

    // Disable search highlighting (nohlsearch in Vim)
    pub fn disable_search_highlight(&mut self) {
        self.highlight_enabled = false;
        self.add_notification("Search highlighting disabled".to_string());
    }

    pub fn save_and_exit(&mut self) {
        if !self.workbook.is_modified() {
            self.add_notification("No changes to save".to_string());
            self.should_quit = true;
            return;
        }

        // Try to save the file
        match self.workbook.save() {
            Ok(_) => {
                self.add_notification("File saved".to_string());
                self.should_quit = true;
            }
            Err(e) => {
                self.add_notification(format!("Save failed: {}", e));
                self.input_mode = InputMode::Normal;
            }
        }
    }

    pub fn save(&mut self) -> Result<()> {
        if !self.workbook.is_modified() {
            self.add_notification("No changes to save".to_string());
            return Ok(());
        }

        match self.workbook.save() {
            Ok(_) => {
                self.add_notification("File saved".to_string());
            }
            Err(e) => {
                self.add_notification(format!("Save failed: {}", e));
            }
        }
        Ok(())
    }

    pub fn exit_without_saving(&mut self) {
        self.should_quit = true;
    }

    pub fn next_sheet(&mut self) -> Result<()> {
        let sheet_count = self.workbook.get_sheet_names().len();
        let current_index = self.workbook.get_current_sheet_index();

        if current_index >= sheet_count - 1 {
            self.add_notification("Already at the last sheet".to_string());
            return Ok(());
        }

        self.switch_sheet_by_index(current_index + 1)
    }

    pub fn prev_sheet(&mut self) -> Result<()> {
        let current_index = self.workbook.get_current_sheet_index();

        if current_index == 0 {
            self.add_notification("Already at the first sheet".to_string());
            return Ok(());
        }

        self.switch_sheet_by_index(current_index - 1)
    }

    pub fn switch_sheet_by_index(&mut self, index: usize) -> Result<()> {
        let current_sheet_name = self.workbook.get_current_sheet_name();
        self.sheet_column_widths
            .insert(current_sheet_name, self.column_widths.clone());

        // Reset cell selection and view position when switching sheets
        self.selected_cell = (1, 1);
        self.start_row = 1;
        self.start_col = 1;

        self.workbook.switch_sheet(index)?;

        let new_sheet_name = self.workbook.get_current_sheet_name();

        // Restore column widths for the new sheet
        if let Some(saved_widths) = self.sheet_column_widths.get(&new_sheet_name) {
            self.column_widths = saved_widths.clone();
        } else {
            let max_cols = self.workbook.get_current_sheet().max_cols;
            let default_width = 15;
            self.column_widths = vec![default_width; max_cols + 1];

            self.sheet_column_widths
                .insert(new_sheet_name.clone(), self.column_widths.clone());
        }

        // Clear search results as they're specific to the previous sheet
        self.search_results.clear();
        self.current_search_idx = None;

        self.add_notification(format!("Switched to sheet: {}", new_sheet_name));
        Ok(())
    }

    pub fn switch_to_sheet(&mut self, name_or_index: &str) {
        // Get all sheet names
        let sheet_names = self.workbook.get_sheet_names();

        // Try to parse as index first
        if let Ok(index) = name_or_index.parse::<usize>() {
            // Convert to 0-based index
            let zero_based_index = index.saturating_sub(1);

            if zero_based_index < sheet_names.len() {
                match self.switch_sheet_by_index(zero_based_index) {
                    Ok(_) => return,
                    Err(e) => {
                        self.add_notification(format!(
                            "Failed to switch to sheet {}: {}",
                            index, e
                        ));
                        return;
                    }
                }
            }
        }

        // Try to find by name
        for (i, name) in sheet_names.iter().enumerate() {
            if name.eq_ignore_ascii_case(name_or_index) {
                match self.switch_sheet_by_index(i) {
                    Ok(_) => return,
                    Err(e) => {
                        self.add_notification(format!(
                            "Failed to switch to sheet '{}': {}",
                            name_or_index, e
                        ));
                        return;
                    }
                }
            }
        }

        // If we get here, no matching sheet was found
        self.add_notification(format!("Sheet '{}' not found", name_or_index));
    }

    pub fn delete_current_sheet(&mut self) {
        let current_sheet_name = self.workbook.get_current_sheet_name();

        match self.workbook.delete_current_sheet() {
            Ok(_) => {
                self.sheet_column_widths.remove(&current_sheet_name);

                // Reset cell selection and view position for the current sheet
                self.selected_cell = (1, 1);
                self.start_row = 1;
                self.start_col = 1;

                let new_sheet_name = self.workbook.get_current_sheet_name();

                if let Some(saved_widths) = self.sheet_column_widths.get(&new_sheet_name) {
                    self.column_widths = saved_widths.clone();
                } else {
                    let max_cols = self.workbook.get_current_sheet().max_cols;
                    let default_width = 15;
                    self.column_widths = vec![default_width; max_cols + 1];

                    self.sheet_column_widths
                        .insert(new_sheet_name.clone(), self.column_widths.clone());
                }

                // Clear search results as they're specific to the previous sheet
                self.search_results.clear();
                self.current_search_idx = None;

                self.add_notification(format!("Deleted sheet: {}", current_sheet_name));
            }
            Err(e) => {
                self.add_notification(format!("Failed to delete sheet: {}", e));
            }
        }
    }

    pub fn delete_current_row(&mut self) -> Result<()> {
        let row = self.selected_cell.0;

        self.workbook.delete_row(row)?;

        // Adjust selected cell if needed
        let sheet = self.workbook.get_current_sheet();
        if row > sheet.max_rows {
            self.selected_cell.0 = sheet.max_rows.max(1);
        }

        self.handle_scrolling();

        self.search_results.clear();
        self.current_search_idx = None;

        self.add_notification(format!("Deleted row {}", row));
        Ok(())
    }

    pub fn delete_row(&mut self, row: usize) -> Result<()> {
        self.workbook.delete_row(row)?;

        // Adjust selected cell if needed
        let sheet = self.workbook.get_current_sheet();
        if self.selected_cell.0 > sheet.max_rows {
            self.selected_cell.0 = sheet.max_rows.max(1);
        }

        self.handle_scrolling();

        self.search_results.clear();
        self.current_search_idx = None;

        self.add_notification(format!("Deleted row {}", row));
        Ok(())
    }

    pub fn delete_rows(&mut self, start_row: usize, end_row: usize) -> Result<()> {
        self.workbook.delete_rows(start_row, end_row)?;

        // Adjust selected cell if needed
        let sheet = self.workbook.get_current_sheet();
        if self.selected_cell.0 > sheet.max_rows {
            self.selected_cell.0 = sheet.max_rows.max(1);
        }

        self.handle_scrolling();

        self.search_results.clear();
        self.current_search_idx = None;

        self.add_notification(format!("Deleted rows {} to {}", start_row, end_row));
        Ok(())
    }

    pub fn delete_current_column(&mut self) -> Result<()> {
        let col = self.selected_cell.1;

        self.workbook.delete_column(col)?;

        // Adjust selected cell if needed
        let sheet = self.workbook.get_current_sheet();
        if col > sheet.max_cols {
            self.selected_cell.1 = sheet.max_cols.max(1);
        }

        if self.column_widths.len() > col {
            self.column_widths.remove(col);
            // Add a default width for the last column to maintain the vector size
            self.column_widths.push(15);
        }

        self.handle_scrolling();

        self.search_results.clear();
        self.current_search_idx = None;

        self.add_notification(format!("Deleted column {}", index_to_col_name(col)));
        Ok(())
    }

    pub fn delete_column(&mut self, col: usize) -> Result<()> {
        self.workbook.delete_column(col)?;

        // Adjust selected cell if needed
        let sheet = self.workbook.get_current_sheet();
        if self.selected_cell.1 > sheet.max_cols {
            self.selected_cell.1 = sheet.max_cols.max(1);
        }

        if self.column_widths.len() > col {
            self.column_widths.remove(col);
            // Add a default width for the last column to maintain the vector size
            self.column_widths.push(15);
        }

        self.handle_scrolling();

        self.search_results.clear();
        self.current_search_idx = None;

        self.add_notification(format!("Deleted column {}", index_to_col_name(col)));
        Ok(())
    }

    pub fn delete_columns(&mut self, start_col: usize, end_col: usize) -> Result<()> {
        self.workbook.delete_columns(start_col, end_col)?;

        // Adjust selected cell if needed
        let sheet = self.workbook.get_current_sheet();
        if self.selected_cell.1 > sheet.max_cols {
            self.selected_cell.1 = sheet.max_cols.max(1);
        }

        let cols_to_remove = end_col - start_col + 1;
        for col in (start_col..=end_col).rev() {
            if self.column_widths.len() > col {
                self.column_widths.remove(col);
            }
        }
        // Add default widths for the removed columns to maintain the vector size
        for _ in 0..cols_to_remove {
            self.column_widths.push(15);
        }

        self.handle_scrolling();

        self.search_results.clear();
        self.current_search_idx = None;

        self.add_notification(format!(
            "Deleted columns {} to {}",
            index_to_col_name(start_col),
            index_to_col_name(end_col)
        ));
        Ok(())
    }

    pub fn copy_cell(&mut self) {
        let content = self.get_cell_content(self.selected_cell.0, self.selected_cell.1);
        self.clipboard = Some(content);
        self.add_notification("Cell content copied".to_string());
    }

    pub fn cut_cell(&mut self) -> Result<()> {
        let content = self.get_cell_content(self.selected_cell.0, self.selected_cell.1);
        self.clipboard = Some(content);

        // Clear the cell
        self.workbook
            .set_cell_value(self.selected_cell.0, self.selected_cell.1, String::new())?;

        self.add_notification("Cell content cut".to_string());
        Ok(())
    }

    pub fn paste_cell(&mut self) -> Result<()> {
        if let Some(content) = &self.clipboard {
            self.workbook.set_cell_value(
                self.selected_cell.0,
                self.selected_cell.1,
                content.clone(),
            )?;
            self.add_notification("Content pasted".to_string());
        } else {
            self.add_notification("Clipboard is empty".to_string());
        }
        Ok(())
    }

    pub fn auto_adjust_column_width(&mut self, col: Option<usize>) {
        let sheet = self.workbook.get_current_sheet();
        let default_min_width = 5;

        match col {
            // Adjust specific column
            Some(column) => {
                if column < self.column_widths.len() {
                    // Record original column width for debugging
                    let old_width = self.column_widths[column];

                    // Calculate and set new column width
                    let width = self.calculate_column_width(column);
                    self.column_widths[column] = width.max(default_min_width);

                    self.ensure_column_visible(column);

                    self.add_notification(format!(
                        "Column {} width adjusted from {} to {}",
                        index_to_col_name(column),
                        old_width,
                        self.column_widths[column]
                    ));
                }
            }
            // Adjust all columns
            None => {
                for col_idx in 1..=sheet.max_cols {
                    let width = self.calculate_column_width(col_idx);
                    self.column_widths[col_idx] = width.max(default_min_width);
                }

                let column = self.selected_cell.1;
                self.ensure_column_visible(column);

                self.add_notification("All column widths adjusted".to_string());
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

    // Adjust column width to a reasonable minimum (max 15 or actual content width)
    pub fn shrink_column_width(&mut self, col: Option<usize>) {
        let max_width = 15; // Maximum column width limit
        let min_width = 5; // Minimum column width

        match col {
            // Minimize specific column
            Some(column) => {
                if column < self.column_widths.len() {
                    // Record current width
                    let current_width = self.column_widths[column];

                    // Calculate actual content width
                    let content_width = self.calculate_column_width(column);

                    // Set width to the minimum of content width and max width, but not less than min width
                    let new_width = content_width.min(max_width).max(min_width);
                    self.column_widths[column] = new_width;

                    self.ensure_column_visible(column);

                    self.add_notification(format!(
                        "Column {} width minimized from {} to {}",
                        index_to_col_name(column),
                        current_width,
                        new_width
                    ));
                }
            }
            // Minimize all columns
            None => {
                let sheet = self.workbook.get_current_sheet();

                // Track how many columns changed for status message
                let mut _changed_columns = 0;

                for col_idx in 1..=sheet.max_cols {
                    // Calculate actual content width
                    let content_width = self.calculate_column_width(col_idx);

                    // Set width to the minimum of content width and max width, but not less than min width
                    let new_width = content_width.min(max_width).max(min_width);

                    // Track if width changed
                    if self.column_widths[col_idx] != new_width {
                        self.column_widths[col_idx] = new_width;
                        _changed_columns += 1;
                    }
                }

                let column = self.selected_cell.1;
                self.ensure_column_visible(column);

                self.add_notification("Minimized the width of all columns".to_string());
            }
        }
    }

    fn handle_column_scrolling(&mut self) {
        self.ensure_column_visible(self.selected_cell.1);
    }

    pub fn ensure_column_visible(&mut self, column: usize) {
        let desired_right_margin_chars = 15;

        if column < self.start_col {
            // If column is to the left of visible area, adjust start_col
            self.start_col = column;
            return;
        }

        let last_visible_col = self.start_col + self.visible_cols - 1;

        if column > last_visible_col {
            // If column is to the right of visible area, adjust start_col to make it visible
            self.start_col = column - self.visible_cols + 1;
            self.start_col = self.start_col.max(1);
            return;
        }

        // If we're here, the column is already visible
        // add a right margin if possible

        let sheet = self.workbook.get_current_sheet();
        let max_col = sheet.max_cols;

        if column < max_col {
            let cols_to_right = last_visible_col - column;

            if cols_to_right > 0 {
                return;
            }

            let next_col = column + 1;
            if next_col <= max_col {
                let next_col_width = self.get_column_width(next_col);

                if next_col_width <= desired_right_margin_chars {
                    if self.visible_cols > 1 {
                        self.start_col = column - (self.visible_cols - 2);
                        self.start_col = self.start_col.max(1);
                    }
                } else {
                    if self.visible_cols > 1 {
                        self.start_col = column - (self.visible_cols - 2);
                        self.start_col = self.start_col.max(1);
                    }
                }
            }
        }
    }

    pub fn start_command_mode(&mut self) {
        self.input_mode = InputMode::Command;
        self.input_buffer = String::new();
    }
}

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

pub use crate::utils::col_name_to_index;

pub use crate::utils::index_to_col_name;
