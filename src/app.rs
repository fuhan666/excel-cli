use anyhow::Result;
use ratatui_textarea::TextArea;
use std::collections::HashMap;
use std::path::PathBuf;

use crate::excel::{Cell, Workbook};
use crate::undo::{
    ActionData, ActionType, CellAction, ColumnAction, MultiColumnAction, MultiRowAction, RowAction,
    SheetAction, UndoAction, UndoHistory,
};
use crate::utils::{Direction, find_non_empty_cell};

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
    pub undo_history: UndoHistory,
}

impl AppState<'_> {
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
            info_panel_height: 10,
            notification_messages: Vec::new(),
            max_notifications: 5,
            help_text: String::new(),
            help_scroll: 0,
            help_visible_lines: 20,
            undo_history: UndoHistory::new(),
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
                let mut result = String::with_capacity(9 + cell.value.len());
                result.push_str("Formula: ");
                result.push_str(&cell.value);
                result
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
        let new_height = (self.info_panel_height as isize + delta).clamp(6, 16) as usize;
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

        self.help_text = "FILE OPERATIONS:\n\
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
             :put, :pu   - Paste to current cell\n\
             u           - Undo last operation\n\
             Ctrl+r      - Redo last undone operation\n\n\
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
             -           - Decrease info panel height"
            .to_string();

        self.input_mode = InputMode::Help;
    }

    pub fn confirm_edit(&mut self) -> Result<()> {
        if let InputMode::Editing = self.input_mode {
            // Get content from TextArea
            let content = self.text_area.lines().join("\n");
            let (row, col) = self.selected_cell;

            let sheet_index = self.workbook.get_current_sheet_index();
            let sheet_name = self.workbook.get_current_sheet_name();

            let old_cell = self.workbook.get_current_sheet().data[row][col].clone();

            let mut new_cell = old_cell.clone();
            new_cell.value = content.clone();

            let cell_action = CellAction {
                sheet_index,
                sheet_name,
                row,
                col,
                old_value: old_cell,
                new_value: new_cell,
            };

            let undo_action = UndoAction {
                action_type: ActionType::Edit,
                action_data: ActionData::Cell(cell_action),
            };

            self.undo_history.push(undo_action);

            self.workbook.set_cell_value(row, col, content)?;
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

    /// Generic method for navigating to non-empty cells
    fn jump_to_non_empty_cell(&mut self, direction: Direction) {
        let sheet = self.workbook.get_current_sheet();
        let max_bounds = (sheet.max_rows, sheet.max_cols);
        let current_pos = self.selected_cell;

        if let Some(new_pos) = find_non_empty_cell(sheet, current_pos, direction, max_bounds) {
            self.selected_cell = new_pos;
            self.handle_scrolling();

            let dir_name = match direction {
                Direction::Left => "left",
                Direction::Right => "right",
                Direction::Up => "up",
                Direction::Down => "down",
            };

            // Re-fetch sheet to avoid borrow conflict
            let sheet = self.workbook.get_current_sheet();

            let (row, col) = self.selected_cell;
            let is_cell_empty = row >= sheet.data.len()
                || col >= sheet.data[row].len()
                || sheet.data[row][col].value.is_empty();

            let message = if is_cell_empty {
                format!("Jumped to first non-empty cell ({})", dir_name)
            } else {
                format!("Jumped to last non-empty cell ({})", dir_name)
            };

            self.add_notification(message);
        }
    }

    // Direction-specific convenience methods

    pub fn jump_to_prev_non_empty_cell_left(&mut self) {
        self.jump_to_non_empty_cell(Direction::Left);
    }

    pub fn jump_to_prev_non_empty_cell_right(&mut self) {
        self.jump_to_non_empty_cell(Direction::Right);
    }

    pub fn jump_to_prev_non_empty_cell_up(&mut self) {
        self.jump_to_non_empty_cell(Direction::Up);
    }

    pub fn jump_to_prev_non_empty_cell_down(&mut self) {
        self.jump_to_non_empty_cell(Direction::Down);
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

        let mut results = Vec::with_capacity(32);

        // Search through all cells in the sheet using row-first, column-second order
        for row in 1..=sheet.max_rows {
            for col in 1..=sheet.max_cols {
                if row < sheet.data.len() && col < sheet.data[row].len() {
                    let cell_content = &sheet.data[row][col].value;

                    if self.case_insensitive_contains(cell_content, &query_lower) {
                        results.push((row, col));
                    }
                }
            }
        }

        results
    }

    fn case_insensitive_contains(&self, haystack: &str, needle: &str) -> bool {
        if needle.is_empty() {
            return true;
        }
        if haystack.is_empty() {
            return false;
        }

        let haystack_lower = haystack.to_lowercase();
        haystack_lower.contains(needle)
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

        match self.workbook.save() {
            Ok(_) => {
                self.undo_history.clear();
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
                self.undo_history.clear();
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

        if !self.sheet_column_widths.contains_key(&current_sheet_name)
            || self.sheet_column_widths[&current_sheet_name] != self.column_widths
        {
            self.sheet_column_widths
                .insert(current_sheet_name, self.column_widths.clone());
        }

        // Reset cell selection and view position when switching sheets
        self.selected_cell = (1, 1);
        self.start_row = 1;
        self.start_col = 1;

        self.workbook.switch_sheet(index)?;

        let new_sheet_name = self.workbook.get_current_sheet_name();

        // Restore column widths for the new sheet
        if let Some(saved_widths) = self.sheet_column_widths.get(&new_sheet_name) {
            if &self.column_widths != saved_widths {
                self.column_widths = saved_widths.clone();
            }
        } else {
            let max_cols = self.workbook.get_current_sheet().max_cols;
            let default_width = 15;
            self.column_widths = vec![default_width; max_cols + 1];

            self.sheet_column_widths
                .insert(new_sheet_name.clone(), self.column_widths.clone());
        }

        // Clear search results as they're specific to the previous sheet
        if !self.search_results.is_empty() {
            self.search_results.clear();
            self.current_search_idx = None;
        }

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
        let sheet_index = self.workbook.get_current_sheet_index();

        // Save the sheet data for undo
        let sheet_data = self.workbook.get_current_sheet().clone();
        let column_widths = self.column_widths.clone();

        match self.workbook.delete_current_sheet() {
            Ok(_) => {
                // Create the undo action
                let sheet_action = SheetAction {
                    sheet_index,
                    sheet_name: current_sheet_name.clone(),
                    sheet_data,
                    column_widths,
                };

                let undo_action = UndoAction {
                    action_type: ActionType::DeleteSheet,
                    action_data: ActionData::Sheet(sheet_action),
                };

                self.undo_history.push(undo_action);

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

        // Save row data for undo
        let sheet_index = self.workbook.get_current_sheet_index();
        let sheet_name = self.workbook.get_current_sheet_name();
        let sheet = self.workbook.get_current_sheet();

        // Create a copy of the row data before deletion
        let row_data = if row < sheet.data.len() {
            sheet.data[row].clone()
        } else {
            Vec::new()
        };

        // Create and add undo action
        let row_action = RowAction {
            sheet_index,
            sheet_name,
            row,
            row_data,
        };

        let undo_action = UndoAction {
            action_type: ActionType::DeleteRow,
            action_data: ActionData::Row(row_action),
        };

        self.undo_history.push(undo_action);

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
        // Save row data for undo
        let sheet_index = self.workbook.get_current_sheet_index();
        let sheet_name = self.workbook.get_current_sheet_name();
        let sheet = self.workbook.get_current_sheet();

        // Create a copy of the row data before deletion
        let row_data = if row < sheet.data.len() {
            sheet.data[row].clone()
        } else {
            Vec::new()
        };

        // Create and add undo action
        let row_action = RowAction {
            sheet_index,
            sheet_name,
            row,
            row_data,
        };

        let undo_action = UndoAction {
            action_type: ActionType::DeleteRow,
            action_data: ActionData::Row(row_action),
        };

        self.undo_history.push(undo_action);

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
        if start_row == end_row {
            return self.delete_row(start_row);
        }

        // Save all row data for batch undo
        let sheet_index = self.workbook.get_current_sheet_index();
        let sheet_name = self.workbook.get_current_sheet_name();
        let sheet = self.workbook.get_current_sheet();

        // Save row data in the original order from top to bottom
        let mut rows_data = Vec::with_capacity(end_row - start_row + 1);

        for row in start_row..=end_row {
            if row < sheet.data.len() {
                rows_data.push(sheet.data[row].clone());
            } else {
                rows_data.push(Vec::new());
            }
        }

        // Create and add batch undo action
        let multi_row_action = MultiRowAction {
            sheet_index,
            sheet_name,
            start_row,
            end_row,
            rows_data,
        };

        let undo_action = UndoAction {
            action_type: ActionType::DeleteMultiRows,
            action_data: ActionData::MultiRow(multi_row_action),
        };

        self.undo_history.push(undo_action);

        self.workbook.delete_rows(start_row, end_row)?;

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

        // Save column data for undo
        let sheet_index = self.workbook.get_current_sheet_index();
        let sheet_name = self.workbook.get_current_sheet_name();
        let sheet = self.workbook.get_current_sheet();

        // Extract the column data from each row
        let mut column_data = Vec::new();
        for row in &sheet.data {
            if col < row.len() {
                column_data.push(row[col].clone());
            } else {
                column_data.push(Cell::empty());
            }
        }

        // Save the column width
        let column_width = if col < self.column_widths.len() {
            self.column_widths[col]
        } else {
            15 // Default width
        };

        let column_action = ColumnAction {
            sheet_index,
            sheet_name,
            col,
            column_data,
            column_width,
        };

        let undo_action = UndoAction {
            action_type: ActionType::DeleteColumn,
            action_data: ActionData::Column(column_action),
        };

        self.undo_history.push(undo_action);

        self.workbook.delete_column(col)?;

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
        // Save column data for undo
        let sheet_index = self.workbook.get_current_sheet_index();
        let sheet_name = self.workbook.get_current_sheet_name();
        let sheet = self.workbook.get_current_sheet();

        // Extract the column data from each row
        let mut column_data = Vec::new();
        for row in &sheet.data {
            if col < row.len() {
                column_data.push(row[col].clone());
            } else {
                column_data.push(Cell::empty());
            }
        }

        // Save the column width
        let column_width = if col < self.column_widths.len() {
            self.column_widths[col]
        } else {
            15 // Default width
        };

        let column_action = ColumnAction {
            sheet_index,
            sheet_name,
            col,
            column_data,
            column_width,
        };

        let undo_action = UndoAction {
            action_type: ActionType::DeleteColumn,
            action_data: ActionData::Column(column_action),
        };

        self.undo_history.push(undo_action);

        self.workbook.delete_column(col)?;

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
        if start_col == end_col {
            return self.delete_column(start_col);
        }

        // For multiple columns, save all column data for batch undo
        let sheet_index = self.workbook.get_current_sheet_index();
        let sheet_name = self.workbook.get_current_sheet_name();
        let sheet = self.workbook.get_current_sheet();

        // Save column data and widths for batch undo
        let mut columns_data = Vec::with_capacity(end_col - start_col + 1);
        let mut column_widths = Vec::with_capacity(end_col - start_col + 1);

        for col in start_col..=end_col {
            // Extract the column data from each row
            let mut column_data = Vec::new();
            for row in &sheet.data {
                if col < row.len() {
                    column_data.push(row[col].clone());
                } else {
                    column_data.push(Cell::empty());
                }
            }
            columns_data.push(column_data);

            // Save the column width
            let column_width = if col < self.column_widths.len() {
                self.column_widths[col]
            } else {
                15 // Default width
            };
            column_widths.push(column_width);
        }

        // Create and add batch undo action
        let multi_column_action = MultiColumnAction {
            sheet_index,
            sheet_name,
            start_col,
            end_col,
            columns_data,
            column_widths,
        };

        let undo_action = UndoAction {
            action_type: ActionType::DeleteMultiColumns,
            action_data: ActionData::MultiColumn(multi_column_action),
        };

        self.undo_history.push(undo_action);

        self.workbook.delete_columns(start_col, end_col)?;

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
        let (row, col) = self.selected_cell;
        let content = self.get_cell_content(row, col);
        self.clipboard = Some(content);

        let sheet_index = self.workbook.get_current_sheet_index();
        let sheet_name = self.workbook.get_current_sheet_name();

        let old_cell = self.workbook.get_current_sheet().data[row][col].clone();

        let mut new_cell = old_cell.clone();
        new_cell.value = String::new();

        let cell_action = CellAction {
            sheet_index,
            sheet_name,
            row,
            col,
            old_value: old_cell,
            new_value: new_cell,
        };

        let undo_action = UndoAction {
            action_type: ActionType::Cut,
            action_data: ActionData::Cell(cell_action),
        };

        self.undo_history.push(undo_action);

        self.workbook.set_cell_value(row, col, String::new())?;

        self.add_notification("Cell content cut".to_string());
        Ok(())
    }

    pub fn paste_cell(&mut self) -> Result<()> {
        if let Some(content) = &self.clipboard {
            let (row, col) = self.selected_cell;

            let sheet_index = self.workbook.get_current_sheet_index();
            let sheet_name = self.workbook.get_current_sheet_name();

            let old_cell = self.workbook.get_current_sheet().data[row][col].clone();

            let mut new_cell = old_cell.clone();
            new_cell.value = content.clone();

            let cell_action = CellAction {
                sheet_index,
                sheet_name,
                row,
                col,
                old_value: old_cell,
                new_value: new_cell,
            };

            let undo_action = UndoAction {
                action_type: ActionType::Paste,
                action_data: ActionData::Cell(cell_action),
            };

            self.undo_history.push(undo_action);

            self.workbook.set_cell_value(row, col, content.clone())?;
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
                    // Calculate and set new column width
                    let width = self.calculate_column_width(column);
                    self.column_widths[column] = width.max(default_min_width);

                    self.ensure_column_visible(column);

                    self.add_notification(format!(
                        "Column {} width adjusted",
                        index_to_col_name(column)
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

        // Start with minimum width and header width
        let col_name = index_to_col_name(col);
        let mut max_width = 3.max(col_name.len());

        // Calculate max width from all cells in the column
        for row in 1..=sheet.max_rows {
            if row >= sheet.data.len() || col >= sheet.data[row].len() {
                continue;
            }

            let content = &sheet.data[row][col].value;
            if content.is_empty() {
                continue;
            }

            let mut is_ascii_only = true;
            let mut display_width = 0;

            for c in content.chars() {
                if c.is_ascii() {
                    display_width += 1;
                } else {
                    is_ascii_only = false;
                    display_width += 2;
                }
            }

            let cell_width = if is_ascii_only {
                display_width + (display_width / 3)
            } else {
                display_width
            };

            max_width = max_width.max(cell_width);
        }

        // Add appropriate padding based on content width
        if max_width > 20 {
            max_width + 3 // Less padding for wide content
        } else {
            max_width + 4 // More padding for narrow content
        }
    }

    pub fn get_column_width(&self, col: usize) -> usize {
        if col < self.column_widths.len() {
            self.column_widths[col]
        } else {
            15 // Default width
        }
    }

    fn handle_column_scrolling(&mut self) {
        self.ensure_column_visible(self.selected_cell.1);
    }

    pub fn ensure_column_visible(&mut self, column: usize) {
        // If column is to the left of visible area, adjust start_col
        if column < self.start_col {
            self.start_col = column;
            return;
        }

        let last_visible_col = self.start_col + self.visible_cols - 1;

        // If column is to the right of visible area, adjust start_col to make it visible
        if column > last_visible_col {
            self.start_col = (column - self.visible_cols + 1).max(1);
            return;
        }

        // If the column is already visible but at the right edge, try to add a margin
        let sheet = self.workbook.get_current_sheet();
        let max_col = sheet.max_cols;

        // Only apply margin logic if not at the max column
        if column < max_col && column == last_visible_col && self.visible_cols > 1 {
            // Adjust start column to show more columns to the left
            // This creates a margin on the right
            self.start_col = (column - (self.visible_cols - 2)).max(1);
        }
    }

    pub fn start_command_mode(&mut self) {
        self.input_mode = InputMode::Command;
        self.input_buffer = String::new();
    }

    fn apply_action(&mut self, action: &UndoAction, is_undo: bool) -> Result<()> {
        match &action.action_data {
            ActionData::Cell(cell_action) => {
                let value = if is_undo {
                    &cell_action.old_value
                } else {
                    &cell_action.new_value
                };
                self.apply_cell_action(cell_action, value, is_undo, &action.action_type)?;
            }
            ActionData::Row(row_action) => {
                self.apply_row_action(row_action, is_undo)?;
            }
            ActionData::Column(column_action) => {
                self.apply_column_action(column_action, is_undo)?;
            }
            ActionData::Sheet(sheet_action) => {
                self.apply_sheet_action(sheet_action, is_undo)?;
            }
            ActionData::MultiRow(multi_row_action) => {
                self.apply_multi_row_action(multi_row_action, is_undo)?;
            }
            ActionData::MultiColumn(multi_column_action) => {
                self.apply_multi_column_action(multi_column_action, is_undo)?;
            }
        }
        Ok(())
    }

    fn apply_cell_action(
        &mut self,
        cell_action: &CellAction,
        value: &Cell,
        is_undo: bool,
        action_type: &ActionType,
    ) -> Result<()> {
        let current_sheet_index = self.workbook.get_current_sheet_index();

        if current_sheet_index != cell_action.sheet_index {
            if let Err(e) = self.switch_sheet_by_index(cell_action.sheet_index) {
                self.add_notification(format!(
                    "Cannot switch to sheet {}: {}",
                    cell_action.sheet_name, e
                ));
                return Ok(());
            }
        }

        self.workbook.get_current_sheet_mut().data[cell_action.row][cell_action.col] =
            value.clone();

        self.selected_cell = (cell_action.row, cell_action.col);
        self.handle_scrolling();

        let cell_ref = format!(
            "{}{}",
            crate::utils::index_to_col_name(cell_action.col),
            cell_action.row
        );

        let operation_text = match action_type {
            ActionType::Edit => "edit",
            ActionType::Cut => "cut",
            ActionType::Paste => "paste",
            _ => "cell operation",
        };

        if current_sheet_index != cell_action.sheet_index {
            let action_word = if is_undo { "Undid" } else { "Redid" };
            self.add_notification(format!(
                "{} {} operation on cell {} in sheet {}",
                action_word, operation_text, cell_ref, cell_action.sheet_name
            ));
        } else {
            let action_word = if is_undo { "Undid" } else { "Redid" };
            self.add_notification(format!(
                "{} {} operation on cell {}",
                action_word, operation_text, cell_ref
            ));
        }

        Ok(())
    }

    fn apply_row_action(&mut self, row_action: &RowAction, is_undo: bool) -> Result<()> {
        let current_sheet_index = self.workbook.get_current_sheet_index();

        if current_sheet_index != row_action.sheet_index {
            if let Err(e) = self.switch_sheet_by_index(row_action.sheet_index) {
                self.add_notification(format!(
                    "Cannot switch to sheet {}: {}",
                    row_action.sheet_name, e
                ));
                return Ok(());
            }
        }

        let sheet = self.workbook.get_current_sheet_mut();

        if is_undo {
            sheet
                .data
                .insert(row_action.row, row_action.row_data.clone());

            sheet.max_rows = sheet.max_rows.saturating_add(1);

            self.add_notification(format!("Undid row {} deletion", row_action.row));
        } else if row_action.row < sheet.data.len() {
            sheet.data.remove(row_action.row);
            sheet.max_rows = sheet.max_rows.saturating_sub(1);

            if self.selected_cell.0 > sheet.max_rows {
                self.selected_cell.0 = sheet.max_rows.max(1);
            }

            self.add_notification(format!("Redid row {} deletion", row_action.row));
        }

        self.handle_scrolling();

        self.search_results.clear();
        self.current_search_idx = None;

        Ok(())
    }

    fn apply_column_action(&mut self, column_action: &ColumnAction, is_undo: bool) -> Result<()> {
        let current_sheet_index = self.workbook.get_current_sheet_index();

        if current_sheet_index != column_action.sheet_index {
            if let Err(e) = self.switch_sheet_by_index(column_action.sheet_index) {
                self.add_notification(format!(
                    "Cannot switch to sheet {}: {}",
                    column_action.sheet_name, e
                ));
                return Ok(());
            }
        }

        let sheet = self.workbook.get_current_sheet_mut();
        let col = column_action.col;

        if is_undo {
            let column_data = &column_action.column_data;

            for (i, row) in sheet.data.iter_mut().enumerate() {
                if i < column_data.len() {
                    if col <= row.len() {
                        row.insert(col, column_data[i].clone());
                    } else {
                        while row.len() < col {
                            row.push(Cell::empty());
                        }
                        row.push(column_data[i].clone());
                    }
                }
            }

            sheet.max_cols = sheet.max_cols.saturating_add(1);

            if col < self.column_widths.len() {
                self.column_widths.insert(col, column_action.column_width);
                if !self.column_widths.is_empty() {
                    self.column_widths.pop();
                }
            } else {
                while self.column_widths.len() < col {
                    self.column_widths.push(15); // Default width
                }
                self.column_widths.push(column_action.column_width);
            }

            self.ensure_column_visible(col);

            self.add_notification(format!("Undid column {} deletion", index_to_col_name(col)));
        } else {
            for row in sheet.data.iter_mut() {
                if col < row.len() {
                    row.remove(col);
                }
            }

            sheet.max_cols = sheet.max_cols.saturating_sub(1);

            if self.column_widths.len() > col {
                self.column_widths.remove(col);
                self.column_widths.push(15);
            }

            if self.selected_cell.1 > sheet.max_cols {
                self.selected_cell.1 = sheet.max_cols.max(1);
            }

            self.add_notification(format!("Redid column {} deletion", index_to_col_name(col)));
        }

        self.handle_scrolling();

        self.search_results.clear();
        self.current_search_idx = None;

        Ok(())
    }

    fn apply_sheet_action(&mut self, sheet_action: &SheetAction, is_undo: bool) -> Result<()> {
        if is_undo {
            let sheet_index = sheet_action.sheet_index;

            if let Err(e) = self
                .workbook
                .insert_sheet_at_index(sheet_action.sheet_data.clone(), sheet_index)
            {
                self.add_notification(format!(
                    "Failed to restore sheet {}: {}",
                    sheet_action.sheet_name, e
                ));
                return Ok(());
            }

            self.sheet_column_widths.insert(
                sheet_action.sheet_name.clone(),
                sheet_action.column_widths.clone(),
            );

            if let Err(e) = self.switch_sheet_by_index(sheet_index) {
                self.add_notification(format!(
                    "Restored sheet {} but couldn't switch to it: {}",
                    sheet_action.sheet_name, e
                ));
            } else {
                self.add_notification(format!("Undid sheet {} deletion", sheet_action.sheet_name));
            }
        } else {
            if let Err(e) = self.switch_sheet_by_index(sheet_action.sheet_index) {
                self.add_notification(format!(
                    "Cannot switch to sheet {} to delete it: {}",
                    sheet_action.sheet_name, e
                ));
                return Ok(());
            }

            if let Err(e) = self.workbook.delete_current_sheet() {
                self.add_notification(format!("Failed to delete sheet: {}", e));
                return Ok(());
            }

            self.cleanup_after_sheet_deletion(&sheet_action.sheet_name);

            self.add_notification(format!(
                "Redid deletion of sheet {}",
                sheet_action.sheet_name
            ));
        }

        Ok(())
    }

    fn cleanup_after_sheet_deletion(&mut self, sheet_name: &str) {
        self.sheet_column_widths.remove(sheet_name);

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

        self.search_results.clear();
        self.current_search_idx = None;
    }

    fn apply_multi_row_action(
        &mut self,
        multi_row_action: &MultiRowAction,
        is_undo: bool,
    ) -> Result<()> {
        let current_sheet_index = self.workbook.get_current_sheet_index();

        if current_sheet_index != multi_row_action.sheet_index {
            if let Err(e) = self.switch_sheet_by_index(multi_row_action.sheet_index) {
                self.add_notification(format!(
                    "Cannot switch to sheet {}: {}",
                    multi_row_action.sheet_name, e
                ));
                return Ok(());
            }
        }

        let start_row = multi_row_action.start_row;
        let end_row = multi_row_action.end_row;
        let rows_to_restore = end_row - start_row + 1;

        if is_undo {
            let rows_data = &multi_row_action.rows_data;

            let sheet = self.workbook.get_current_sheet_mut();

            Self::restore_rows(sheet, start_row, rows_data);

            sheet.max_rows = sheet.max_rows.saturating_add(rows_to_restore);

            self.add_notification(format!("Undid rows {} to {} deletion", start_row, end_row));
        } else {
            self.workbook.delete_rows(start_row, end_row)?;

            let sheet = self.workbook.get_current_sheet();

            if self.selected_cell.0 > sheet.max_rows {
                self.selected_cell.0 = sheet.max_rows.max(1);
            }

            self.add_notification(format!("Redid rows {} to {} deletion", start_row, end_row));
        }

        self.handle_scrolling();

        self.search_results.clear();
        self.current_search_idx = None;

        Ok(())
    }

    fn apply_multi_column_action(
        &mut self,
        multi_column_action: &MultiColumnAction,
        is_undo: bool,
    ) -> Result<()> {
        let current_sheet_index = self.workbook.get_current_sheet_index();

        if current_sheet_index != multi_column_action.sheet_index {
            if let Err(e) = self.switch_sheet_by_index(multi_column_action.sheet_index) {
                self.add_notification(format!(
                    "Cannot switch to sheet {}: {}",
                    multi_column_action.sheet_name, e
                ));
                return Ok(());
            }
        }

        let start_col = multi_column_action.start_col;
        let end_col = multi_column_action.end_col;
        let cols_to_restore = end_col - start_col + 1;

        if is_undo {
            let columns_data = &multi_column_action.columns_data;
            let column_widths = &multi_column_action.column_widths;

            let sheet = self.workbook.get_current_sheet_mut();

            for col_idx in (0..cols_to_restore).rev() {
                if col_idx < columns_data.len() {
                    let column_data = &columns_data[col_idx];
                    Self::restore_column_at_position(sheet, start_col, column_data);

                    Self::restore_column_width(
                        &mut self.column_widths,
                        start_col,
                        col_idx,
                        column_widths,
                    );
                }
            }

            sheet.max_cols = sheet.max_cols.saturating_add(cols_to_restore);

            Self::trim_column_widths(&mut self.column_widths, cols_to_restore);

            self.ensure_column_visible(start_col);

            self.add_notification(format!(
                "Undid columns {} to {} deletion",
                index_to_col_name(start_col),
                index_to_col_name(end_col)
            ));
        } else {
            self.workbook.delete_columns(start_col, end_col)?;

            let sheet = self.workbook.get_current_sheet();

            Self::remove_column_widths(&mut self.column_widths, start_col, end_col);

            if self.selected_cell.1 > sheet.max_cols {
                self.selected_cell.1 = sheet.max_cols.max(1);
            }

            self.add_notification(format!(
                "Redid columns {} to {} deletion",
                index_to_col_name(start_col),
                index_to_col_name(end_col)
            ));
        }

        self.handle_scrolling();

        self.search_results.clear();
        self.current_search_idx = None;

        Ok(())
    }

    fn restore_rows(sheet: &mut crate::excel::Sheet, position: usize, rows_data: &[Vec<Cell>]) {
        for row_data in rows_data.iter().rev() {
            sheet.data.insert(position, row_data.clone());
        }
    }

    fn restore_column_at_position(
        sheet: &mut crate::excel::Sheet,
        position: usize,
        column_data: &[Cell],
    ) {
        for (i, row) in sheet.data.iter_mut().enumerate() {
            if i < column_data.len() {
                if position <= row.len() {
                    row.insert(position, column_data[i].clone());
                } else {
                    while row.len() < position {
                        row.push(Cell::empty());
                    }
                    row.push(column_data[i].clone());
                }
            }
        }
    }

    fn restore_column_width(
        column_widths: &mut Vec<usize>,
        position: usize,
        col_idx: usize,
        width_values: &[usize],
    ) {
        if position < column_widths.len() {
            let width = if col_idx < width_values.len() {
                width_values[col_idx]
            } else {
                15 // Default width
            };
            column_widths.insert(position, width);
        }
    }

    // Helper method to trim excess column widths
    fn trim_column_widths(column_widths: &mut Vec<usize>, count: usize) {
        for _ in 0..count {
            if !column_widths.is_empty() {
                column_widths.pop();
            }
        }
    }

    fn remove_column_widths(column_widths: &mut Vec<usize>, start_col: usize, end_col: usize) {
        let cols_to_remove = end_col - start_col + 1;

        for col in (start_col..=end_col).rev() {
            if column_widths.len() > col {
                column_widths.remove(col);
            }
        }

        for _ in 0..cols_to_remove {
            column_widths.push(15);
        }
    }

    pub fn undo(&mut self) -> Result<()> {
        if let Some(action) = self.undo_history.undo() {
            self.apply_action(&action, true)?;

            if self.undo_history.all_undone() {
                self.workbook.set_modified(false);
            } else {
                self.workbook.set_modified(true);
            }
        } else {
            self.add_notification("No operations to undo".to_string());
        }
        Ok(())
    }

    pub fn redo(&mut self) -> Result<()> {
        if let Some(action) = self.undo_history.redo() {
            self.apply_action(&action, false)?;
            self.workbook.set_modified(true);
        } else {
            self.add_notification("No operations to redo".to_string());
        }
        Ok(())
    }
}

pub use crate::utils::index_to_col_name;
