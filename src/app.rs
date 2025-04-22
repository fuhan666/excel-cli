use anyhow::Result;
use std::path::{Path, PathBuf};

use crate::excel::Workbook;
use crate::json_export::{export_json, HeaderDirection};

pub enum InputMode {
    Normal,
    Editing,
    Confirm,
    Command,
    SearchForward,  // For / search
    SearchBackward, // For ? search
}

pub struct AppState {
    pub workbook: Workbook,
    pub file_path: PathBuf,
    pub selected_cell: (usize, usize), // (row, col)
    pub start_row: usize,
    pub start_col: usize,
    pub visible_rows: usize,
    pub visible_cols: usize,
    pub input_mode: InputMode,
    pub input_buffer: String,
    pub status_message: String,
    pub should_quit: bool,
    pub column_widths: Vec<usize>, // Store width for each column
    pub clipboard: Option<String>, // Store copied/cut cell content
    pub g_pressed: bool, // Track if 'g' was pressed for 'gg' command
    pub search_query: String, // Current search query
    pub search_results: Vec<(usize, usize)>, // List of cells matching the search query
    pub current_search_idx: Option<usize>, // Index of current search result
    pub search_direction: bool, // true for forward, false for backward
    pub highlight_enabled: bool, // Control whether search results are highlighted
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
            selected_cell: (1, 1), // Excel uses 1-based indexing
            start_row: 1,
            start_col: 1,
            visible_rows: 30, // Default values, will be adjusted based on window size
            visible_cols: 15, // Default values, will be adjusted based on window size
            input_mode: InputMode::Normal,
            input_buffer: String::new(),
            status_message: String::new(),
            should_quit: false,
            column_widths,
            clipboard: None,
            g_pressed: false,
            search_query: String::new(),
            search_results: Vec::new(),
            current_search_idx: None,
            search_direction: true, // Default to forward search
            highlight_enabled: true, // Default to showing highlights
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
        }
        else if self.selected_cell.0 >= self.start_row + self.visible_rows {
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

    pub fn start_editing(&mut self) {
        self.input_mode = InputMode::Editing;
        self.input_buffer = self.get_cell_content(self.selected_cell.0, self.selected_cell.1);
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
                self.input_buffer.clone(),
            )?;
            self.input_mode = InputMode::Normal;
            self.input_buffer = String::new();
        }
        Ok(())
    }




    pub fn jump_to_first_row(&mut self) {
        let current_col = self.selected_cell.1;
        self.selected_cell = (1, current_col);
        self.handle_scrolling();
        self.status_message = "Jumped to first row".to_string();
    }


    pub fn jump_to_last_row(&mut self) {
        let sheet = self.workbook.get_current_sheet();
        let current_col = self.selected_cell.1;


        let max_row = sheet.max_rows;

        self.selected_cell = (max_row, current_col);
        self.handle_scrolling();
        self.status_message = "Jumped to last row".to_string();
    }


    pub fn jump_to_first_column(&mut self) {
        let current_row = self.selected_cell.0;
        self.selected_cell = (current_row, 1); // First column is 1 in our system
        self.handle_scrolling();
        self.status_message = "Jumped to first column".to_string();
    }


    pub fn jump_to_first_non_empty_column(&mut self) {
        let sheet = self.workbook.get_current_sheet();
        let current_row = self.selected_cell.0;


        let mut first_non_empty_col = 1; // Default to first column

        if current_row < sheet.data.len() {
            for col in 1..=sheet.max_cols {
                if col < sheet.data[current_row].len() && !sheet.data[current_row][col].value.is_empty() {
                    first_non_empty_col = col;
                    break;
                }
            }
        }

        self.selected_cell = (current_row, first_non_empty_col);
        self.handle_scrolling();
        self.status_message = "Jumped to first non-empty column".to_string();
    }


    pub fn jump_to_last_column(&mut self) {
        let sheet = self.workbook.get_current_sheet();
        let current_row = self.selected_cell.0;


        let max_col = sheet.max_cols;

        self.selected_cell = (current_row, max_col);
        self.handle_scrolling();
        self.status_message = "Jumped to last column".to_string();
    }




    pub fn jump_to_prev_non_empty_cell_left(&mut self) {
        let sheet = self.workbook.get_current_sheet();
        let (row, col) = self.selected_cell;

        if col <= 1 {

            return;
        }


        let current_cell_is_empty = row >= sheet.data.len() ||
                                   col >= sheet.data[row].len() ||
                                   sheet.data[row][col].value.is_empty();

        if current_cell_is_empty {

            let mut target_col = 1;
            let mut found_non_empty = false;


            for c in (1..col).rev() {
                if row < sheet.data.len() && c < sheet.data[row].len() && !sheet.data[row][c].value.is_empty() {

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
            self.status_message = "Jumped to first non-empty cell (left)".to_string();
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
            self.status_message = "Jumped to last non-empty cell (left)".to_string();
        }
    }


    pub fn jump_to_prev_non_empty_cell_right(&mut self) {
        let sheet = self.workbook.get_current_sheet();
        let (row, col) = self.selected_cell;
        let max_col = sheet.max_cols;

        if col >= max_col {

            return;
        }


        let current_cell_is_empty = row >= sheet.data.len() ||
                                   col >= sheet.data[row].len() ||
                                   sheet.data[row][col].value.is_empty();

        if current_cell_is_empty {

            let mut target_col = max_col;
            let mut found_non_empty = false;


            for c in (col + 1)..=max_col {
                if row < sheet.data.len() && c < sheet.data[row].len() && !sheet.data[row][c].value.is_empty() {

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
            self.status_message = "Jumped to first non-empty cell (right)".to_string();
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
            self.status_message = "Jumped to last non-empty cell (right)".to_string();
        }
    }


    pub fn jump_to_prev_non_empty_cell_up(&mut self) {
        let sheet = self.workbook.get_current_sheet();
        let (row, col) = self.selected_cell;

        if row <= 1 {

            return;
        }


        let current_cell_is_empty = row >= sheet.data.len() ||
                                   col >= sheet.data[row].len() ||
                                   sheet.data[row][col].value.is_empty();

        if current_cell_is_empty {

            let mut target_row = 1;
            let mut found_non_empty = false;


            for r in (1..row).rev() {
                if r < sheet.data.len() && col < sheet.data[r].len() && !sheet.data[r][col].value.is_empty() {

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
            self.status_message = "Jumped to first non-empty cell (up)".to_string();
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
            self.status_message = "Jumped to last non-empty cell (up)".to_string();
        }
    }


    pub fn jump_to_prev_non_empty_cell_down(&mut self) {
        let sheet = self.workbook.get_current_sheet();
        let (row, col) = self.selected_cell;
        let max_row = sheet.max_rows;

        if row >= max_row {

            return;
        }


        let current_cell_is_empty = row >= sheet.data.len() ||
                                   col >= sheet.data[row].len() ||
                                   sheet.data[row][col].value.is_empty();

        if current_cell_is_empty {

            let mut target_row = max_row;
            let mut found_non_empty = false;


            for r in (row + 1)..=max_row {
                if r < sheet.data.len() && col < sheet.data[r].len() && !sheet.data[r][col].value.is_empty() {

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
            self.status_message = "Jumped to first non-empty cell (down)".to_string();
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
            self.status_message = "Jumped to last non-empty cell (down)".to_string();
        }
    }

    pub fn add_char_to_input(&mut self, c: char) {
        self.input_buffer.push(c);
    }

    pub fn delete_char_from_input(&mut self) {
        self.input_buffer.pop();
    }

    // Start forward search mode (/ in Vim)
    pub fn start_search_forward(&mut self) {
        self.input_mode = InputMode::SearchForward;
        self.input_buffer = String::new();
        self.status_message = String::new();
        self.highlight_enabled = true; // Ensure highlighting is enabled when starting a search
    }

    // Start backward search mode (? in Vim)
    pub fn start_search_backward(&mut self) {
        self.input_mode = InputMode::SearchBackward;
        self.input_buffer = String::new();
        self.status_message = String::new();
        self.highlight_enabled = true; // Ensure highlighting is enabled when starting a search
    }

    // Execute search based on current input buffer
    pub fn execute_search(&mut self) {
        let query = self.input_buffer.clone();
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
            self.status_message = format!("Pattern not found: {}", query);
            self.current_search_idx = None;
        } else {
            // Find the appropriate result to jump to based on search direction and current position
            self.jump_to_next_search_result();
            self.status_message = format!("{} matches found for: {}", self.search_results.len(), query);
        }

        self.input_mode = InputMode::Normal;
        self.input_buffer = String::new();
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

        if self.search_direction { // Forward search
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
                    self.status_message = "Search wrapped to top".to_string();
                }
            }
        } else { // Backward search
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
                    self.status_message = "Search wrapped to bottom".to_string();
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
        self.status_message = "Search highlighting disabled".to_string();
    }

    // Re-enable search highlighting
    pub fn enable_search_highlight(&mut self) {
        self.highlight_enabled = true;
        self.status_message = "Search highlighting enabled".to_string();
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
        // Check if there are changes to save
        if !self.workbook.is_modified() {
            self.status_message = "No changes to save".to_string();
            self.should_quit = true;
            return;
        }

        // Try to save the file
        match self.workbook.save() {
            Ok(_) => {
                self.status_message = "File saved".to_string();
                self.should_quit = true;
            }
            Err(e) => {
                self.status_message = format!("Save failed: {}", e);
                self.input_mode = InputMode::Normal;
            }
        }
    }

    // Save without exiting
    pub fn save(&mut self) -> Result<()> {
        if !self.workbook.is_modified() {
            self.status_message = "No changes to save".to_string();
            return Ok(());
        }

        match self.workbook.save() {
            Ok(_) => {
                self.status_message = "File saved".to_string();
            }
            Err(e) => {
                self.status_message = format!("Save failed: {}", e);
            }
        }
        Ok(())
    }

    pub fn exit_without_saving(&mut self) {
        self.should_quit = true;
    }

    pub fn cancel_exit(&mut self) {
        self.input_mode = InputMode::Normal;
        self.status_message = String::new();
    }

    // Switch to next sheet (without cycling)
    pub fn next_sheet(&mut self) -> Result<()> {
        let sheet_count = self.workbook.get_sheet_names().len();
        let current_index = self.workbook.get_current_sheet_index();

        // Check if we're already at the last sheet
        if current_index >= sheet_count - 1 {
            self.status_message = "Already at the last sheet".to_string();
            return Ok(());
        }

        // Move to the next sheet
        self.switch_sheet_by_index(current_index + 1)
    }

    // Switch to previous sheet (without cycling)
    pub fn prev_sheet(&mut self) -> Result<()> {
        let current_index = self.workbook.get_current_sheet_index();

        // Check if we're already at the first sheet
        if current_index == 0 {
            self.status_message = "Already at the first sheet".to_string();
            return Ok(());
        }

        // Move to the previous sheet
        self.switch_sheet_by_index(current_index - 1)
    }

    // Switch to sheet by index
    pub fn switch_sheet_by_index(&mut self, index: usize) -> Result<()> {
        // Reset cell selection and view position when switching sheets
        self.selected_cell = (1, 1);
        self.start_row = 1;
        self.start_col = 1;

        // Switch sheet in workbook
        self.workbook.switch_sheet(index)?;

        // Update column widths for the new sheet
        let max_cols = self.workbook.get_current_sheet().max_cols;
        let default_width = 15;
        self.column_widths = vec![default_width; max_cols + 1];

        // Clear search results as they're specific to the previous sheet
        self.search_results.clear();
        self.current_search_idx = None;

        self.status_message = format!("Switched to sheet: {}", self.workbook.get_current_sheet_name());
        Ok(())
    }

    // Switch to sheet by name or index from command
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
                        self.status_message = format!("Failed to switch to sheet {}: {}", index, e);
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
                        self.status_message = format!("Failed to switch to sheet '{}': {}", name_or_index, e);
                        return;
                    }
                }
            }
        }

        // If we get here, no matching sheet was found
        self.status_message = format!("Sheet '{}' not found", name_or_index);
    }



    // Copy current cell content to clipboard
    pub fn copy_cell(&mut self) {
        let content = self.get_cell_content(self.selected_cell.0, self.selected_cell.1);
        self.clipboard = Some(content);
        self.status_message = "Cell content copied".to_string();
    }

    // Cut current cell content to clipboard
    pub fn cut_cell(&mut self) -> Result<()> {
        let content = self.get_cell_content(self.selected_cell.0, self.selected_cell.1);
        self.clipboard = Some(content);

        // Clear the cell
        self.workbook.set_cell_value(
            self.selected_cell.0,
            self.selected_cell.1,
            String::new(),
        )?;

        self.status_message = "Cell content cut".to_string();
        Ok(())
    }

    // Paste clipboard content to current cell
    pub fn paste_cell(&mut self) -> Result<()> {
        if let Some(content) = &self.clipboard {
            self.workbook.set_cell_value(
                self.selected_cell.0,
                self.selected_cell.1,
                content.clone(),
            )?;
            self.status_message = "Content pasted".to_string();
        } else {
            self.status_message = "Clipboard is empty".to_string();
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

                    self.status_message = format!(
                        "Column {} width adjusted from {} to {}",
                        index_to_col_name(column),
                        old_width,
                        self.column_widths[column]
                    );
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

                    self.status_message = format!(
                        "Column {} width minimized from {} to {}",
                        index_to_col_name(column),
                        current_width,
                        new_width
                    );
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

                self.status_message = "Minimized the width of all columns".to_string();
            }
        }
    }

    fn handle_column_scrolling(&mut self) {
        let target_col = self.selected_cell.1;

        if target_col < self.start_col {
            self.start_col = target_col;
            return;
        }

        let mut current_col = self.start_col;
        let mut visible_cols = 0;

        while visible_cols < self.visible_cols {
            if current_col == target_col {
                return;
            }

            visible_cols += 1;
            current_col += 1;
        }

        self.start_col = target_col - self.visible_cols + 1;
        self.start_col = self.start_col.max(1);
    }

    pub fn ensure_column_visible(&mut self, column: usize) {
        if column < self.start_col {
            self.start_col = column;
            return;
        }

        let last_visible_col = self.start_col + self.visible_cols - 1;

        if column > last_visible_col {
            self.start_col = column - self.visible_cols + 1;
            self.start_col = self.start_col.max(1);
        }
    }

    // Enter command mode
    pub fn start_command_mode(&mut self) {
        self.input_mode = InputMode::Command;
        self.input_buffer = String::new();
        self.status_message = "Commands: :w, :wq, :q, :q!, :y, :d, :put, :cw fit, :cw fit all, :cw min, :cw min all, :cw [number], :export json [h|v] [rows], :ej [h|v] [rows], :sheet [name/number], :help".to_string();
    }

    // Handle JSON export command
    fn handle_json_export_command(&mut self, cmd: &str) {
        // Parse command
        let parts: Vec<&str> = if cmd.starts_with("export json ") {
            cmd.strip_prefix("export json ")
                .unwrap()
                .split_whitespace()
                .collect()
        } else if cmd.starts_with("ej ") {
            cmd.strip_prefix("ej ")
                .unwrap()
                .split_whitespace()
                .collect()
        } else if cmd == "export json" || cmd == "ej" {
            // No arguments provided, use default values
            vec!["h", "1"] // Default to horizontal headers with 1 header row
        } else {
            self.status_message = "Invalid JSON export command".to_string();
            return;
        };

        // Check if we have enough arguments for direction and header count
        if parts.len() < 2 {
            self.status_message = "Usage: :export json [h|v] [rows] or :ej [h|v] [rows]".to_string();
            return;
        }

        // Extract arguments
        let direction_str = parts[0];
        let header_count_str = parts[1];

        // Parse header direction
        let direction = match HeaderDirection::from_str(direction_str) {
            Some(dir) => dir,
            None => {
                self.status_message = format!(
                    "Invalid header direction: {}. Use 'h' or 'v'",
                    direction_str
                );
                return;
            }
        };

        // Parse header count
        let header_count = match header_count_str.parse::<usize>() {
            Ok(count) => count,
            Err(_) => {
                self.status_message = format!("Invalid header count: {}", header_count_str);
                return;
            }
        };

        // Get current sheet name for filename
        let sheet_name = self.workbook.get_current_sheet_name();

        // Get original file name without extension
        let file_path = self.workbook.get_file_path().to_string();
        let original_file = Path::new(&file_path);
        let file_stem = original_file.file_stem()
            .and_then(|s| s.to_str())
            .unwrap_or("export");

        // Create timestamp for the filename
        let now = chrono::Local::now();
        let timestamp = now.format("%Y%m%d_%H%M%S").to_string();

        // Create new filename with original name, sheet name and timestamp
        let new_filename = format!("{}_sheet_{}_{}.json", file_stem, sheet_name, timestamp);

        // Export to JSON
        match export_json(
            self.workbook.get_current_sheet(),
            direction,
            header_count,
            Path::new(&new_filename),
        ) {
            Ok(_) => {
                self.status_message = format!("Successfully exported to {}", new_filename);
            }
            Err(e) => {
                self.status_message = format!("Failed to export JSON: {}", e);
            }
        }
    }

    // Execute command
    pub fn execute_command(&mut self) {
        if let InputMode::Command = self.input_mode {
            let cmd = self.input_buffer.trim();

            // Handle Vim-style save and exit commands
            match cmd {
                // Save and continue editing
                "w" => {
                    if let Err(e) = self.save() {
                        self.status_message = format!("Save failed: {}", e);
                    }
                }
                // Save and quit
                "wq" | "x" => {
                    self.save_and_exit();
                }
                // Quit without saving
                "q" => {
                    if self.workbook.is_modified() {
                        self.status_message = "File modified. Use :q! to force quit without saving".to_string();
                    } else {
                        self.should_quit = true;
                    }
                }
                // Force quit without saving
                "q!" => {
                    self.exit_without_saving();
                }
                // Handle column width commands
                _ if cmd.starts_with("cw ") => {
                    if let Some(subcmd) = cmd.strip_prefix("cw ") {
                        match subcmd {
                            // Auto-adjust current column width
                            "fit" => {
                                self.auto_adjust_column_width(Some(self.selected_cell.1));
                            }
                            // Auto-adjust all column widths
                            "fit all" => {
                                self.auto_adjust_column_width(None);
                            }
                            // Minimize current column width
                            "min" => {
                                self.shrink_column_width(Some(self.selected_cell.1));
                            }
                            // Minimize all column widths
                            "min all" => {
                                self.shrink_column_width(None);
                            }
                            // Try to parse subcommand as a number to set current column width
                            _ => {
                                if let Ok(width) = subcmd.parse::<usize>() {
                                    let min_width = 5;
                                    let max_width = 100;
                                    let column = self.selected_cell.1;

                                    // Ensure width is within reasonable range
                                    let width = width.max(min_width).min(max_width);

                                    if column < self.column_widths.len() {
                                        // Record original width for status message
                                        let old_width = self.column_widths[column];

                                        // Record starting column to detect changes (unused but kept for future use)
                                        let _original_start_col = self.start_col;

                                        // Set new column width
                                        self.column_widths[column] = width;

                                        self.ensure_column_visible(column);

                                        self.status_message = format!(
                                            "Column {} width changed from {} to {}",
                                            index_to_col_name(column),
                                            old_width,
                                            width
                                        );
                                    }
                                } else {
                                    self.status_message = format!("Unknown column command: {}", subcmd);
                                }
                            }
                        }
                    }
                }
                // JSON export command
                _ if cmd.starts_with("export json ") || cmd.starts_with("ej ") => {
                    let cmd_str = cmd.to_string(); // Clone the command string to avoid borrowing issues
                    self.handle_json_export_command(&cmd_str);
                }
                // Sheet switching command
                _ if cmd.starts_with("sheet ") => {
                    if let Some(sheet_name_or_index) = cmd.strip_prefix("sheet ") {
                        let name_or_index = sheet_name_or_index.trim().to_string();
                        self.switch_to_sheet(&name_or_index);
                    }
                }

                // Copy command
                "y" => {
                    self.copy_cell();
                }
                // Cut command
                "d" => {
                    if let Err(e) = self.cut_cell() {
                        self.status_message = format!("Cut failed: {}", e);
                    }
                }
                // Paste command (using Vim's Ex mode paste command)
                "put" | "pu" => {
                    if let Err(e) = self.paste_cell() {
                        self.status_message = format!("Paste failed: {}", e);
                    }
                }
                // Disable search highlighting
                "nohlsearch" | "noh" => {
                    self.disable_search_highlight();
                }
                // Help command
                "help" => {
                    self.status_message = format!(
                        "Commands:\n\
                         :w - Save file\n\
                         :wq, :x - Save and quit\n\
                         :q - Quit (will warn if unsaved changes)\n\
                         :q! - Force quit without saving\n\
                         :y - Copy current cell\n\
                         :d - Cut current cell\n\
                         :put, :pu - Paste to current cell\n\
                         :[cell] - Jump to cell (e.g., :A1, :B10)\n\
                         :nohlsearch, :noh - Disable search highlighting\n\
                         :cw fit, :cw fit all, :cw min, :cw min all, :cw [number] - Column width commands\n\
                         :export json [h|v] [rows] - Export to JSON (h=horizontal, v=vertical)\n\
                         :ej [h|v] [rows] - Shorthand for export json\n\
                         :sheet [name/number] - Switch to sheet by name or index"
                    );
                }
                // Try to parse as cell reference (e.g., A1, B10)
                _ => {
                    // Clone cmd to avoid borrowing issues
                    let cmd_clone = cmd.to_string();
                    if let Some(cell_ref) = parse_cell_reference(cmd) {
                        self.selected_cell = cell_ref;
                        self.handle_scrolling();
                        self.status_message = format!("Jumped to cell {}", cmd_clone);
                    } else {
                        self.status_message = format!("Unknown command: {}", cmd_clone);
                    }
                }
            }

            self.input_mode = InputMode::Normal;
            self.input_buffer = String::new();
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
