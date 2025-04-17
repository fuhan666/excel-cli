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
}

impl AppState {
    pub fn new(workbook: Workbook, file_path: PathBuf) -> Result<Self> {
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
        // If selected cell is not in visible area, scroll
        if self.selected_cell.0 < self.start_row {
            self.start_row = self.selected_cell.0;
        } else if self.selected_cell.0 >= self.start_row + self.visible_rows {
            self.start_row = self.selected_cell.0 - self.visible_rows + 1;
        }
        
        if self.selected_cell.1 < self.start_col {
            self.start_col = self.selected_cell.1;
        } else if self.selected_cell.1 >= self.start_col + self.visible_cols {
            self.start_col = self.selected_cell.1 - self.visible_cols + 1;
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