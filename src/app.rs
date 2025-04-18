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
        
        // Handle column scrolling with column width consideration
        if self.selected_cell.1 < self.start_col {
            // If selected column is before visible area, make it the first visible column
            self.start_col = self.selected_cell.1;
        } else {
            // Check if selected column is visible with current start_col
            let mut width_sum = 0;
            let mut last_visible_col = self.start_col;
            
            // Calculate which columns are visible
            for col in self.start_col.. {
                width_sum += self.get_column_width(col);
                
                // If we've accumulated enough width to fill the visible area
                if width_sum >= self.visible_cols * 15 { // Approximate width threshold
                    last_visible_col = col;
                    break;
                }
            }
            
            // If selected column is after last visible column, adjust start_col
            if self.selected_cell.1 > last_visible_col {
                // Move start_col forward until selected column becomes visible
                // This is a simplified approach - a more precise calculation would involve
                // working backwards from the selected column
                self.start_col = self.selected_cell.1 - (self.visible_cols / 2);
                self.start_col = self.start_col.max(1); // Ensure start_col is at least 1
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
                // Calculate display width considering Unicode width
                let display_width = content.chars().fold(0, |acc, c| {
                    acc + if c.is_ascii() { 1 } else { 2 } // Consider CJK characters as double-width
                });
                max_width = max_width.max(display_width);
            }
        }
        
        // Add more generous padding to ensure content fits
        max_width + 6
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