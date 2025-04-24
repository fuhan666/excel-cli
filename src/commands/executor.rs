use std::path::Path;

use crate::app::AppState;
use crate::json_export::{HeaderDirection, export_all_sheets_json, export_json};
use crate::utils::col_name_to_index;

impl AppState<'_> {
    pub fn execute_command(&mut self) {
        let command = self.input_buffer.clone();
        self.input_mode = crate::app::InputMode::Normal;
        self.input_buffer = String::new();

        if command.is_empty() {
            return;
        }

        // Handle cell navigation (e.g., :A1, :B10)
        if let Some(cell_ref) = parse_cell_reference(&command) {
            self.jump_to_cell(cell_ref);
            return;
        }

        // Handle commands
        match command.as_str() {
            "w" => {
                if let Err(e) = self.save() {
                    self.add_notification(format!("Save failed: {}", e));
                }
            }
            "wq" | "x" => self.save_and_exit(),
            "q" => {
                if self.workbook.is_modified() {
                    self.add_notification(
                        "File has unsaved changes. Use :q! to force quit or :wq to save and quit."
                            .to_string(),
                    );
                } else {
                    self.should_quit = true;
                }
            }
            "q!" => self.exit_without_saving(),
            "y" => self.copy_cell(),
            "d" => {
                if let Err(e) = self.cut_cell() {
                    self.add_notification(format!("Cut failed: {}", e));
                }
            }
            "put" | "pu" => {
                if let Err(e) = self.paste_cell() {
                    self.add_notification(format!("Paste failed: {}", e));
                }
            }
            "nohlsearch" | "noh" => self.disable_search_highlight(),
            "help" => self.show_help(),
            "delsheet" => self.delete_current_sheet(),
            _ => {
                // Handle commands with parameters
                if command.starts_with("cw ") {
                    self.handle_column_width_command(&command);
                } else if command.starts_with("ej") {
                    self.handle_json_export_command(&command);
                } else if command.starts_with("sheet ") {
                    let sheet_name = command.strip_prefix("sheet ").unwrap().trim();
                    self.switch_to_sheet(sheet_name);
                } else if command.starts_with("dr") {
                    self.handle_delete_row_command(&command);
                } else if command.starts_with("dc") {
                    self.handle_delete_column_command(&command);
                } else {
                    self.add_notification(format!("Unknown command: {}", command));
                }
            }
        }
    }

    fn handle_column_width_command(&mut self, cmd: &str) {
        let parts: Vec<&str> = cmd.split_whitespace().collect();

        if parts.len() < 2 {
            self.add_notification("Usage: :cw [fit|min|number] [all]".to_string());
            return;
        }

        let action = parts[1];
        let apply_to_all = parts.len() > 2 && parts[2] == "all";

        match action {
            "fit" => {
                if apply_to_all {
                    self.auto_adjust_column_width(None);
                } else {
                    self.auto_adjust_column_width(Some(self.selected_cell.1));
                }
            }
            "min" => {
                if apply_to_all {
                    // Set all columns to minimum width
                    let sheet = self.workbook.get_current_sheet();
                    for col in 1..=sheet.max_cols {
                        self.column_widths[col] = 5; // Minimum width
                    }
                    self.add_notification("All columns set to minimum width".to_string());
                } else {
                    // Set current column to minimum width
                    let col = self.selected_cell.1;
                    self.column_widths[col] = 5; // Minimum width
                    self.add_notification(format!("Column {} set to minimum width", col));
                }
            }
            _ => {
                // Try to parse as a number
                if let Ok(width) = action.parse::<usize>() {
                    let col = self.selected_cell.1;
                    self.column_widths[col] = width.clamp(5, 50); // Clamp between 5 and 50
                    self.add_notification(format!("Column {} width set to {}", col, width));
                } else {
                    self.add_notification(format!("Invalid column width: {}", action));
                }
            }
        }
    }

    fn handle_delete_row_command(&mut self, cmd: &str) {
        let parts: Vec<&str> = cmd.split_whitespace().collect();

        if parts.len() == 1 {
            // Delete current row
            if let Err(e) = self.delete_current_row() {
                self.add_notification(format!("Failed to delete row: {}", e));
            }
            return;
        }

        if parts.len() == 2 {
            // Delete specific row
            if let Ok(row) = parts[1].parse::<usize>() {
                if let Err(e) = self.delete_row(row) {
                    self.add_notification(format!("Failed to delete row {}: {}", row, e));
                }
            } else {
                self.add_notification(format!("Invalid row number: {}", parts[1]));
            }
            return;
        }

        if parts.len() == 3 {
            // Delete range of rows
            if let (Ok(start_row), Ok(end_row)) =
                (parts[1].parse::<usize>(), parts[2].parse::<usize>())
            {
                if let Err(e) = self.delete_rows(start_row, end_row) {
                    self.add_notification(format!(
                        "Failed to delete rows {} to {}: {}",
                        start_row, end_row, e
                    ));
                }
            } else {
                self.add_notification("Invalid row range".to_string());
            }
            return;
        }

        self.add_notification("Usage: :dr [row] [end_row]".to_string());
    }

    fn handle_delete_column_command(&mut self, cmd: &str) {
        let parts: Vec<&str> = cmd.split_whitespace().collect();

        if parts.len() == 1 {
            // Delete current column
            if let Err(e) = self.delete_current_column() {
                self.add_notification(format!("Failed to delete column: {}", e));
            }
            return;
        }

        if parts.len() == 2 {
            // Delete specific column
            let col_str = parts[1].to_uppercase();

            // Try to parse as a column letter (A, B, C, etc.)
            if let Some(col) = col_name_to_index(&col_str) {
                if let Err(e) = self.delete_column(col) {
                    self.add_notification(format!("Failed to delete column {}: {}", col_str, e));
                }
                return;
            }

            // Try to parse as a column number
            if let Ok(col) = col_str.parse::<usize>() {
                if let Err(e) = self.delete_column(col) {
                    self.add_notification(format!("Failed to delete column {}: {}", col, e));
                }
                return;
            }

            self.add_notification(format!("Invalid column: {}", col_str));
            return;
        }

        if parts.len() == 3 {
            // Delete range of columns
            let start_col_str = parts[1].to_uppercase();
            let end_col_str = parts[2].to_uppercase();

            let start_col =
                col_name_to_index(&start_col_str).or_else(|| start_col_str.parse::<usize>().ok());
            let end_col =
                col_name_to_index(&end_col_str).or_else(|| end_col_str.parse::<usize>().ok());

            if let (Some(start), Some(end)) = (start_col, end_col) {
                if let Err(e) = self.delete_columns(start, end) {
                    self.add_notification(format!(
                        "Failed to delete columns {} to {}: {}",
                        start_col_str, end_col_str, e
                    ));
                }
            } else {
                self.add_notification("Invalid column range".to_string());
            }
            return;
        }

        self.add_notification("Usage: :dc [col] [end_col]".to_string());
    }

    fn handle_json_export_command(&mut self, cmd: &str) {
        // Check if this is an export all command
        let export_all = cmd.starts_with("eja ") || cmd == "eja";

        // Parse command
        let parts: Vec<&str> = if cmd.starts_with("ej ") {
            cmd.strip_prefix("ej ")
                .unwrap()
                .split_whitespace()
                .collect()
        } else if cmd == "ej" {
            // No arguments provided, use default values
            vec!["h", "1"] // Default to horizontal headers with 1 header row
        } else if cmd.starts_with("eja ") {
            cmd.strip_prefix("eja ")
                .unwrap()
                .split_whitespace()
                .collect()
        } else if cmd == "eja" {
            // No arguments provided, use default values
            vec!["h", "1"] // Default to horizontal headers with 1 header row
        } else {
            self.add_notification("Invalid JSON export command".to_string());
            return;
        };

        // Check if we have enough arguments for direction and header count
        if parts.len() < 2 {
            if export_all {
                self.add_notification("Usage: :eja [h|v] [rows]".to_string());
            } else {
                self.add_notification("Usage: :ej [h|v] [rows]".to_string());
            }
            return;
        }

        let direction_str = parts[0];
        let header_count_str = parts[1];

        let direction = match HeaderDirection::from_str(direction_str) {
            Some(dir) => dir,
            None => {
                self.add_notification(format!(
                    "Invalid header direction: {}. Use 'h' or 'v'",
                    direction_str
                ));
                return;
            }
        };

        let header_count = match header_count_str.parse::<usize>() {
            Ok(count) => count,
            Err(_) => {
                self.add_notification(format!("Invalid header count: {}", header_count_str));
                return;
            }
        };

        let sheet_name = self.workbook.get_current_sheet_name();

        // Get original file name without extension
        let file_path = self.workbook.get_file_path().to_string();
        let original_file = Path::new(&file_path);
        let file_stem = original_file
            .file_stem()
            .and_then(|s| s.to_str())
            .unwrap_or("export");

        let now = chrono::Local::now();
        let timestamp = now.format("%Y%m%d_%H%M%S").to_string();

        let new_filename = if export_all {
            format!("{}_all_sheets_{}.json", file_stem, timestamp)
        } else {
            format!("{}_sheet_{}_{}.json", file_stem, sheet_name, timestamp)
        };

        // Export to JSON
        let result = if export_all {
            export_all_sheets_json(
                &self.workbook,
                direction,
                header_count,
                Path::new(&new_filename),
            )
        } else {
            export_json(
                self.workbook.get_current_sheet(),
                direction,
                header_count,
                Path::new(&new_filename),
            )
        };

        match result {
            Ok(_) => {
                self.add_notification(format!("Exported to {}", new_filename));
            }
            Err(e) => {
                self.add_notification(format!("Export failed: {}", e));
            }
        }
    }

    fn jump_to_cell(&mut self, cell_ref: (usize, usize)) {
        let (row, col) = cell_ref; // Fixed: cell_ref is already (row, col)

        let sheet = self.workbook.get_current_sheet();

        // Validate row and column
        if row > sheet.max_rows || col > sheet.max_cols {
            self.add_notification(format!(
                "Cell reference out of range: {}{}",
                crate::utils::index_to_col_name(col),
                row
            ));
            return;
        }

        self.selected_cell = (row, col);
        // Handle scrolling
        if self.selected_cell.0 < self.start_row {
            self.start_row = self.selected_cell.0;
        } else if self.selected_cell.0 >= self.start_row + self.visible_rows {
            self.start_row = self.selected_cell.0 - self.visible_rows + 1;
        }

        self.ensure_column_visible(self.selected_cell.1);

        self.add_notification(format!(
            "Jumped to cell {}{}",
            crate::utils::index_to_col_name(col),
            row
        ));
    }
}

// Parse a cell reference like "A1", "B10", etc.
fn parse_cell_reference(input: &str) -> Option<(usize, usize)> {
    // Cell references should have at least 2 characters (e.g., A1)
    if input.len() < 2 {
        return None;
    }

    // Find the first digit to separate column and row parts
    let mut col_end = 0;
    for (i, c) in input.chars().enumerate() {
        if c.is_ascii_digit() {
            col_end = i;
            break;
        }
    }

    if col_end == 0 {
        return None; // No digits found
    }

    let col_part = &input[0..col_end];
    let row_part = &input[col_end..];

    // Convert column letters to index
    let col = col_name_to_index(&col_part.to_uppercase())?;

    // Parse row number
    let row = row_part.parse::<usize>().ok()?;

    Some((row, col))
}
