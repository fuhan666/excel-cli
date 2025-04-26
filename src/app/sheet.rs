use crate::actions::{
    ActionCommand, ColumnAction, MultiColumnAction, MultiRowAction, RowAction, SheetAction,
};
use crate::app::AppState;
use crate::utils::index_to_col_name;
use anyhow::Result;

impl AppState<'_> {
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

                self.undo_history.push(ActionCommand::Sheet(sheet_action));
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

        self.undo_history.push(ActionCommand::Row(row_action));
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

        self.undo_history.push(ActionCommand::Row(row_action));
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
        let rows_to_save = end_row - start_row + 1;
        let mut rows_data = Vec::with_capacity(rows_to_save);

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

        self.undo_history
            .push(ActionCommand::MultiRow(multi_row_action));
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
        let mut column_data = Vec::with_capacity(sheet.data.len());
        for row in &sheet.data {
            if col < row.len() {
                column_data.push(row[col].clone());
            } else {
                column_data.push(crate::excel::Cell::empty());
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

        self.undo_history.push(ActionCommand::Column(column_action));
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
        let mut column_data = Vec::with_capacity(sheet.data.len());
        for row in &sheet.data {
            if col < row.len() {
                column_data.push(row[col].clone());
            } else {
                column_data.push(crate::excel::Cell::empty());
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

        self.undo_history.push(ActionCommand::Column(column_action));
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
        let cols_to_save = end_col - start_col + 1;
        let mut columns_data = Vec::with_capacity(cols_to_save);
        let mut column_widths = Vec::with_capacity(cols_to_save);

        for col in start_col..=end_col {
            // Extract the column data from each row
            let mut column_data = Vec::with_capacity(sheet.data.len());
            for row in &sheet.data {
                if col < row.len() {
                    column_data.push(row[col].clone());
                } else {
                    column_data.push(crate::excel::Cell::empty());
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

        self.undo_history
            .push(ActionCommand::MultiColumn(multi_column_action));
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
        let mut defaults = vec![15; cols_to_remove];
        self.column_widths.append(&mut defaults);

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
}
