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

        // Save current column widths if they've changed
        if !self.sheet_column_widths.contains_key(&current_sheet_name)
            || self.sheet_column_widths[&current_sheet_name] != self.column_widths
        {
            self.sheet_column_widths
                .insert(current_sheet_name.clone(), self.column_widths.clone());
        }

        // Save current cell position and view position
        let current_position = crate::app::CellPosition {
            selected: self.selected_cell,
            view: (self.start_row, self.start_col),
        };
        self.sheet_cell_positions
            .insert(current_sheet_name, current_position);

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

        // Restore cell position and view position for the new sheet
        if let Some(saved_position) = self.sheet_cell_positions.get(&new_sheet_name) {
            // Ensure the saved position is valid for the current sheet
            let sheet = self.workbook.get_current_sheet();
            let valid_row = saved_position.selected.0.min(sheet.max_rows.max(1));
            let valid_col = saved_position.selected.1.min(sheet.max_cols.max(1));

            self.selected_cell = (valid_row, valid_col);
            self.start_row = saved_position.view.0;
            self.start_col = saved_position.view.1;

            // Make sure the view position is valid relative to the selected cell
            self.handle_scrolling();
        } else {
            // If no saved position exists, use default position
            self.selected_cell = (1, 1);
            self.start_row = 1;
            self.start_col = 1;
        }

        // Clear search results as they're specific to the previous sheet
        if !self.search_results.is_empty() {
            self.search_results.clear();
            self.current_search_idx = None;
        }

        self.update_row_number_width();

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
                self.sheet_cell_positions.remove(&current_sheet_name);

                let new_sheet_name = self.workbook.get_current_sheet_name();

                // Restore saved cell position for the new current sheet or use default
                if let Some(saved_position) = self.sheet_cell_positions.get(&new_sheet_name) {
                    // Ensure the saved position is valid for the current sheet
                    let sheet = self.workbook.get_current_sheet();
                    let valid_row = saved_position.selected.0.min(sheet.max_rows.max(1));
                    let valid_col = saved_position.selected.1.min(sheet.max_cols.max(1));

                    self.selected_cell = (valid_row, valid_col);
                    self.start_row = saved_position.view.0;
                    self.start_col = saved_position.view.1;

                    // Make sure the view position is valid relative to the selected cell
                    self.handle_scrolling();
                } else {
                    // If no saved position exists, use default position
                    self.selected_cell = (1, 1);
                    self.start_row = 1;
                    self.start_col = 1;
                }

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
        let sheet = self.workbook.get_current_sheet();

        // If row is outside the valid range, return success
        if row < 1 || row > sheet.max_rows {
            return Ok(());
        }

        // Save row data for undo
        let sheet_index = self.workbook.get_current_sheet_index();
        let sheet_name = self.workbook.get_current_sheet_name();

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

        self.workbook.recalculate_max_rows();
        self.workbook.recalculate_max_cols();
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

    pub fn delete_row(&mut self, row: usize) -> Result<()> {
        let sheet = self.workbook.get_current_sheet();

        // If row is outside the valid range, return success
        if row < 1 || row > sheet.max_rows {
            return Ok(());
        }

        // Save row data for undo
        let sheet_index = self.workbook.get_current_sheet_index();
        let sheet_name = self.workbook.get_current_sheet_name();

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

        self.workbook.recalculate_max_rows();
        self.workbook.recalculate_max_cols();
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

        let sheet = self.workbook.get_current_sheet();

        // If the entire range is outside the valid range, return success
        if start_row < 1 || start_row > sheet.max_rows || start_row > end_row {
            return Ok(());
        }

        // If start_row is valid but end_row exceeds max_rows, adjust end_row to max_rows
        let effective_end_row = end_row.min(sheet.max_rows);

        // Save all row data for batch undo
        let sheet_index = self.workbook.get_current_sheet_index();
        let sheet_name = self.workbook.get_current_sheet_name();

        // Save row data in the original order from top to bottom
        let rows_to_save = effective_end_row - start_row + 1;
        let mut rows_data = Vec::with_capacity(rows_to_save);

        for row in start_row..=effective_end_row {
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
            end_row: effective_end_row,
            rows_data,
        };

        self.undo_history
            .push(ActionCommand::MultiRow(multi_row_action));
        self.workbook.delete_rows(start_row, effective_end_row)?;

        self.workbook.recalculate_max_rows();
        self.workbook.recalculate_max_cols();
        let sheet = self.workbook.get_current_sheet();

        if self.selected_cell.0 > sheet.max_rows {
            self.selected_cell.0 = sheet.max_rows.max(1);
        }

        self.handle_scrolling();
        self.search_results.clear();
        self.current_search_idx = None;

        self.add_notification(format!(
            "Deleted rows {} to {}",
            start_row, effective_end_row
        ));
        Ok(())
    }

    pub fn delete_current_column(&mut self) -> Result<()> {
        let col = self.selected_cell.1;
        let sheet = self.workbook.get_current_sheet();

        // If column is outside the valid range, return success
        if col < 1 || col > sheet.max_cols {
            return Ok(());
        }

        // Save column data for undo
        let sheet_index = self.workbook.get_current_sheet_index();
        let sheet_name = self.workbook.get_current_sheet_name();

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

        self.workbook.recalculate_max_rows();
        self.workbook.recalculate_max_cols();
        let sheet = self.workbook.get_current_sheet();

        if col > sheet.max_cols {
            self.selected_cell.1 = sheet.max_cols.max(1);
        }

        if self.selected_cell.0 > sheet.max_rows {
            self.selected_cell.0 = sheet.max_rows.max(1);
        }

        if self.column_widths.len() > col {
            self.column_widths.remove(col);
        }

        self.adjust_column_widths(sheet.max_cols);

        self.handle_scrolling();
        self.search_results.clear();
        self.current_search_idx = None;

        self.add_notification(format!("Deleted column {}", index_to_col_name(col)));
        Ok(())
    }

    pub fn delete_column(&mut self, col: usize) -> Result<()> {
        let sheet = self.workbook.get_current_sheet();

        // If column is outside the valid range, return success
        if col < 1 || col > sheet.max_cols {
            return Ok(());
        }

        // Save column data for undo
        let sheet_index = self.workbook.get_current_sheet_index();
        let sheet_name = self.workbook.get_current_sheet_name();

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

        self.workbook.recalculate_max_rows();
        self.workbook.recalculate_max_cols();
        let sheet = self.workbook.get_current_sheet();

        if self.selected_cell.1 > sheet.max_cols {
            self.selected_cell.1 = sheet.max_cols.max(1);
        }

        if self.selected_cell.0 > sheet.max_rows {
            self.selected_cell.0 = sheet.max_rows.max(1);
        }

        if self.column_widths.len() > col {
            self.column_widths.remove(col);
        }

        self.adjust_column_widths(sheet.max_cols);

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

        let sheet = self.workbook.get_current_sheet();

        // If the entire range is outside the valid range, return success
        if start_col < 1 || start_col > sheet.max_cols || start_col > end_col {
            return Ok(());
        }

        // If start_col is valid but end_col exceeds max_cols, adjust end_col to max_cols
        let effective_end_col = end_col.min(sheet.max_cols);

        // For multiple columns, save all column data for batch undo
        let sheet_index = self.workbook.get_current_sheet_index();
        let sheet_name = self.workbook.get_current_sheet_name();

        // Save column data and widths for batch undo
        let cols_to_save = effective_end_col - start_col + 1;
        let mut columns_data = Vec::with_capacity(cols_to_save);
        let mut column_widths = Vec::with_capacity(cols_to_save);

        for col in start_col..=effective_end_col {
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
            end_col: effective_end_col,
            columns_data,
            column_widths,
        };

        self.undo_history
            .push(ActionCommand::MultiColumn(multi_column_action));
        self.workbook.delete_columns(start_col, effective_end_col)?;

        self.workbook.recalculate_max_rows();
        self.workbook.recalculate_max_cols();
        let sheet = self.workbook.get_current_sheet();

        if self.selected_cell.1 > sheet.max_cols {
            self.selected_cell.1 = sheet.max_cols.max(1);
        }

        if self.selected_cell.0 > sheet.max_rows {
            self.selected_cell.0 = sheet.max_rows.max(1);
        }

        for col in (start_col..=effective_end_col).rev() {
            if self.column_widths.len() > col {
                self.column_widths.remove(col);
            }
        }

        self.adjust_column_widths(sheet.max_cols);

        self.handle_scrolling();
        self.search_results.clear();
        self.current_search_idx = None;

        self.add_notification(format!(
            "Deleted columns {} to {}",
            index_to_col_name(start_col),
            index_to_col_name(effective_end_col)
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

            let mut display_width = 0;

            for c in content.chars() {
                if c.is_ascii() {
                    display_width += 1;
                } else {
                    display_width += 2;
                }
            }

            max_width = max_width.max(display_width);
        }
        max_width
    }

    pub fn get_column_width(&self, col: usize) -> usize {
        if col < self.column_widths.len() {
            self.column_widths[col]
        } else {
            15 // Default width
        }
    }

    pub fn ensure_column_widths(&mut self) {
        let sheet = self.workbook.get_current_sheet();
        self.adjust_column_widths(sheet.max_cols);
    }

    fn adjust_column_widths(&mut self, max_cols: usize) {
        match self.column_widths.len().cmp(&(max_cols + 1)) {
            std::cmp::Ordering::Greater => {
                self.column_widths.truncate(max_cols + 1);
            }
            std::cmp::Ordering::Less => {
                let additional = max_cols + 1 - self.column_widths.len();
                self.column_widths.extend(vec![15; additional]);
            }
            std::cmp::Ordering::Equal => {
                // Column widths are already correct, do nothing
            }
        }
    }
}
