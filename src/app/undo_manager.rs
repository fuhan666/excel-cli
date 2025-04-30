use crate::actions::{
    ActionCommand, ActionExecutor, ActionType, CellAction, ColumnAction, MultiColumnAction,
    MultiRowAction, RowAction, SheetAction,
};
use crate::app::AppState;
use crate::utils::index_to_col_name;
use anyhow::Result;
use std::rc::Rc;

impl AppState<'_> {
    pub fn undo(&mut self) -> Result<()> {
        if let Some(action) = self.undo_history.undo() {
            self.apply_action(&action, true)?;

            self.workbook.recalculate_max_rows();
            self.workbook.recalculate_max_cols();
            self.ensure_column_widths();

            // Update cursor position if it's outside the valid range
            let sheet = self.workbook.get_current_sheet();
            if self.selected_cell.0 > sheet.max_rows {
                self.selected_cell.0 = sheet.max_rows.max(1);
            }
            if self.selected_cell.1 > sheet.max_cols {
                self.selected_cell.1 = sheet.max_cols.max(1);
            }

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

            self.workbook.recalculate_max_rows();
            self.workbook.recalculate_max_cols();
            self.ensure_column_widths();

            // Update cursor position if it's outside the valid range
            let sheet = self.workbook.get_current_sheet();
            if self.selected_cell.0 > sheet.max_rows {
                self.selected_cell.0 = sheet.max_rows.max(1);
            }
            if self.selected_cell.1 > sheet.max_cols {
                self.selected_cell.1 = sheet.max_cols.max(1);
            }

            self.workbook.set_modified(true);
        } else {
            self.add_notification("No operations to redo".to_string());
        }
        Ok(())
    }

    fn apply_action(&mut self, action: &Rc<ActionCommand>, is_undo: bool) -> Result<()> {
        match action.as_ref() {
            ActionCommand::Cell(cell_action) => {
                let value = if is_undo {
                    &cell_action.old_value
                } else {
                    &cell_action.new_value
                };
                self.apply_cell_action(cell_action, value, is_undo, &cell_action.action_type)?;
            }
            ActionCommand::Row(row_action) => {
                self.apply_row_action(row_action, is_undo)?;
            }
            ActionCommand::Column(column_action) => {
                self.apply_column_action(column_action, is_undo)?;
            }
            ActionCommand::Sheet(sheet_action) => {
                self.apply_sheet_action(sheet_action, is_undo)?;
            }
            ActionCommand::MultiRow(multi_row_action) => {
                self.apply_multi_row_action(multi_row_action, is_undo)?;
            }
            ActionCommand::MultiColumn(multi_column_action) => {
                self.apply_multi_column_action(multi_column_action, is_undo)?;
            }
        }
        Ok(())
    }

    fn apply_cell_action(
        &mut self,
        cell_action: &CellAction,
        value: &crate::excel::Cell,
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

            // Recalculate max_cols since restoring a row might affect the maximum column count
            // This is especially important if the row contained data beyond the current max_cols
            self.workbook.recalculate_max_cols();

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
                            row.push(crate::excel::Cell::empty());
                        }
                        row.push(column_data[i].clone());
                    }
                }
            }

            // Update both max_cols and max_rows when restoring a column
            sheet.max_cols = sheet.max_cols.saturating_add(1);

            // Recalculate max_rows since restoring a column might affect the maximum row count
            // This is especially important if the column contained data beyond the current max_rows
            self.workbook.recalculate_max_rows();

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

            // Initialize cell position for the restored sheet with default values
            self.sheet_cell_positions.insert(
                sheet_action.sheet_name.clone(),
                crate::app::CellPosition {
                    selected: (1, 1),
                    view: (1, 1),
                },
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
                self.add_notification(format!("Failed to delete sheet: {e}"));
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
        self.sheet_cell_positions.remove(sheet_name);

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

            // Optimized restore function
            Self::restore_rows(sheet, start_row, rows_data);

            sheet.max_rows = sheet.max_rows.saturating_add(rows_to_restore);

            // Recalculate max_cols since restoring rows might affect the maximum column count
            self.workbook.recalculate_max_cols();

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

            // Recalculate max_rows since restoring columns might affect the maximum row count
            self.workbook.recalculate_max_rows();

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

    fn restore_rows(
        sheet: &mut crate::excel::Sheet,
        position: usize,
        rows_data: &[Vec<crate::excel::Cell>],
    ) {
        // Pre-allocate space by extending the vector
        for row_data in rows_data.iter().rev() {
            sheet.data.insert(position, row_data.clone());
        }
    }

    fn restore_column_at_position(
        sheet: &mut crate::excel::Sheet,
        position: usize,
        column_data: &[crate::excel::Cell],
    ) {
        for (i, row) in sheet.data.iter_mut().enumerate() {
            if i < column_data.len() {
                if position <= row.len() {
                    row.insert(position, column_data[i].clone());
                } else {
                    let additional = position - row.len();
                    row.reserve(additional + 1);
                    while row.len() < position {
                        row.push(crate::excel::Cell::empty());
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

    fn trim_column_widths(column_widths: &mut Vec<usize>, count: usize) {
        if count >= column_widths.len() {
            return;
        }
        column_widths.truncate(column_widths.len() - count);
    }

    fn remove_column_widths(column_widths: &mut Vec<usize>, start_col: usize, end_col: usize) {
        let cols_to_remove = end_col - start_col + 1;

        // Pre-allocate space for new entries to avoid multiple resizes
        column_widths.reserve(cols_to_remove);

        for col in (start_col..=end_col).rev() {
            if column_widths.len() > col {
                column_widths.remove(col);
            }
        }

        // Add default widths in a single batch to avoid multiple resizes
        let mut defaults = vec![15; cols_to_remove];
        column_widths.append(&mut defaults);
    }
}

impl ActionExecutor for AppState<'_> {
    fn execute_action(&mut self, action: &ActionCommand) -> Result<()> {
        match action {
            ActionCommand::Cell(action) => self.execute_cell_action(action),
            ActionCommand::Row(action) => self.execute_row_action(action),
            ActionCommand::Column(action) => self.execute_column_action(action),
            ActionCommand::Sheet(action) => self.execute_sheet_action(action),
            ActionCommand::MultiRow(action) => self.execute_multi_row_action(action),
            ActionCommand::MultiColumn(action) => self.execute_multi_column_action(action),
        }
    }

    fn execute_cell_action(&mut self, action: &CellAction) -> Result<()> {
        self.workbook
            .set_cell_value(action.row, action.col, action.new_value.value.clone())
    }

    fn execute_row_action(&mut self, action: &RowAction) -> Result<()> {
        self.workbook.delete_row(action.row)
    }

    fn execute_column_action(&mut self, action: &ColumnAction) -> Result<()> {
        self.workbook.delete_column(action.col)
    }

    fn execute_sheet_action(&mut self, action: &SheetAction) -> Result<()> {
        self.switch_sheet_by_index(action.sheet_index)?;
        self.workbook.delete_current_sheet()
    }

    fn execute_multi_row_action(&mut self, action: &MultiRowAction) -> Result<()> {
        self.workbook.delete_rows(action.start_row, action.end_row)
    }

    fn execute_multi_column_action(&mut self, action: &MultiColumnAction) -> Result<()> {
        self.workbook
            .delete_columns(action.start_col, action.end_col)
    }
}
