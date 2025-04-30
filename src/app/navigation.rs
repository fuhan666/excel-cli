use crate::app::AppState;
use crate::utils::find_non_empty_cell;
use crate::utils::Direction;

impl AppState<'_> {
    pub fn move_cursor(&mut self, delta_row: isize, delta_col: isize) {
        // Calculate new position
        let new_row = (self.selected_cell.0 as isize + delta_row).max(1) as usize;
        let new_col = (self.selected_cell.1 as isize + delta_col).max(1) as usize;

        // Update selected position
        self.selected_cell = (new_row, new_col);

        // Handle scrolling
        self.handle_scrolling();
    }

    pub fn handle_scrolling(&mut self) {
        if self.selected_cell.0 < self.start_row {
            self.start_row = self.selected_cell.0;
        } else if self.selected_cell.0 >= self.start_row + self.visible_rows {
            self.start_row = self.selected_cell.0 - self.visible_rows + 1;
        }

        self.handle_column_scrolling();
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
        self.selected_cell = (current_row, 1);
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
                format!("Jumped to first non-empty cell ({dir_name})")
            } else {
                format!("Jumped to last non-empty cell ({dir_name})")
            };

            self.add_notification(message);
        }
    }

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
}
