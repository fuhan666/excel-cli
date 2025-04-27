use crate::actions::{ActionCommand, ActionType, CellAction};
use crate::app::AppState;
use crate::app::InputMode;
use anyhow::Result;

impl AppState<'_> {
    pub fn start_editing(&mut self) {
        self.input_mode = InputMode::Editing;
        let content = self.get_cell_content(self.selected_cell.0, self.selected_cell.1);
        self.input_buffer = content.clone();

        // Set up TextArea for editing
        self.text_area = ratatui_textarea::TextArea::default();
        self.text_area.insert_str(&content);
    }

    pub fn confirm_edit(&mut self) -> Result<()> {
        if let InputMode::Editing = self.input_mode {
            // Get content from TextArea
            let content = self.text_area.lines().join("\n");
            let (row, col) = self.selected_cell;

            self.workbook.ensure_cell_exists(row, col);

            self.ensure_column_widths();

            let sheet_index = self.workbook.get_current_sheet_index();
            let sheet_name = self.workbook.get_current_sheet_name();

            let old_cell = self.workbook.get_current_sheet().data[row][col].clone();

            let mut new_cell = old_cell.clone();
            new_cell.value = content.clone();

            let cell_action = CellAction::new(
                sheet_index,
                sheet_name,
                row,
                col,
                old_cell,
                new_cell,
                ActionType::Edit,
            );

            self.undo_history.push(ActionCommand::Cell(cell_action));

            self.workbook.set_cell_value(row, col, content)?;
            self.input_mode = InputMode::Normal;
            self.input_buffer = String::new();
            self.text_area = ratatui_textarea::TextArea::default();
        }
        Ok(())
    }

    pub fn copy_cell(&mut self) {
        let content = self.get_cell_content_mut(self.selected_cell.0, self.selected_cell.1);
        self.clipboard = Some(content);
        self.add_notification("Cell content copied".to_string());
    }

    pub fn cut_cell(&mut self) -> Result<()> {
        let (row, col) = self.selected_cell;

        self.workbook.ensure_cell_exists(row, col);

        self.ensure_column_widths();

        let content = self.get_cell_content(row, col);
        self.clipboard = Some(content);

        let sheet_index = self.workbook.get_current_sheet_index();
        let sheet_name = self.workbook.get_current_sheet_name();

        let old_cell = self.workbook.get_current_sheet().data[row][col].clone();

        let mut new_cell = old_cell.clone();
        new_cell.value = String::new();

        let cell_action = CellAction::new(
            sheet_index,
            sheet_name,
            row,
            col,
            old_cell,
            new_cell,
            ActionType::Cut,
        );

        self.undo_history.push(ActionCommand::Cell(cell_action));
        self.workbook.set_cell_value(row, col, String::new())?;

        self.add_notification("Cell content cut".to_string());
        Ok(())
    }

    pub fn paste_cell(&mut self) -> Result<()> {
        if let Some(content) = self.clipboard.clone() {
            let (row, col) = self.selected_cell;

            self.workbook.ensure_cell_exists(row, col);
            self.ensure_column_widths();

            let sheet_index = self.workbook.get_current_sheet_index();
            let sheet_name = self.workbook.get_current_sheet_name();

            let old_cell = self.workbook.get_current_sheet().data[row][col].clone();

            let mut new_cell = old_cell.clone();
            new_cell.value = content.clone();

            let cell_action = CellAction::new(
                sheet_index,
                sheet_name,
                row,
                col,
                old_cell,
                new_cell,
                ActionType::Paste,
            );

            self.undo_history.push(ActionCommand::Cell(cell_action));
            self.workbook.set_cell_value(row, col, content)?;
            self.add_notification("Content pasted".to_string());
        } else {
            self.add_notification("Clipboard is empty".to_string());
        }
        Ok(())
    }
}
