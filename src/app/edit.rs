use crate::actions::{ActionCommand, ActionType, CellAction};
use crate::app::AppState;
use crate::app::InputMode;
use crate::app::{Transition, VimMode, VimState};
use anyhow::Result;
use ratatui::style::{Modifier, Style};
use tui_textarea::Input;

impl AppState<'_> {
    pub fn start_editing(&mut self) {
        self.input_mode = InputMode::Editing;
        let content = self.get_cell_content(self.selected_cell.0, self.selected_cell.1);
        self.input_buffer.clone_from(&content);

        // Initialize TextArea with content and settings
        let mut text_area = tui_textarea::TextArea::default();
        text_area.insert_str(&content);
        text_area.set_tab_length(4);
        text_area.set_cursor_line_style(Style::default());
        text_area.set_cursor_style(Style::default().add_modifier(Modifier::REVERSED));

        self.text_area = text_area;
        self.vim_state = Some(VimState::new(VimMode::Normal));
    }

    pub fn handle_vim_input(&mut self, input: Input) -> Result<()> {
        if let Some(vim_state) = &mut self.vim_state {
            match vim_state.transition(input, &mut self.text_area) {
                Transition::Mode(mode) => {
                    self.vim_state = Some(VimState::new(mode));
                }
                Transition::Pending(pending) => {
                    self.vim_state = Some(vim_state.clone().with_pending(pending));
                }
                Transition::Exit => {
                    // Confirm edit and exit Vim mode
                    self.confirm_edit()?;
                }
                Transition::Nop => {}
            }
        }
        Ok(())
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
            new_cell.value.clone_from(&content);

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
            self.text_area = tui_textarea::TextArea::default();
            self.vim_state = None;
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
            new_cell.value.clone_from(&content);

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
