use crate::app::AppState;
use crate::app::InputMode;

impl AppState<'_> {
    pub fn show_help(&mut self) {
        self.help_scroll = 0;
        self.help_text = crate::app::help_reference_text();
        self.help_total_lines = crate::app::help_reference_line_count();

        self.input_mode = InputMode::Help;
    }

    pub fn save_and_exit(&mut self) {
        if !self.workbook.is_modified() {
            self.add_notification("No changes to save".to_string());
            self.should_quit = true;
            return;
        }

        match self.workbook.save() {
            Ok(_) => {
                self.undo_history.clear();
                self.add_notification("File saved".to_string());
                self.should_quit = true;
            }
            Err(e) => {
                self.add_notification(format!("Save failed: {e}"));
                self.input_mode = InputMode::Normal;
            }
        }
    }

    pub fn save(&mut self) -> Result<(), anyhow::Error> {
        if !self.workbook.is_modified() {
            self.add_notification("No changes to save".to_string());
            return Ok(());
        }

        match self.workbook.save() {
            Ok(_) => {
                self.undo_history.clear();
                self.add_notification("File saved".to_string());
            }
            Err(e) => {
                self.add_notification(format!("Save failed: {e}"));
            }
        }
        Ok(())
    }

    pub fn exit_without_saving(&mut self) {
        self.should_quit = true;
    }
}
