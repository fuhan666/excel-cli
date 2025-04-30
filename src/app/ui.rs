use crate::app::AppState;
use crate::app::InputMode;

impl AppState<'_> {
    pub fn show_help(&mut self) {
        self.help_scroll = 0;

        self.help_text = "FILE OPERATIONS:\n\
             :w          - Save file\n\
             :wq, :x     - Save and quit\n\
             :q          - Quit (will warn if unsaved changes)\n\
             :q!         - Force quit without saving\n\n\
             NAVIGATION:\n\
             :[cell]     - Jump to cell (e.g., :B10)\n\
             hjkl        - Move cursor (left, down, up, right)\n\
             0           - Jump to first column\n\
             ^           - Jump to first non-empty column\n\
             $           - Jump to last column\n\
             gg          - Jump to first row\n\
             G           - Jump to last row\n\
             Ctrl+arrows - Jump to next non-empty cell\n\
             [           - Switch to previous sheet\n\
             ]           - Switch to next sheet\n\
             :sheet [name/number] - Switch to sheet by name or index\n\n\
             EDITING:\n\
             Enter       - Edit current cell\n\
             :y          - Copy current cell\n\
             :d          - Cut current cell\n\
             :put, :pu   - Paste to current cell\n\
             u           - Undo last operation\n\
             Ctrl+r      - Redo last undone operation\n\n\
             SEARCH:\n\
             /           - Search forward\n\
             ?           - Search backward\n\
             n           - Jump to next search result\n\
             N           - Jump to previous search result\n\
             :nohlsearch, :noh - Disable search highlighting\n\n\
             COLUMN OPERATIONS:\n\
             :cw fit     - Adjust width of current column to fit its content\n\
             :cw fit all - Adjust width of all columns to fit their content\n\
             :cw min     - Set current column width to minimum (5 characters)\n\
             :cw min all - Set all columns width to minimum\n\
             :cw [number] - Set current column width to specific number of characters\n\
             :dc         - Delete current column\n\
             :dc [col]   - Delete specific column (e.g., :dc A or :dc 1)\n\
             :dc [start] [end] - Delete columns from start to end (e.g., :dc A C)\n\n\
             ROW OPERATIONS:\n\
             :dr         - Delete current row\n\
             :dr [row]   - Delete specific row\n\
             :dr [start] [end] - Delete rows from start to end\n\n\
             EXPORT:\n\
             :ej [h|v] [rows]  - Export current sheet to JSON\n\
             :eja [h|v] [rows] - Export all sheets to a single JSON file\n\
                                h=horizontal (default), v=vertical\n\
                                [rows]=number of header rows (default: 1)\n\n\
             SHEET OPERATIONS:\n\
             :delsheet   - Delete the current sheet\n\n\
             UI ADJUSTMENTS:\n\
             +/=         - Increase info panel height\n\
             -           - Decrease info panel height\n\n\
             EDITING MODE:\n\
             Esc         - Exit Vim mode and save changes\n\
             i           - Enter Insert mode\n\
             v           - Enter Visual mode\n\
             y           - Yank (copy) text in Visual mode or with operator\n\
             d           - Delete text in Visual mode or with operator\n\
             c           - Change text in Visual mode or with operator\n\
             p           - Paste yanked or deleted text\n\
             u           - Undo last change\n\
             Ctrl+r      - Redo last undone change\n\
             h,j,k,l     - Move cursor left, down, up, right\n\
             w           - Move to next word\n\
             b           - Move to beginning of word\n\
             e           - Move to end of word\n\
             $           - Move to end of line\n\
             ^           - Move to first non-blank character of line\n\
             gg          - Move to first line\n\
             G           - Move to last line\n\
             x           - Delete character under cursor\n\
             D           - Delete to end of line\n\
             C           - Change to end of line\n\
             o           - Open new line below and enter Insert mode\n\
             O           - Open new line above and enter Insert mode\n\
             A           - Append at end of line\n\
             I           - Insert at beginning of line"
            .to_string();

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
