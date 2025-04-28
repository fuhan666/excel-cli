use anyhow::Result;
use std::collections::HashMap;
use std::path::PathBuf;
use tui_textarea::TextArea;

use crate::actions::UndoHistory;
use crate::app::VimState;
use crate::excel::Workbook;

pub enum InputMode {
    Normal,
    Editing,
    Command,
    SearchForward,
    SearchBackward,
    Help,
}

pub struct AppState<'a> {
    pub workbook: Workbook,
    pub file_path: PathBuf,
    pub selected_cell: (usize, usize), // (row, col)
    pub start_row: usize,
    pub start_col: usize,
    pub visible_rows: usize,
    pub visible_cols: usize,
    pub input_mode: InputMode,
    pub input_buffer: String,
    pub text_area: TextArea<'a>,
    pub should_quit: bool,
    pub column_widths: Vec<usize>, // Store width for current sheet's columns
    pub sheet_column_widths: HashMap<String, Vec<usize>>, // Store column widths for each sheet
    pub clipboard: Option<String>, // Store copied/cut cell content
    pub g_pressed: bool,           // Track if 'g' was pressed for 'gg' command
    pub search_query: String,      // Current search query
    pub search_results: Vec<(usize, usize)>, // List of cells matching the search query
    pub current_search_idx: Option<usize>, // Index of current search result
    pub search_direction: bool,    // true for forward, false for backward
    pub highlight_enabled: bool,   // Control whether search results are highlighted
    pub info_panel_height: usize,
    pub notification_messages: Vec<String>,
    pub max_notifications: usize,
    pub help_text: String,
    pub help_scroll: usize,
    pub help_visible_lines: usize,
    pub undo_history: UndoHistory,
    pub vim_state: Option<VimState>,
}

impl AppState<'_> {
    pub fn new(workbook: Workbook, file_path: PathBuf) -> Result<Self> {
        // Initialize default column widths for current sheet
        let max_cols = workbook.get_current_sheet().max_cols;
        let default_width = 15;
        let column_widths = vec![default_width; max_cols + 1];

        // Initialize column widths for all sheets
        let mut sheet_column_widths = HashMap::with_capacity(workbook.get_sheet_names().len());
        let sheet_names = workbook.get_sheet_names();

        for (i, name) in sheet_names.iter().enumerate() {
            if i == workbook.get_current_sheet_index() {
                sheet_column_widths.insert(name.clone(), column_widths.clone());
            } else {
                let sheet_max_cols = if let Some(sheet) = workbook.get_sheet_by_index(i) {
                    sheet.max_cols
                } else {
                    max_cols // Fallback to current sheet's max_cols
                };
                sheet_column_widths.insert(name.clone(), vec![default_width; sheet_max_cols + 1]);
            }
        }

        // Initialize TextArea
        let text_area = TextArea::default();

        Ok(Self {
            workbook,
            file_path,
            selected_cell: (1, 1), // Excel uses 1-based indexing
            start_row: 1,
            start_col: 1,
            visible_rows: 30, // Default values, will be adjusted based on window size
            visible_cols: 15, // Default values, will be adjusted based on window size
            input_mode: InputMode::Normal,
            input_buffer: String::new(),
            text_area,
            should_quit: false,
            column_widths,
            sheet_column_widths,
            clipboard: None,
            g_pressed: false,
            search_query: String::new(),
            search_results: Vec::new(),
            current_search_idx: None,
            search_direction: true,  // Default to forward search
            highlight_enabled: true, // Default to showing highlights
            info_panel_height: 10,
            notification_messages: Vec::new(),
            max_notifications: 5,
            help_text: String::new(),
            help_scroll: 0,
            help_visible_lines: 20,
            undo_history: UndoHistory::new(),
            vim_state: None,
        })
    }

    pub fn add_notification(&mut self, message: String) {
        self.notification_messages.push(message);

        if self.notification_messages.len() > self.max_notifications {
            self.notification_messages.remove(0);
        }
    }

    pub fn adjust_info_panel_height(&mut self, delta: isize) {
        let new_height = (self.info_panel_height as isize + delta).clamp(6, 16) as usize;
        if new_height != self.info_panel_height {
            self.info_panel_height = new_height;
            self.add_notification(format!("Info panel height: {}", self.info_panel_height));
        }
    }

    pub fn get_cell_content(&self, row: usize, col: usize) -> String {
        let sheet = self.workbook.get_current_sheet();

        if row < sheet.data.len() && col < sheet.data[0].len() {
            let cell = &sheet.data[row][col];
            if cell.is_formula {
                let mut result = String::with_capacity(9 + cell.value.len());
                result.push_str("Formula: ");
                result.push_str(&cell.value);
                result
            } else {
                cell.value.clone()
            }
        } else {
            String::new()
        }
    }

    pub fn get_cell_content_mut(&mut self, row: usize, col: usize) -> String {
        self.workbook.ensure_cell_exists(row, col);

        self.ensure_column_widths();

        let sheet = self.workbook.get_current_sheet();
        let cell = &sheet.data[row][col];

        if cell.is_formula {
            let mut result = String::with_capacity(9 + cell.value.len());
            result.push_str("Formula: ");
            result.push_str(&cell.value);
            result
        } else {
            cell.value.clone()
        }
    }

    pub fn cancel_input(&mut self) {
        // If in help mode, just close the help window
        if let InputMode::Help = self.input_mode {
            self.input_mode = InputMode::Normal;
            return;
        }

        // Otherwise, cancel the current input
        self.input_mode = InputMode::Normal;
        self.input_buffer = String::new();
        self.text_area = TextArea::default();
    }

    pub fn add_char_to_input(&mut self, c: char) {
        self.input_buffer.push(c);
    }

    pub fn delete_char_from_input(&mut self) {
        self.input_buffer.pop();
    }

    pub fn start_command_mode(&mut self) {
        self.input_mode = InputMode::Command;
        self.input_buffer = String::new();
    }
}
