use crate::app::AppState;
use crate::app::InputMode;

impl AppState<'_> {
    pub fn start_search_forward(&mut self) {
        self.input_mode = InputMode::SearchForward;
        self.input_buffer = String::new();
        self.text_area = tui_textarea::TextArea::default();
        self.add_notification("Search forward mode".to_string());
        self.highlight_enabled = true;
    }

    pub fn start_search_backward(&mut self) {
        self.input_mode = InputMode::SearchBackward;
        self.input_buffer = String::new();
        self.text_area = tui_textarea::TextArea::default();
        self.add_notification("Search backward mode".to_string());
        self.highlight_enabled = true;
    }

    pub fn execute_search(&mut self) {
        let query = self.text_area.lines().join("\n");
        self.input_buffer = query.clone();

        if query.is_empty() {
            self.input_mode = InputMode::Normal;
            return;
        }

        // Save the query for n/N commands
        self.search_query = query.clone();

        // Set search direction based on mode
        match self.input_mode {
            InputMode::SearchForward => self.search_direction = true,
            InputMode::SearchBackward => self.search_direction = false,
            _ => {}
        }

        self.search_results = self.find_all_matches(&query);

        if self.search_results.is_empty() {
            self.add_notification(format!("Pattern not found: {}", query));
            self.current_search_idx = None;
        } else {
            // Find the appropriate result to jump to based on search direction and current position
            self.jump_to_next_search_result();
            self.add_notification(format!(
                "{} matches found for: {}",
                self.search_results.len(),
                query
            ));
        }

        self.input_mode = InputMode::Normal;
        self.input_buffer = String::new();
        self.text_area = tui_textarea::TextArea::default();
    }

    pub fn find_all_matches(&self, query: &str) -> Vec<(usize, usize)> {
        let sheet = self.workbook.get_current_sheet();
        let query_lower = query.to_lowercase();

        // Pre-allocate with reasonable capacity
        let mut results = Vec::with_capacity(32);

        // row-first, column-second order
        for row in 1..=sheet.max_rows {
            for col in 1..=sheet.max_cols {
                if row < sheet.data.len() && col < sheet.data[row].len() {
                    let cell_content = &sheet.data[row][col].value;

                    if cell_content.is_empty() {
                        continue;
                    }

                    if self.case_insensitive_contains(cell_content, &query_lower) {
                        results.push((row, col));
                    }
                }
            }
        }

        results
    }

    fn case_insensitive_contains(&self, haystack: &str, needle: &str) -> bool {
        if needle.is_empty() {
            return true;
        }
        if haystack.is_empty() {
            return false;
        }

        haystack.to_lowercase().contains(needle)
    }

    pub fn jump_to_next_search_result(&mut self) {
        if self.search_results.is_empty() {
            return;
        }

        self.highlight_enabled = true;

        let current_pos = self.selected_cell;

        if self.search_direction {
            // Forward search
            let next_idx = self.search_results.iter().position(|&pos| {
                pos.0 > current_pos.0 || (pos.0 == current_pos.0 && pos.1 > current_pos.1)
            });

            match next_idx {
                Some(idx) => {
                    self.current_search_idx = Some(idx);
                    self.selected_cell = self.search_results[idx];
                }
                None => {
                    // Wrap around to the first result
                    self.current_search_idx = Some(0);
                    self.selected_cell = self.search_results[0];
                    self.add_notification("Search wrapped to top".to_string());
                }
            }
        } else {
            // Backward search
            let prev_idx = self.search_results.iter().rposition(|&pos| {
                pos.0 < current_pos.0 || (pos.0 == current_pos.0 && pos.1 < current_pos.1)
            });

            match prev_idx {
                Some(idx) => {
                    self.current_search_idx = Some(idx);
                    self.selected_cell = self.search_results[idx];
                }
                None => {
                    // Wrap around to the last result
                    let last_idx = self.search_results.len() - 1;
                    self.current_search_idx = Some(last_idx);
                    self.selected_cell = self.search_results[last_idx];
                    self.add_notification("Search wrapped to bottom".to_string());
                }
            }
        }

        self.handle_scrolling();
    }

    pub fn jump_to_prev_search_result(&mut self) {
        if self.search_results.is_empty() {
            return;
        }

        // Temporarily flip the search direction
        self.search_direction = !self.search_direction;
        self.jump_to_next_search_result();
        // Restore original search direction
        self.search_direction = !self.search_direction;
    }

    pub fn disable_search_highlight(&mut self) {
        self.highlight_enabled = false;
        self.add_notification("Search highlighting disabled".to_string());
    }
}
