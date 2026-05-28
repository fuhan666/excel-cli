use crossterm::event::{KeyCode, KeyEvent, KeyModifiers};
use tui_textarea::{Input, Key, TextArea};

use crate::app::{help_reference_line_count, AppState, InputMode};

pub fn handle_key_event(app_state: &mut AppState, key: KeyEvent) {
    match app_state.input_mode {
        InputMode::Normal => {
            if key.modifiers.contains(KeyModifiers::CONTROL)
                || key.modifiers.contains(KeyModifiers::SUPER)
            {
                handle_ctrl_key(app_state, key.code);
            } else {
                handle_normal_mode(app_state, key.code);
            }
        }
        InputMode::Editing => handle_editing_mode(app_state, key),
        InputMode::Command => handle_command_mode(app_state, key.code),
        InputMode::CommandInLazyLoading => handle_command_in_lazy_loading_mode(app_state, key.code),
        InputMode::SearchForward => handle_search_mode(app_state, key.code),
        InputMode::SearchBackward => handle_search_mode(app_state, key.code),
        InputMode::Help => handle_help_mode(app_state, key.code),
        InputMode::LazyLoading => handle_lazy_loading_mode(app_state, key.code),
    }
}

// Handles both Ctrl+key and Command+key (on Mac) combinations
fn handle_ctrl_key(app_state: &mut AppState, key_code: KeyCode) {
    match key_code {
        KeyCode::Left => {
            app_state.jump_to_prev_non_empty_cell_left();
        }
        KeyCode::Right => {
            app_state.jump_to_prev_non_empty_cell_right();
        }
        KeyCode::Up => {
            app_state.jump_to_prev_non_empty_cell_up();
        }
        KeyCode::Down => {
            app_state.jump_to_prev_non_empty_cell_down();
        }
        KeyCode::Char('r') => {
            if let Err(e) = app_state.redo() {
                app_state.add_notification(format!("Redo failed: {e}"));
            }
        }
        _ => {}
    }
}

fn handle_command_mode(app_state: &mut AppState, key_code: KeyCode) {
    match key_code {
        KeyCode::Enter => app_state.execute_command(),
        KeyCode::Esc => app_state.cancel_input(),
        KeyCode::Backspace => app_state.delete_char_from_input(),
        KeyCode::Char(c) => app_state.add_char_to_input(c),
        _ => {}
    }
}

fn handle_command_in_lazy_loading_mode(app_state: &mut AppState, key_code: KeyCode) {
    match key_code {
        KeyCode::Enter => {
            // Execute the command but stay in lazy loading mode if needed
            let current_index = app_state.workbook.get_current_sheet_index();
            let is_sheet_loaded = app_state.workbook.is_sheet_loaded(current_index);

            // Execute the command
            app_state.execute_command();

            // If the sheet is still not loaded after command execution, switch back to LazyLoading mode
            if !is_sheet_loaded
                && !app_state
                    .workbook
                    .is_sheet_loaded(app_state.workbook.get_current_sheet_index())
                && matches!(app_state.input_mode, InputMode::Normal)
            {
                app_state.input_mode = InputMode::LazyLoading;
            }
        }
        KeyCode::Esc => {
            // Return to LazyLoading mode
            app_state.input_mode = InputMode::LazyLoading;
            app_state.input_buffer = String::new();
        }
        KeyCode::Backspace => app_state.delete_char_from_input(),
        KeyCode::Char(c) => app_state.add_char_to_input(c),
        _ => {}
    }
}

fn handle_normal_mode(app_state: &mut AppState, key_code: KeyCode) {
    match key_code {
        KeyCode::Enter => {
            app_state.g_pressed = false;

            // Check if the current sheet is loaded
            let index = app_state.workbook.get_current_sheet_index();
            let sheet_name = app_state.workbook.get_current_sheet_name();

            if app_state.workbook.is_lazy_loading() && !app_state.workbook.is_sheet_loaded(index) {
                // If the sheet is not loaded, load it first
                if let Err(e) = app_state.workbook.ensure_sheet_loaded(index, &sheet_name) {
                    app_state.add_notification(format!("Failed to load sheet: {e}"));
                } else {
                    app_state.start_editing();
                }
            } else {
                app_state.start_editing();
            }
        }
        KeyCode::Char('h') => {
            app_state.g_pressed = false;
            app_state.move_cursor(0, -1);
        }
        KeyCode::Char('j') => {
            app_state.g_pressed = false;
            app_state.move_cursor(1, 0);
        }
        KeyCode::Char('k') => {
            app_state.g_pressed = false;
            app_state.move_cursor(-1, 0);
        }
        KeyCode::Char('l') => {
            app_state.g_pressed = false;
            app_state.move_cursor(0, 1);
        }
        KeyCode::Char('u') => {
            app_state.g_pressed = false;
            if let Err(e) = app_state.undo() {
                app_state.add_notification(format!("Undo failed: {e}"));
            }
        }
        KeyCode::Char('=' | '+') => {
            app_state.g_pressed = false;
            app_state.adjust_info_panel_height(1);
        }
        KeyCode::Char('-') => {
            app_state.g_pressed = false;
            app_state.adjust_info_panel_height(-1);
        }
        KeyCode::Char('[') => {
            app_state.g_pressed = false;
            if let Err(e) = app_state.prev_sheet() {
                app_state.add_notification(format!("Failed to switch to previous sheet: {e}"));
            }
        }
        KeyCode::Char(']') => {
            app_state.g_pressed = false;
            if let Err(e) = app_state.next_sheet() {
                app_state.add_notification(format!("Failed to switch to next sheet: {e}"));
            }
        }
        KeyCode::Char('g') => {
            if app_state.g_pressed {
                app_state.jump_to_first_row();
                app_state.g_pressed = false;
            } else {
                app_state.g_pressed = true;
            }
        }
        KeyCode::Char('G') => {
            app_state.g_pressed = false;
            app_state.jump_to_last_row();
        }
        KeyCode::Char('0') => {
            app_state.g_pressed = false;
            app_state.jump_to_first_column();
        }
        KeyCode::Char('^') => {
            app_state.g_pressed = false;
            app_state.jump_to_first_non_empty_column();
        }
        KeyCode::Char('$') => {
            app_state.g_pressed = false;
            app_state.jump_to_last_column();
        }
        KeyCode::Char('y') => {
            app_state.g_pressed = false;
            app_state.copy_cell();
        }
        KeyCode::Char('d') => {
            app_state.g_pressed = false;
            if let Err(e) = app_state.cut_cell() {
                app_state.add_notification(format!("Cut failed: {e}"));
            }
        }
        KeyCode::Char('p') => {
            app_state.g_pressed = false;
            if let Err(e) = app_state.paste_cell() {
                app_state.add_notification(format!("Paste failed: {e}"));
            }
        }
        KeyCode::Char(':') => {
            app_state.g_pressed = false;
            app_state.start_command_mode();
        }
        KeyCode::Char('/') => {
            app_state.g_pressed = false;
            app_state.start_search_forward();
        }
        KeyCode::Char('?') => {
            app_state.g_pressed = false;
            app_state.start_search_backward();
        }
        KeyCode::Char('n') => {
            app_state.g_pressed = false;
            if !app_state.search_results.is_empty() {
                app_state.jump_to_next_search_result();
            } else if !app_state.search_query.is_empty() {
                // Re-run the last search if we have a query but no results
                app_state.search_results = app_state.find_all_matches(&app_state.search_query);
                if !app_state.search_results.is_empty() {
                    app_state.jump_to_next_search_result();
                }
            }
        }

        KeyCode::Char('N') => {
            app_state.g_pressed = false;
            if !app_state.search_results.is_empty() {
                app_state.jump_to_prev_search_result();
            } else if !app_state.search_query.is_empty() {
                // Re-run the last search if we have a query but no results
                app_state.search_results = app_state.find_all_matches(&app_state.search_query);
                if !app_state.search_results.is_empty() {
                    app_state.jump_to_prev_search_result();
                }
            }
        }

        KeyCode::Left => {
            app_state.g_pressed = false;
            app_state.move_cursor(0, -1);
        }
        KeyCode::Right => {
            app_state.g_pressed = false;
            app_state.move_cursor(0, 1);
        }
        KeyCode::Up => {
            app_state.g_pressed = false;
            app_state.move_cursor(-1, 0);
        }
        KeyCode::Down => {
            app_state.g_pressed = false;
            app_state.move_cursor(1, 0);
        }
        _ => {
            app_state.g_pressed = false;
        }
    }
}

fn handle_editing_mode(app_state: &mut AppState, key: KeyEvent) {
    // Convert KeyEvent to Input for tui-textarea
    let input = Input {
        key: key_code_to_tui_key(key.code),
        ctrl: key.modifiers.contains(KeyModifiers::CONTROL),
        alt: key.modifiers.contains(KeyModifiers::ALT),
        shift: key.modifiers.contains(KeyModifiers::SHIFT),
    };

    if let Err(e) = app_state.handle_vim_input(input) {
        app_state.add_notification(format!("Vim input error: {e}"));
    }
}

fn handle_search_mode(app_state: &mut AppState, key_code: KeyCode) {
    match key_code {
        KeyCode::Enter => app_state.execute_search(),
        KeyCode::Esc => {
            app_state.input_mode = InputMode::Normal;
            app_state.input_buffer = String::new();
            app_state.text_area = TextArea::default();
        }
        _ => {
            let input = Input {
                key: key_code_to_tui_key(key_code),
                ctrl: false,
                alt: false,
                shift: false,
            };
            app_state.text_area.input(input);
        }
    }
}

// Convert crossterm::event::KeyCode to tui_textarea::Key
fn key_code_to_tui_key(key_code: KeyCode) -> Key {
    match key_code {
        KeyCode::Backspace => Key::Backspace,
        KeyCode::Enter => Key::Enter,
        KeyCode::Left => Key::Left,
        KeyCode::Right => Key::Right,
        KeyCode::Up => Key::Up,
        KeyCode::Down => Key::Down,
        KeyCode::Home => Key::Home,
        KeyCode::End => Key::End,
        KeyCode::PageUp => Key::PageUp,
        KeyCode::PageDown => Key::PageDown,
        KeyCode::Tab => Key::Tab,
        KeyCode::Delete => Key::Delete,
        // BackTab and Insert not supported in tui-textarea
        KeyCode::BackTab | KeyCode::Insert => Key::Null,
        KeyCode::Esc => Key::Esc,
        KeyCode::Char(c) => Key::Char(c),
        KeyCode::F(n) => Key::F(n),
        _ => Key::Null,
    }
}

fn handle_lazy_loading_mode(app_state: &mut AppState, key_code: KeyCode) {
    match key_code {
        KeyCode::Enter => {
            let index = app_state.workbook.get_current_sheet_index();
            let sheet_name = app_state.workbook.get_current_sheet_name();

            // Load the sheet
            if let Err(e) = app_state.workbook.ensure_sheet_loaded(index, &sheet_name) {
                app_state.add_notification(format!("Failed to load sheet: {e}"));
            }

            app_state.input_mode = InputMode::Normal;
        }
        KeyCode::Char('[') => {
            // Switch to previous sheet
            let current_index = app_state.workbook.get_current_sheet_index();

            if current_index == 0 {
                app_state.add_notification("Already at the first sheet".to_string());
            } else {
                // The method will automatically set the input mode to LazyLoading if the sheet is not loaded
                if let Err(e) = app_state.switch_sheet_by_index(current_index - 1) {
                    app_state.add_notification(format!("Failed to switch to previous sheet: {e}"));
                }
            }
        }
        KeyCode::Char(']') => {
            // Switch to next sheet
            let current_index = app_state.workbook.get_current_sheet_index();
            let sheet_count = app_state.workbook.get_sheet_names().len();

            if current_index >= sheet_count - 1 {
                app_state.add_notification("Already at the last sheet".to_string());
            } else {
                // The method will automatically set the input mode to LazyLoading if the sheet is not loaded
                if let Err(e) = app_state.switch_sheet_by_index(current_index + 1) {
                    app_state.add_notification(format!("Failed to switch to next sheet: {e}"));
                }
            }
        }
        KeyCode::Char(':') => {
            // Allow entering command mode from lazy loading mode
            app_state.start_command_in_lazy_loading_mode();
        }
        _ => {
            app_state.add_notification(
                "Press Enter to load the sheet data, or use [ and ] to switch sheets".to_string(),
            );
        }
    }
}

fn handle_help_mode(app_state: &mut AppState, key_code: KeyCode) {
    let line_count = app_state.help_total_lines.max(help_reference_line_count());
    let visible_lines = app_state.help_visible_lines.max(1);
    let max_scroll = line_count.saturating_sub(visible_lines);

    match key_code {
        KeyCode::Enter | KeyCode::Esc | KeyCode::Char('q') => {
            app_state.input_mode = InputMode::Normal;
        }
        KeyCode::Char('j') | KeyCode::Down => {
            app_state.help_scroll = (app_state.help_scroll + 1).min(max_scroll);
        }
        KeyCode::Char('k') | KeyCode::Up => {
            app_state.help_scroll = app_state.help_scroll.saturating_sub(1);
        }
        KeyCode::PageDown => {
            app_state.help_scroll = (app_state.help_scroll + visible_lines).min(max_scroll);
        }
        KeyCode::PageUp => {
            app_state.help_scroll = app_state.help_scroll.saturating_sub(visible_lines);
        }
        KeyCode::Home => {
            app_state.help_scroll = 0;
        }
        KeyCode::End => {
            app_state.help_scroll = max_scroll;
        }
        _ => {}
    }
}

#[cfg(test)]
mod tests {
    use crossterm::event::{KeyCode, KeyEvent, KeyModifiers};
    use std::path::PathBuf;

    use super::handle_key_event;
    use crate::app::{AppState, InputMode};
    use crate::excel::{Cell, FreezePanes, Sheet, Workbook, EXCEL_MAX_COLS, EXCEL_MAX_ROWS};
    use crate::utils::index_to_col_name;

    fn app_with_sheet() -> AppState<'static> {
        let mut data = vec![vec![Cell::empty(); 3]; 3];
        data[1][1] = Cell::new("Name".to_string(), false);
        data[1][2] = Cell::new("Name".to_string(), false);
        data[2][1] = Cell::new("Ada".to_string(), false);
        data[2][2] = Cell::new("10".to_string(), false);
        let sheet = Sheet {
            name: "Data".to_string(),
            data,
            max_rows: 2,
            max_cols: 2,
            is_loaded: true,
            freeze_panes: FreezePanes::none(),
        };
        let app = AppState::new(
            Workbook::from_sheets_for_test(vec![sheet]),
            PathBuf::from("test.xlsx"),
        )
        .unwrap();
        app
    }

    #[test]
    fn question_mark_starts_backward_search_from_normal_mode() {
        let mut app = app_with_sheet();

        handle_key_event(
            &mut app,
            KeyEvent::new(KeyCode::Char('?'), KeyModifiers::empty()),
        );

        assert!(matches!(app.input_mode, InputMode::SearchBackward));
    }

    #[test]
    fn chinese_question_mark_does_not_open_help_overlay_from_normal_mode() {
        let mut app = app_with_sheet();

        handle_key_event(
            &mut app,
            KeyEvent::new(KeyCode::Char('？'), KeyModifiers::empty()),
        );

        assert!(matches!(app.input_mode, InputMode::Normal));
    }

    #[test]
    fn right_movement_can_enter_blank_columns_beyond_used_range() {
        let mut app = app_with_sheet();
        app.selected_cell = (2, 2);

        handle_key_event(
            &mut app,
            KeyEvent::new(KeyCode::Char('l'), KeyModifiers::empty()),
        );

        assert_eq!(app.selected_cell, (2, 3));
        assert_eq!(app.get_cell_content(2, 3), "");
    }

    #[test]
    fn down_movement_can_enter_blank_rows_beyond_used_range() {
        let mut app = app_with_sheet();
        app.selected_cell = (2, 2);

        handle_key_event(
            &mut app,
            KeyEvent::new(KeyCode::Char('j'), KeyModifiers::empty()),
        );

        assert_eq!(app.selected_cell, (3, 2));
        assert_eq!(app.get_cell_content(3, 2), "");
    }

    #[test]
    fn down_movement_stops_at_excel_row_limit() {
        let mut app = app_with_sheet();
        app.selected_cell = (EXCEL_MAX_ROWS, 2);

        handle_key_event(
            &mut app,
            KeyEvent::new(KeyCode::Down, KeyModifiers::empty()),
        );

        assert_eq!(app.selected_cell, (EXCEL_MAX_ROWS, 2));
    }

    #[test]
    fn right_movement_stops_at_excel_column_limit() {
        let mut app = app_with_sheet();
        app.selected_cell = (2, EXCEL_MAX_COLS);

        handle_key_event(
            &mut app,
            KeyEvent::new(KeyCode::Char('l'), KeyModifiers::empty()),
        );

        assert_eq!(app.selected_cell, (2, EXCEL_MAX_COLS));
        assert_eq!(index_to_col_name(app.selected_cell.1), "XFD");
    }

    #[test]
    fn q_closes_help_overlay_without_quitting() {
        let mut app = app_with_sheet();
        app.show_help();

        handle_key_event(
            &mut app,
            KeyEvent::new(KeyCode::Char('q'), KeyModifiers::empty()),
        );

        assert!(matches!(app.input_mode, InputMode::Normal));
        assert!(!app.should_quit);
    }

    #[test]
    fn page_keys_scroll_help_overlay_by_visible_page() {
        let mut app = app_with_sheet();
        app.show_help();
        app.help_visible_lines = 8;

        handle_key_event(
            &mut app,
            KeyEvent::new(KeyCode::PageDown, KeyModifiers::empty()),
        );

        assert_eq!(app.help_scroll, 8);

        handle_key_event(
            &mut app,
            KeyEvent::new(KeyCode::PageUp, KeyModifiers::empty()),
        );

        assert_eq!(app.help_scroll, 0);
    }

    #[test]
    fn end_key_scrolls_help_overlay_to_full_reference_bottom() {
        let mut app = app_with_sheet();
        app.show_help();
        app.help_visible_lines = 8;

        handle_key_event(&mut app, KeyEvent::new(KeyCode::End, KeyModifiers::empty()));

        let full_reference_bottom = crate::app::help_reference_text()
            .lines()
            .count()
            .saturating_sub(app.help_visible_lines);
        assert_eq!(app.help_scroll, full_reference_bottom);
    }

    #[test]
    fn end_key_uses_rendered_help_total_lines_when_available() {
        let mut app = app_with_sheet();
        app.show_help();
        app.help_visible_lines = 8;
        app.help_total_lines = 120;

        handle_key_event(&mut app, KeyEvent::new(KeyCode::End, KeyModifiers::empty()));

        assert_eq!(app.help_scroll, 112);
    }
}
