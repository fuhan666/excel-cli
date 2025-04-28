use crossterm::event::{KeyCode, KeyEvent, KeyModifiers};
use tui_textarea::{TextArea, Input, Key};

use crate::app::{AppState, InputMode};

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
        InputMode::Editing => handle_editing_mode(app_state, key.code),
        InputMode::Command => handle_command_mode(app_state, key.code),
        InputMode::SearchForward => handle_search_mode(app_state, key.code),
        InputMode::SearchBackward => handle_search_mode(app_state, key.code),
        InputMode::Help => handle_help_mode(app_state, key.code),
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
                app_state.add_notification(format!("Redo failed: {}", e));
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

fn handle_normal_mode(app_state: &mut AppState, key_code: KeyCode) {
    match key_code {
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
                app_state.add_notification(format!("Undo failed: {}", e));
            }
        }
        KeyCode::Char('=') | KeyCode::Char('+') => {
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
                app_state.add_notification(format!("Failed to switch to previous sheet: {}", e));
            }
        }
        KeyCode::Char(']') => {
            app_state.g_pressed = false;
            if let Err(e) = app_state.next_sheet() {
                app_state.add_notification(format!("Failed to switch to next sheet: {}", e));
            }
        }
        KeyCode::Char('i') => {
            app_state.g_pressed = false;
            app_state.start_editing();
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
                app_state.add_notification(format!("Cut failed: {}", e));
            }
        }
        KeyCode::Char('p') => {
            app_state.g_pressed = false;
            if let Err(e) = app_state.paste_cell() {
                app_state.add_notification(format!("Paste failed: {}", e));
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

fn handle_editing_mode(app_state: &mut AppState, key_code: KeyCode) {
    match key_code {
        KeyCode::Enter => {
            if let Err(e) = app_state.confirm_edit() {
                app_state.add_notification(format!("Error: {}", e));
            }
        }
        KeyCode::Esc => app_state.cancel_input(),
        _ => {
            let input = Input {
                key: key_code_to_tui_key(key_code),
                ctrl: false,
                alt: false,
                shift: false,
            };
            app_state.text_area.input(input);

            // Update input_buffer with the current TextArea content to sync with cell display
            app_state.input_buffer = app_state.text_area.lines().join("\n");
        }
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
        KeyCode::BackTab => Key::Null, // BackTab not supported in tui-textarea
        KeyCode::Delete => Key::Delete,
        KeyCode::Insert => Key::Null, // Insert not supported in tui-textarea
        KeyCode::Esc => Key::Esc,
        KeyCode::Char(c) => Key::Char(c),
        KeyCode::F(n) => Key::F(n),
        KeyCode::Null => Key::Null,
        _ => Key::Null,
    }
}

fn handle_help_mode(app_state: &mut AppState, key_code: KeyCode) {
    let line_count = app_state.help_text.lines().count();

    let visible_lines = app_state.help_visible_lines;

    let max_scroll = line_count.saturating_sub(visible_lines).max(0);

    match key_code {
        KeyCode::Enter | KeyCode::Esc => {
            app_state.input_mode = InputMode::Normal;
        }
        KeyCode::Char('j') | KeyCode::Down => {
            // Scroll down, but not beyond the last line
            app_state.help_scroll = (app_state.help_scroll + 1).min(max_scroll);
        }
        KeyCode::Char('k') | KeyCode::Up => {
            // Scroll up
            app_state.help_scroll = app_state.help_scroll.saturating_sub(1);
        }
        KeyCode::Home => {
            // Scroll to the top
            app_state.help_scroll = 0;
        }
        KeyCode::End => {
            // Scroll to the bottom
            app_state.help_scroll = max_scroll;
        }
        _ => {}
    }
}
