use anyhow::Result;
use crossterm::{
    event::{self, Event, KeyCode, KeyEventKind},
    terminal::{disable_raw_mode, enable_raw_mode, EnterAlternateScreen, LeaveAlternateScreen},
    ExecutableCommand,
};
use ratatui::{
    backend::CrosstermBackend,
    layout::{Constraint, Direction, Layout, Rect},
    style::{Color, Modifier, Style},
    widgets::{Block, Borders, Cell, Paragraph, Row, Table},
    Frame, Terminal,
};
use std::{io, time::Duration};

use crate::app::{index_to_col_name, AppState, InputMode};

pub fn run_app(mut app_state: AppState) -> Result<()> {
    // Setup terminal
    enable_raw_mode()?;
    let mut stdout = io::stdout();
    stdout.execute(EnterAlternateScreen)?;
    let backend = CrosstermBackend::new(stdout);
    let mut terminal = Terminal::new(backend)?;

    // Main loop
    while !app_state.should_quit {
        terminal.draw(|f| ui(f, &mut app_state))?;

        if event::poll(Duration::from_millis(50))? {
            if let Event::Key(key) = event::read()? {
                if key.kind == KeyEventKind::Press {
                    handle_key_event(&mut app_state, key.code);
                }
            }
        }
    }

    // Restore terminal
    disable_raw_mode()?;
    terminal.backend_mut().execute(LeaveAlternateScreen)?;
    terminal.show_cursor()?;

    Ok(())
}

fn update_visible_area(app_state: &mut AppState, area: Rect) {
    let visible_rows = (area.height as usize).saturating_sub(3);
    app_state.visible_rows = visible_rows;

    let total_width = area.width as usize;
    let available_width = total_width.saturating_sub(5 + 2); // 5 for row numbers, 2 for borders

    app_state.ensure_column_visible(app_state.selected_cell.1);

    let mut visible_cols = 0;
    let mut width_used = 0;

    for col_idx in app_state.start_col.. {
        let col_width = app_state.get_column_width(col_idx);

        if col_idx == app_state.start_col {
            width_used += col_width;
            visible_cols += 1;

            if width_used >= available_width {
                break;
            }
        } else {
            if width_used + col_width <= available_width {
                width_used += col_width;
                visible_cols += 1;
            }
            // Excel-like behavior: include a partially visible column if there's space
            else if width_used < available_width {
                visible_cols += 1;
                break;
            } else {
                break;
            }
        }
    }

    app_state.visible_cols = visible_cols.max(1);
}

fn ui(f: &mut Frame, app_state: &mut AppState) {
    let chunks = Layout::default()
        .direction(Direction::Vertical)
        .constraints([
            Constraint::Length(1),
            Constraint::Min(1),
            Constraint::Length(1),
        ])
        .split(f.size());

    let title = format!(
        " {} - Sheet: {} ",
        app_state.file_path.display(),
        app_state.workbook.get_current_sheet_name()
    );

    let title_widget =
        Paragraph::new(title).style(Style::default().bg(Color::Blue).fg(Color::White));

    f.render_widget(title_widget, chunks[0]);

    update_visible_area(app_state, chunks[1]);
    draw_spreadsheet(f, app_state, chunks[1]);
    draw_status_bar(f, app_state, chunks[2]);
}

fn draw_spreadsheet(f: &mut Frame, app_state: &AppState, area: Rect) {
    let visible_rows = (app_state.start_row..=(app_state.start_row + app_state.visible_rows - 1))
        .collect::<Vec<_>>();
    let visible_cols = (app_state.start_col..=(app_state.start_col + app_state.visible_cols - 1))
        .collect::<Vec<_>>();

    let mut col_constraints = vec![Constraint::Length(5)];
    col_constraints.extend(
        visible_cols
            .iter()
            .map(|col| Constraint::Length(app_state.get_column_width(*col) as u16)),
    );

    let header_cells = {
        let mut cells = vec![Cell::from("")];
        cells.extend(visible_cols.iter().map(|col| {
            let col_name = index_to_col_name(*col);
            Cell::from(col_name).style(Style::default().bg(Color::Blue).fg(Color::White))
        }));
        Row::new(cells).height(1)
    };

    let rows = visible_rows
        .iter()
        .map(|row| {
            let row_header = Cell::from(row.to_string())
                .style(Style::default().bg(Color::Blue).fg(Color::White));

            let row_cells = visible_cols
                .iter()
                .map(|col| {
                    let content = if app_state.selected_cell == (*row, *col)
                        && matches!(app_state.input_mode, InputMode::Editing)
                    {
                        let buf = &app_state.input_buffer;
                        let col_width = app_state.get_column_width(*col);

                        let display_width = buf
                            .chars()
                            .fold(0, |acc, c| acc + if c.is_ascii() { 1 } else { 2 });

                        if display_width > col_width.saturating_sub(2) {
                            let mut cumulative_width = 0;
                            let chars_to_show = buf
                                .chars()
                                .rev()
                                .take_while(|&c| {
                                    let char_width = if c.is_ascii() { 1 } else { 2 };
                                    if cumulative_width + char_width <= col_width.saturating_sub(2)
                                    {
                                        cumulative_width += char_width;
                                        true
                                    } else {
                                        false
                                    }
                                })
                                .collect::<Vec<_>>();

                            chars_to_show.into_iter().rev().collect()
                        } else {
                            buf.clone()
                        }
                    } else {
                        let content = app_state.get_cell_content(*row, *col);
                        let col_width = app_state.get_column_width(*col);

                        let display_width = content
                            .chars()
                            .fold(0, |acc, c| acc + if c.is_ascii() { 1 } else { 2 });

                        if display_width > col_width {
                            let mut result = String::new();
                            let mut current_width = 0;

                            for c in content.chars() {
                                let char_width = if c.is_ascii() { 1 } else { 2 };
                                if current_width + char_width + 1 <= col_width {
                                    result.push(c);
                                    current_width += char_width;
                                } else {
                                    break;
                                }
                            }

                            if !content.is_empty() && result.len() < content.len() {
                                result.push('…');
                            }

                            result
                        } else {
                            content
                        }
                    };

                    let style = if app_state.selected_cell == (*row, *col) {
                        Style::default().bg(Color::DarkGray).fg(Color::White)
                    } else {
                        Style::default()
                    };

                    Cell::from(content).style(style)
                })
                .collect::<Vec<_>>();

            let mut cells = vec![row_header];
            cells.extend(row_cells);

            Row::new(cells)
        })
        .collect::<Vec<_>>();

    let table = Table::new(std::iter::once(header_cells).chain(rows), col_constraints)
        .block(Block::default().borders(Borders::ALL))
        .highlight_style(Style::default().add_modifier(Modifier::BOLD))
        .highlight_symbol(">> ");

    f.render_widget(table, area);
}

fn draw_status_bar(f: &mut Frame, app_state: &AppState, area: Rect) {
    let status = match app_state.input_mode {
        InputMode::Normal => {
            if !app_state.status_message.is_empty() {
                app_state.status_message.clone()
            } else {
                format!(
                    " Cell: {} | hjkl=move(1) HJKL=move(5) e=edit g=goto :=command",
                    cell_reference(app_state.selected_cell)
                )
            }
        }
        InputMode::Editing => {
            format!(
                " Editing cell {}: {}",
                cell_reference(app_state.selected_cell),
                app_state.input_buffer
            )
        }
        InputMode::Goto => {
            format!(" Go to cell: {}", app_state.input_buffer)
        }
        InputMode::Confirm => app_state.status_message.clone(),

        InputMode::Command => {
            format!(":{}", app_state.input_buffer)
        }
    };

    let status_style = Style::default().bg(Color::Green).fg(Color::White);
    let status_widget = Paragraph::new(status).style(status_style);

    f.render_widget(status_widget, area);
}

fn handle_key_event(app_state: &mut AppState, key_code: KeyCode) {
    match app_state.input_mode {
        InputMode::Normal => handle_normal_mode(app_state, key_code),
        InputMode::Editing => handle_editing_mode(app_state, key_code),
        InputMode::Goto => handle_goto_mode(app_state, key_code),
        InputMode::Confirm => handle_confirm_mode(app_state, key_code),
        InputMode::Command => handle_command_mode(app_state, key_code),
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
        KeyCode::Char('h') => app_state.move_cursor(0, -1),
        KeyCode::Char('j') => app_state.move_cursor(1, 0),
        KeyCode::Char('k') => app_state.move_cursor(-1, 0),
        KeyCode::Char('l') => app_state.move_cursor(0, 1),
        KeyCode::Char('H') => app_state.move_cursor(0, -5),
        KeyCode::Char('J') => app_state.move_cursor(5, 0),
        KeyCode::Char('K') => app_state.move_cursor(-5, 0),
        KeyCode::Char('L') => app_state.move_cursor(0, 5),
        KeyCode::Char('e') => app_state.start_editing(),
        KeyCode::Char('g') => app_state.start_goto(),
        // Enter command mode
        KeyCode::Char(':') => app_state.start_command_mode(),

        KeyCode::Left => app_state.move_cursor(0, -1),
        KeyCode::Right => app_state.move_cursor(0, 1),
        KeyCode::Up => app_state.move_cursor(-1, 0),
        KeyCode::Down => app_state.move_cursor(1, 0),
        _ => {}
    }
}

fn handle_editing_mode(app_state: &mut AppState, key_code: KeyCode) {
    match key_code {
        KeyCode::Enter => {
            if let Err(e) = app_state.confirm_edit() {
                app_state.status_message = format!("Error: {}", e);
            }
        }
        KeyCode::Esc => app_state.cancel_input(),
        KeyCode::Backspace => app_state.delete_char_from_input(),
        KeyCode::Char(c) => app_state.add_char_to_input(c),
        _ => {}
    }
}

fn handle_goto_mode(app_state: &mut AppState, key_code: KeyCode) {
    match key_code {
        KeyCode::Enter => app_state.confirm_goto(),
        KeyCode::Esc => app_state.cancel_input(),
        KeyCode::Backspace => app_state.delete_char_from_input(),
        KeyCode::Char(c) => app_state.add_char_to_input(c),
        _ => {}
    }
}

fn handle_confirm_mode(app_state: &mut AppState, key_code: KeyCode) {
    match key_code {
        KeyCode::Char('y') | KeyCode::Char('Y') => app_state.save_and_exit(),
        KeyCode::Char('n') | KeyCode::Char('N') => app_state.exit_without_saving(),
        KeyCode::Char('c') | KeyCode::Char('C') | KeyCode::Esc => app_state.cancel_exit(),
        _ => {}
    }
}

fn cell_reference(cell: (usize, usize)) -> String {
    format!("{}{}", index_to_col_name(cell.1), cell.0)
}
