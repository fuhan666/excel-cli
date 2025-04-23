use anyhow::Result;
use crossterm::{
    ExecutableCommand,
    event::{self, Event, KeyCode, KeyEvent, KeyEventKind, KeyModifiers},
    terminal::{EnterAlternateScreen, LeaveAlternateScreen, disable_raw_mode, enable_raw_mode},
};
use ratatui::{
    Frame, Terminal,
    backend::CrosstermBackend,
    layout::{Constraint, Direction, Layout, Rect},
    style::{Color, Modifier, Style},
    text::{Line, Span},
    widgets::{Block, Borders, Cell, Clear, Paragraph, Row, Table},
};
use ratatui_textarea::TextArea;
use std::{io, time::Duration};

use crate::app::{AppState, InputMode, index_to_col_name};

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
                    handle_key_event(&mut app_state, key);
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
    // Create the main layout
    let chunks = Layout::default()
        .direction(Direction::Vertical)
        .constraints([
            Constraint::Length(1), // Combined title bar and sheet tabs
            Constraint::Min(1),    // Spreadsheet
            Constraint::Length(app_state.info_panel_height as u16), // Info panel
            Constraint::Length(1), // Status bar
        ])
        .split(f.size());

    draw_title_with_tabs(f, app_state, chunks[0]);

    update_visible_area(app_state, chunks[1]);
    draw_spreadsheet(f, app_state, chunks[1]);

    draw_info_panel(f, app_state, chunks[2]);
    draw_status_bar(f, app_state, chunks[3]);

    // If in help mode, draw the help popup over everything else
    if let InputMode::Help = app_state.input_mode {
        draw_help_popup(f, app_state, f.size());
    }
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

                    // Determine cell style based on selection and search results
                    let style = if app_state.selected_cell == (*row, *col) {
                        Style::default().bg(Color::DarkGray).fg(Color::White)
                    } else if app_state.highlight_enabled
                        && app_state.search_results.contains(&(*row, *col))
                    {
                        // Highlight search results only if highlighting is enabled
                        Style::default().bg(Color::Yellow).fg(Color::Black)
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

    // Create a table with the rows
    let mut table = Table::new(std::iter::once(header_cells).chain(rows));

    // Set table properties
    table = table
        .block(Block::default().borders(Borders::ALL))
        .highlight_style(Style::default().add_modifier(Modifier::BOLD))
        .highlight_symbol(">> ");

    // Set column constraints
    table = table.widths(&col_constraints);

    f.render_widget(table, area);
}

// Parse command input and identify keywords and parameters for highlighting
fn parse_command(input: &str) -> Vec<Span> {
    if input.is_empty() {
        return vec![Span::raw("")];
    }

    // Define known commands and their parameters
    let known_commands = [
        "w",
        "wq",
        "q",
        "q!",
        "x",
        "y",
        "d",
        "put",
        "pu",
        "nohlsearch",
        "noh",
        "help",
        "delsheet",
    ];

    // Commands with parameters
    let commands_with_params = ["cw", "ej", "sheet", "dr", "dc"];

    // Check if input is a simple command without parameters
    if known_commands.contains(&input) {
        return vec![Span::styled(input, Style::default().fg(Color::Yellow))];
    }

    // Check for commands with parameters
    for &cmd in &commands_with_params {
        if input.starts_with(cmd)
            && (input.len() == cmd.len() || input.chars().nth(cmd.len()) == Some(' '))
        {
            let mut spans = Vec::new();

            // Add the command part with yellow color
            spans.push(Span::styled(cmd, Style::default().fg(Color::Yellow)));

            // If there are parameters, add them with a different color
            if input.len() > cmd.len() {
                let params = &input[cmd.len()..];
                spans.push(Span::styled(params, Style::default().fg(Color::LightCyan)));
            }

            return spans;
        }
    }

    // Special case for "ej" command
    if input.starts_with("ej ") {
        let mut spans = Vec::new();
        let parts: Vec<&str> = input.split_whitespace().collect();

        // Handle "ej" command
        spans.push(Span::styled("ej", Style::default().fg(Color::Yellow)));

        // Add parameters if they exist
        if parts.len() > 1 {
            spans.push(Span::raw(" "));
            for i in 1..parts.len() {
                spans.push(Span::styled(
                    parts[i],
                    Style::default().fg(Color::LightCyan),
                ));
                if i < parts.len() - 1 {
                    spans.push(Span::raw(" "));
                }
            }
        }

        return spans;
    }

    // Special case for "cw" commands
    if input.starts_with("cw ") {
        let mut spans = Vec::new();
        let parts: Vec<&str> = input.split_whitespace().collect();

        spans.push(Span::styled("cw", Style::default().fg(Color::Yellow)));

        // Add parameters if they exist
        if parts.len() > 1 {
            spans.push(Span::raw(" "));
            for i in 1..parts.len() {
                let style = if parts[i] == "fit" || parts[i] == "min" || parts[i] == "all" {
                    Style::default().fg(Color::Yellow) // Keywords are yellow
                } else {
                    Style::default().fg(Color::LightCyan) // Parameters are cyan
                };

                spans.push(Span::styled(parts[i], style));
                if i < parts.len() - 1 {
                    spans.push(Span::raw(" "));
                }
            }
        }

        return spans;
    }

    // Special case for "dr" commands
    if input.starts_with("dr") {
        let mut spans = Vec::new();
        let parts: Vec<&str> = input.split_whitespace().collect();

        spans.push(Span::styled("dr", Style::default().fg(Color::Yellow)));

        // Add parameters if they exist
        if parts.len() > 1 {
            spans.push(Span::raw(" "));
            for i in 1..parts.len() {
                spans.push(Span::styled(
                    parts[i],
                    Style::default().fg(Color::LightCyan),
                ));
                if i < parts.len() - 1 {
                    spans.push(Span::raw(" "));
                }
            }
        }

        return spans;
    }

    // Special case for "dc" commands
    if input.starts_with("dc") {
        let mut spans = Vec::new();
        let parts: Vec<&str> = input.split_whitespace().collect();

        spans.push(Span::styled("dc", Style::default().fg(Color::Yellow)));

        // Add parameters if they exist
        if parts.len() > 1 {
            spans.push(Span::raw(" "));
            for i in 1..parts.len() {
                spans.push(Span::styled(
                    parts[i],
                    Style::default().fg(Color::LightCyan),
                ));
                if i < parts.len() - 1 {
                    spans.push(Span::raw(" "));
                }
            }
        }

        return spans;
    }

    // For cell references or unknown commands, return as is
    vec![Span::raw(input)]
}

fn draw_info_panel(f: &mut Frame, app_state: &AppState, area: Rect) {
    // Create a layout to split the info panel into two sections
    let chunks = Layout::default()
        .direction(Direction::Horizontal)
        .constraints([
            Constraint::Percentage(50), // Cell content/editing area
            Constraint::Percentage(50), // Notifications
        ])
        .split(area);

    // Get the cell reference
    let (row, col) = app_state.selected_cell;
    let cell_ref = cell_reference(app_state.selected_cell);

    // Handle the left panel based on the input mode
    match app_state.input_mode {
        InputMode::Editing => {
            // In editing mode, show the text area for editing
            let text_area = app_state.text_area.clone();

            // Create a block for the editing area
            let edit_block = Block::default()
                .borders(Borders::ALL)
                .title(format!(" Editing Cell {} ", cell_ref));

            // First render the block
            f.render_widget(edit_block.clone(), chunks[0]);

            let inner_area = edit_block.inner(chunks[0]);
            let padded_area = Rect {
                x: inner_area.x + 1, // Add 1 character padding on the left
                y: inner_area.y,
                width: inner_area.width.saturating_sub(2), // Subtract 2 for left and right padding
                height: inner_area.height,
            };
            f.render_widget(text_area.widget(), padded_area);
        }
        _ => {
            let content = app_state.get_cell_content(row, col);

            let cell_block = Block::default()
                .borders(Borders::ALL)
                .title(format!(" Cell {} Content ", cell_ref));

            f.render_widget(cell_block.clone(), chunks[0]);

            let inner_area = cell_block.inner(chunks[0]);
            let padded_area = Rect {
                x: inner_area.x + 1, // Add 1 character padding on the left
                y: inner_area.y,
                width: inner_area.width.saturating_sub(2), // Subtract 2 for left and right padding
                height: inner_area.height,
            };

            let cell_paragraph = Paragraph::new(content)
                .wrap(ratatui::widgets::Wrap { trim: false })
                .scroll((0, 0));

            f.render_widget(cell_paragraph, padded_area);
        }
    }

    let notification_block = Block::default()
        .borders(Borders::ALL)
        .title(" Notifications ");

    let notification_height = area.height as usize - 2; // Subtract 2 for the border

    let notifications_to_show = if app_state.notification_messages.len() > notification_height {
        let start_idx = app_state.notification_messages.len() - notification_height;
        app_state.notification_messages[start_idx..].to_vec()
    } else {
        app_state.notification_messages.clone()
    };

    let notifications_text = notifications_to_show.join("\n");

    f.render_widget(notification_block.clone(), chunks[1]);

    let inner_area = notification_block.inner(chunks[1]);
    let padded_area = Rect {
        x: inner_area.x + 1, // Add 1 character padding on the left
        y: inner_area.y,
        width: inner_area.width.saturating_sub(2), // Subtract 2 for left and right padding
        height: inner_area.height,
    };

    let notification_paragraph = Paragraph::new(notifications_text)
        .wrap(ratatui::widgets::Wrap { trim: false })
        .scroll((0, 0));

    f.render_widget(notification_paragraph, padded_area);
}

fn draw_status_bar(f: &mut Frame, app_state: &AppState, area: Rect) {
    match app_state.input_mode {
        InputMode::Normal => {
            let status =
                "hjkl=move 0=first-col ^=first-non-empty $=last-col gg=first-row G=last-row Ctrl+arrows=jump i=edit y=copy d=cut p=paste /=search ?=rev-search n=next N=prev :=command [ ]=prev/next-sheet +-=adjust-info-panel".to_string();

            let status_style = Style::default().bg(Color::Black).fg(Color::White);
            let status_widget = Paragraph::new(status).style(status_style);
            f.render_widget(status_widget, area);
        }

        InputMode::Editing => {
            let status = format!(
                "Editing cell {} (Enter=confirm, Esc=cancel)",
                cell_reference(app_state.selected_cell)
            );
            let status_style = Style::default().bg(Color::Black).fg(Color::White);
            let status_widget = Paragraph::new(status).style(status_style);
            f.render_widget(status_widget, area);
        }

        InputMode::Command => {
            // Create a styled text with different colors for command and parameters
            let mut spans = vec![Span::styled(":", Style::default().fg(Color::White))];
            let command_spans = parse_command(&app_state.input_buffer);
            spans.extend(command_spans);

            let text = Line::from(spans);
            let status_style = Style::default().bg(Color::Black);
            let status_widget = Paragraph::new(text).style(status_style);
            f.render_widget(status_widget, area);
        }

        InputMode::SearchForward => {
            let text_area = app_state.text_area.clone();

            let chunks = Layout::default()
                .direction(Direction::Horizontal)
                .constraints([Constraint::Length(1), Constraint::Min(1)])
                .split(area);

            let prefix_widget =
                Paragraph::new("/").style(Style::default().bg(Color::Black).fg(Color::White));
            f.render_widget(prefix_widget, chunks[0]);

            f.render_widget(text_area.widget(), chunks[1]);
        }

        InputMode::SearchBackward => {
            let text_area = app_state.text_area.clone();

            let chunks = Layout::default()
                .direction(Direction::Horizontal)
                .constraints([Constraint::Length(1), Constraint::Min(1)])
                .split(area);

            let prefix_widget =
                Paragraph::new("?").style(Style::default().bg(Color::Black).fg(Color::White));
            f.render_widget(prefix_widget, chunks[0]);

            f.render_widget(text_area.widget(), chunks[1]);
        }

        InputMode::Help => {
            let status = "Help Mode - Press Enter or Esc to close".to_string();
            let status_style = Style::default().bg(Color::Black).fg(Color::White);
            let status_widget = Paragraph::new(status).style(status_style);
            f.render_widget(status_widget, area);
        }
    }
}

fn draw_help_popup(f: &mut Frame, app_state: &mut AppState, area: Rect) {
    let overlay = Block::default()
        .style(Style::default().bg(Color::Black))
        .borders(Borders::NONE);
    f.render_widget(Clear, area);
    f.render_widget(overlay, area);

    let line_count = app_state.help_text.lines().count() as u16;

    let content_height = line_count + 2; // +2 for borders

    let max_line_width = app_state
        .help_text
        .lines()
        .map(|line| line.len() as u16)
        .max()
        .unwrap_or(40) as u16;

    let content_width = max_line_width + 4; // +4 for borders and padding

    let popup_width = content_width.min(area.width.saturating_sub(4));
    let popup_height = content_height.min(area.height.saturating_sub(4));

    // Center the popup on screen
    let popup_x = (area.width.saturating_sub(popup_width)) / 2;
    let popup_y = (area.height.saturating_sub(popup_height)) / 2;

    let popup_area = Rect::new(popup_x, popup_y, popup_width, popup_height);

    let visible_lines = popup_height.saturating_sub(2) as usize; // Subtract 2 for top and bottom borders
    app_state.help_visible_lines = visible_lines;

    let line_count = app_state.help_text.lines().count();
    let max_scroll = line_count.saturating_sub(visible_lines).max(0);

    app_state.help_scroll = app_state.help_scroll.min(max_scroll);

    let help_block = Block::default()
        .title(" HELP ")
        .title_style(
            Style::default()
                .fg(Color::Yellow)
                .add_modifier(Modifier::BOLD),
        )
        .borders(Borders::ALL)
        .border_style(Style::default().fg(Color::LightCyan))
        .style(Style::default().bg(Color::Blue).fg(Color::White));

    f.render_widget(help_block.clone(), popup_area);

    let inner_area = help_block.inner(popup_area);
    let padded_area = Rect {
        x: inner_area.x + 1, // Add 1 character padding on the left
        y: inner_area.y,
        width: inner_area.width.saturating_sub(2), // Subtract 2 for left and right padding
        height: inner_area.height,
    };

    let help_paragraph = Paragraph::new(app_state.help_text.clone())
        .wrap(ratatui::widgets::Wrap { trim: false })
        .scroll((app_state.help_scroll as u16, 0));

    f.render_widget(help_paragraph, padded_area);
}

fn handle_key_event(app_state: &mut AppState, key: KeyEvent) {
    match app_state.input_mode {
        InputMode::Normal => {
            if key.modifiers.contains(KeyModifiers::CONTROL) {
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
            let key_event = KeyEvent {
                code: key_code,
                modifiers: KeyModifiers::empty(),
                kind: KeyEventKind::Press,
                state: crossterm::event::KeyEventState::NONE,
            };
            app_state.text_area.input(key_event);
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
            let key_event = KeyEvent {
                code: key_code,
                modifiers: KeyModifiers::empty(),
                kind: KeyEventKind::Press,
                state: crossterm::event::KeyEventState::NONE,
            };
            app_state.text_area.input(key_event);
        }
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
        KeyCode::PageDown => {
            // Scroll down by a larger amount, but not beyond the last line
            app_state.help_scroll = (app_state.help_scroll + 10).min(max_scroll);
        }
        KeyCode::PageUp => {
            // Scroll up by a larger amount
            app_state.help_scroll = app_state.help_scroll.saturating_sub(10);
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

fn cell_reference(cell: (usize, usize)) -> String {
    format!("{}{}", index_to_col_name(cell.1), cell.0)
}

fn draw_title_with_tabs(f: &mut Frame, app_state: &AppState, area: Rect) {
    let sheet_names = app_state.workbook.get_sheet_names();
    let current_index = app_state.workbook.get_current_sheet_index();

    let file_name = app_state
        .file_path
        .file_name()
        .and_then(|n| n.to_str())
        .unwrap_or("Untitled");
    let title = format!(" {} | ", file_name);
    let title_width = title.len() as u16;

    let available_width = area.width.saturating_sub(title_width) as usize;

    let mut tab_widths = Vec::new();
    let mut total_width = 0;
    let mut visible_tabs = Vec::new();
    for (i, name) in sheet_names.iter().enumerate() {
        let tab_width = name.len() + 4;

        if total_width + tab_width <= available_width {
            tab_widths.push(tab_width as u16);
            total_width += tab_width;
            visible_tabs.push(i);
        } else {
            if !visible_tabs.contains(&current_index) {
                while !visible_tabs.is_empty() && total_width + tab_width > available_width {
                    let removed_width = tab_widths.remove(0) as usize;
                    visible_tabs.remove(0);
                    total_width -= removed_width;
                }

                if total_width + tab_width <= available_width {
                    tab_widths.push(tab_width as u16);
                    visible_tabs.push(current_index);
                }
            }
            break;
        }
    }

    let mut constraints = vec![Constraint::Length(title_width)];
    constraints.extend(tab_widths.iter().map(|&width| Constraint::Length(width)));

    if !constraints.is_empty() {
        let combined_layout = Layout::default()
            .direction(Direction::Horizontal)
            .constraints(constraints)
            .split(area);

        let title_widget =
            Paragraph::new(title).style(Style::default().bg(Color::Blue).fg(Color::White));
        f.render_widget(title_widget, combined_layout[0]);

        for (layout_idx, &sheet_idx) in visible_tabs.iter().enumerate() {
            let name = &sheet_names[sheet_idx];
            let is_current = sheet_idx == current_index;

            let tab_text = if is_current {
                format!("[{}]", name)
            } else {
                format!(" {} ", name)
            };

            let style = if is_current {
                Style::default().bg(Color::LightBlue).fg(Color::Black)
            } else {
                Style::default().bg(Color::DarkGray).fg(Color::White)
            };

            let tab_widget = Paragraph::new(tab_text).style(style);
            f.render_widget(tab_widget, combined_layout[layout_idx + 1]);
        }
    } else {
        let title_widget =
            Paragraph::new(title).style(Style::default().bg(Color::Blue).fg(Color::White));
        f.render_widget(title_widget, area);
    }

    if visible_tabs.len() < sheet_names.len() {
        let more_indicator = "...";
        let indicator_style = Style::default().bg(Color::DarkGray).fg(Color::White);
        let indicator_width = more_indicator.len() as u16;

        let indicator_rect = Rect {
            x: area.x + area.width - indicator_width,
            y: area.y,
            width: indicator_width,
            height: 1,
        };

        let indicator_widget = Paragraph::new(more_indicator).style(indicator_style);
        f.render_widget(indicator_widget, indicator_rect);
    }
}
