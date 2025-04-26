use anyhow::Result;
use crossterm::{
    ExecutableCommand,
    event::{self, Event, KeyEventKind},
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
use std::{io, time::Duration};

use crate::app::AppState;
use crate::app::InputMode;
use crate::ui::handlers::handle_key_event;
use crate::utils::cell_reference;
use crate::utils::index_to_col_name;

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
    // Calculate visible rows based on available height
    app_state.visible_rows = (area.height as usize).saturating_sub(3);

    // Ensure the selected column is visible
    app_state.ensure_column_visible(app_state.selected_cell.1);

    // Calculate available width for columns (subtract row numbers and borders)
    let available_width = (area.width as usize).saturating_sub(7); // 5 for row numbers + 2 for borders

    // Calculate how many columns can fit in the available width
    let mut visible_cols = 0;
    let mut width_used = 0;

    // Start from the leftmost visible column and add columns until run out of space
    for col_idx in app_state.start_col.. {
        let col_width = app_state.get_column_width(col_idx);

        // Always include the first column even if it's wider than available space
        if col_idx == app_state.start_col {
            width_used += col_width;
            visible_cols += 1;

            if width_used >= available_width {
                break;
            }
        }
        // For subsequent columns, add them if they fit completely
        else if width_used + col_width <= available_width {
            width_used += col_width;
            visible_cols += 1;
        }
        // Excel-like behavior: include one partially visible column if there's any space left
        else if width_used < available_width {
            visible_cols += 1;
            break;
        }
        // No more space available
        else {
            break;
        }
    }

    // Ensure at least one column is visible
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
    // Calculate visible row and column ranges
    let start_row = app_state.start_row;
    let end_row = start_row + app_state.visible_rows - 1;
    let start_col = app_state.start_col;
    let end_col = start_col + app_state.visible_cols - 1;

    let mut col_constraints = Vec::with_capacity(app_state.visible_cols + 1);
    col_constraints.push(Constraint::Length(5)); // Row header width

    for col in start_col..=end_col {
        col_constraints.push(Constraint::Length(app_state.get_column_width(col) as u16));
    }

    let mut header_cells = Vec::with_capacity(app_state.visible_cols + 1);
    header_cells.push(Cell::from(""));

    // Add column headers
    for col in start_col..=end_col {
        let col_name = index_to_col_name(col);
        header_cells
            .push(Cell::from(col_name).style(Style::default().bg(Color::Blue).fg(Color::White)));
    }

    let header_row = Row::new(header_cells).height(1);

    let mut rows = Vec::with_capacity(app_state.visible_rows);

    // Create data rows
    for row in start_row..=end_row {
        let row_header =
            Cell::from(row.to_string()).style(Style::default().bg(Color::Blue).fg(Color::White));

        let mut cells = Vec::with_capacity(app_state.visible_cols + 1);
        cells.push(row_header);

        // Add cells for this row
        for col in start_col..=end_col {
            let content = if app_state.selected_cell == (row, col)
                && matches!(app_state.input_mode, InputMode::Editing)
            {
                // Handle editing mode content
                let current_content = app_state.text_area.lines().join("\n");
                let col_width = app_state.get_column_width(col);

                // Calculate display width
                let display_width = current_content
                    .chars()
                    .fold(0, |acc, c| acc + if c.is_ascii() { 1 } else { 2 });

                if display_width > col_width.saturating_sub(2) {
                    // Truncate content if it's too wide
                    let mut result = String::with_capacity(col_width);
                    let mut cumulative_width = 0;

                    // Process characters from the end to show the most recent input
                    for c in current_content.chars().rev().take(col_width * 2) {
                        let char_width = if c.is_ascii() { 1 } else { 2 };
                        if cumulative_width + char_width <= col_width.saturating_sub(2) {
                            cumulative_width += char_width;
                            result.push(c);
                        } else {
                            break;
                        }
                    }

                    // Reverse the characters to get the correct order
                    result.chars().rev().collect::<String>()
                } else {
                    current_content
                }
            } else {
                // Handle normal cell content
                let content = app_state.get_cell_content(row, col);
                let col_width = app_state.get_column_width(col);

                // Calculate display width
                let display_width = content
                    .chars()
                    .fold(0, |acc, c| acc + if c.is_ascii() { 1 } else { 2 });

                if display_width > col_width {
                    // Truncate content if it's too wide
                    let mut result = String::with_capacity(col_width);
                    let mut current_width = 0;

                    for c in content.chars() {
                        let char_width = if c.is_ascii() { 1 } else { 2 };
                        if current_width + char_width < col_width {
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

            // Determine cell style
            let style = if app_state.selected_cell == (row, col) {
                Style::default().bg(Color::DarkGray).fg(Color::White)
            } else if app_state.highlight_enabled && app_state.search_results.contains(&(row, col))
            {
                Style::default().bg(Color::Yellow).fg(Color::Black)
            } else {
                Style::default()
            };

            cells.push(Cell::from(content).style(style));
        }

        rows.push(Row::new(cells));
    }

    let table = Table::new(std::iter::once(header_row).chain(rows))
        .block(Block::default().borders(Borders::ALL))
        .highlight_style(Style::default().add_modifier(Modifier::BOLD))
        .highlight_symbol(">> ")
        .widths(&col_constraints);

    f.render_widget(table, area);
}

// Parse command input and identify keywords and parameters for highlighting
fn parse_command(input: &str) -> Vec<Span> {
    if input.is_empty() {
        return vec![Span::raw("")];
    }

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

    let commands_with_params = ["cw", "ej", "eja", "sheet", "dr", "dc"];

    let special_keywords = ["fit", "min", "all", "h", "v", "horizontal", "vertical"];

    // Check if input is a simple command without parameters
    if known_commands.contains(&input) {
        return vec![Span::styled(input, Style::default().fg(Color::Yellow))];
    }

    // Extract command and parameters
    let parts: Vec<&str> = input.split_whitespace().collect();
    if parts.is_empty() {
        return vec![Span::raw(input)];
    }

    let cmd = parts[0];

    // Check if it's a known command with parameters
    if commands_with_params.contains(&cmd) || (cmd.starts_with("ej") && cmd.len() <= 3) {
        let mut spans = Vec::new();

        // Add the command part with yellow color
        spans.push(Span::styled(cmd, Style::default().fg(Color::Yellow)));

        // Add parameters if they exist
        if parts.len() > 1 {
            spans.push(Span::raw(" "));

            for i in 1..parts.len() {
                // Determine style based on whether it's a special keyword
                let style = if special_keywords.contains(&parts[i]) {
                    Style::default().fg(Color::Yellow) // Keywords are yellow
                } else {
                    Style::default().fg(Color::LightCyan) // Parameters are cyan
                };

                spans.push(Span::styled(parts[i], style));

                // Add space between parameters
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
    let chunks = Layout::default()
        .direction(Direction::Vertical)
        .constraints([
            Constraint::Percentage(50), // Cell content/editing area
            Constraint::Percentage(50), // Notifications
        ])
        .split(area);

    // Get the cell reference
    let (row, col) = app_state.selected_cell;
    let cell_ref = cell_reference(app_state.selected_cell);

    // Handle the top panel based on the input mode
    match app_state.input_mode {
        InputMode::Editing => {
            // In editing mode, show the text area for editing
            // Create a block for the editing area with title
            let title = format!(" Editing Cell {} ", cell_ref);
            let edit_block = Block::default().borders(Borders::ALL).title(title);

            f.render_widget(edit_block.clone(), chunks[0]);

            // Calculate inner area with padding
            let inner_area = edit_block.inner(chunks[0]);
            let padded_area = Rect {
                x: inner_area.x + 1, // Add 1 character padding on the left
                y: inner_area.y,
                width: inner_area.width.saturating_sub(2), // Subtract 2 for left and right padding
                height: inner_area.height,
            };

            f.render_widget(app_state.text_area.widget(), padded_area);
        }
        _ => {
            // Get cell content
            let content = app_state.get_cell_content(row, col);

            // Create block with title
            let title = format!(" Cell {} Content ", cell_ref);
            let cell_block = Block::default().borders(Borders::ALL).title(title);

            f.render_widget(cell_block.clone(), chunks[0]);

            // Calculate inner area with padding
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

    // Create notification block
    let notification_block = Block::default()
        .borders(Borders::ALL)
        .title(" Notifications ");

    f.render_widget(notification_block.clone(), chunks[1]);

    // Calculate inner area with padding
    let inner_area = notification_block.inner(chunks[1]);
    let padded_area = Rect {
        x: inner_area.x + 1, // Add 1 character padding on the left
        y: inner_area.y,
        width: inner_area.width.saturating_sub(2), // Subtract 2 for left and right padding
        height: inner_area.height,
    };

    // Calculate how many notifications can be shown
    let notification_height = inner_area.height as usize;

    let notifications_text = if app_state.notification_messages.is_empty() {
        String::new()
    } else if app_state.notification_messages.len() <= notification_height {
        app_state.notification_messages.join("\n")
    } else {
        let start_idx = app_state.notification_messages.len() - notification_height;

        let mut result = String::with_capacity(
            app_state.notification_messages[start_idx..]
                .iter()
                .map(|s| s.len())
                .sum::<usize>()
                + notification_height, // Account for newlines
        );

        for (i, msg) in app_state.notification_messages[start_idx..]
            .iter()
            .enumerate()
        {
            if i > 0 {
                result.push('\n');
            }
            result.push_str(msg);
        }

        result
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

        InputMode::Help => {}
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
        .unwrap_or(40);

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

    let mut title = " HELP ".to_string();

    let exit_instructions = " [ESC/Enter to close] ";
    title.push_str(exit_instructions);

    // Add scroll indicators if content is scrollable
    if max_scroll > 0 {
        let scroll_indicator = if app_state.help_scroll == 0 {
            " [↓ or j to scroll] "
        } else if app_state.help_scroll >= max_scroll {
            " [↑ or k to scroll] "
        } else {
            " [↑↓ or j/k to scroll] "
        };
        title.push_str(scroll_indicator);
    }

    let help_block = Block::default()
        .title(title)
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
