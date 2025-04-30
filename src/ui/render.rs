use anyhow::Result;
use crossterm::{
    event::{self, Event, KeyEventKind},
    terminal::{disable_raw_mode, enable_raw_mode, EnterAlternateScreen, LeaveAlternateScreen},
    ExecutableCommand,
};
use ratatui::{
    backend::CrosstermBackend,
    layout::{Constraint, Direction, Layout, Rect},
    style::{Color, Modifier, Style},
    text::{Line, Span},
    widgets::{Block, Borders, Cell, Clear, Paragraph, Row, Table},
    Frame, Terminal,
};
use std::{io, time::Duration};

use crate::app::AppState;
use crate::app::InputMode;
use crate::ui::handlers::handle_key_event;
use crate::utils::cell_reference;
use crate::utils::index_to_col_name;

pub fn run_app(mut app_state: AppState) -> Result<()> {
    // Setup terminal
    let mut terminal = setup_terminal()?;

    // Main event loop
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
    restore_terminal(&mut terminal)?;

    Ok(())
}

/// Setup the terminal for the application
fn setup_terminal() -> Result<Terminal<CrosstermBackend<io::Stdout>>> {
    enable_raw_mode()?;
    let mut stdout = io::stdout();
    stdout.execute(EnterAlternateScreen)?;

    let backend = CrosstermBackend::new(stdout);
    let terminal = Terminal::new(backend)?;

    Ok(terminal)
}

/// Restore the terminal to its original state
fn restore_terminal(terminal: &mut Terminal<CrosstermBackend<io::Stdout>>) -> Result<()> {
    disable_raw_mode()?;
    terminal.backend_mut().execute(LeaveAlternateScreen)?;
    terminal.show_cursor()?;

    Ok(())
}

/// Update the visible area of the spreadsheet based on the available space
fn update_visible_area(app_state: &mut AppState, area: Rect) {
    // Calculate visible rows based on available height (subtract header and borders)
    app_state.visible_rows = (area.height as usize).saturating_sub(3);

    // Ensure the selected column is visible
    app_state.ensure_column_visible(app_state.selected_cell.1);

    // Calculate available width for columns (subtract row numbers and borders)
    let available_width = (area.width as usize).saturating_sub(7); // 5 for row numbers + 2 for borders

    // Calculate how many columns can fit in the available width
    let mut visible_cols = 0;
    let mut width_used = 0;

    // Iterate through columns starting from the leftmost visible column
    for col_idx in app_state.start_col.. {
        let col_width = app_state.get_column_width(col_idx);

        if col_idx == app_state.start_col {
            // Always include the first column even if it's wider than available space
            width_used += col_width;
            visible_cols += 1;

            if width_used >= available_width {
                break;
            }
        } else if width_used + col_width <= available_width {
            // Add columns that fit completely
            width_used += col_width;
            visible_cols += 1;
        } else if width_used < available_width {
            // Excel-like behavior: include one partially visible column
            visible_cols += 1;
            break;
        } else {
            // No more space available
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

    let mut constraints = Vec::with_capacity(app_state.visible_cols + 1);
    constraints.push(Constraint::Length(5)); // Row header width

    for col in start_col..=end_col {
        constraints.push(Constraint::Length(app_state.get_column_width(col) as u16));
    }

    // Set table style based on current mode
    let (table_block, header_style, cell_style) =
        if matches!(app_state.input_mode, InputMode::Normal) {
            // In Normal mode, add color to the border of the data display area to indicate current focus
            (
                Block::default()
                    .borders(Borders::ALL)
                    .border_style(Style::default().fg(Color::LightCyan)),
                Style::default().bg(Color::DarkGray).fg(Color::Gray),
                Style::default(),
            )
        } else {
            // In editing mode, dim the data display area
            (
                Block::default().borders(Borders::ALL),
                Style::default().fg(Color::DarkGray),
                Style::default().fg(Color::DarkGray), // Dimmed cell content
            )
        };

    // Create header row
    let mut header_cells = Vec::with_capacity(app_state.visible_cols + 1);
    header_cells.push(Cell::from("").style(header_style));

    // Add column headers
    for col in start_col..=end_col {
        let col_name = index_to_col_name(col);
        header_cells.push(Cell::from(col_name).style(header_style));
    }

    let header = Row::new(header_cells).height(1);

    // Create data rows
    let rows = (start_row..=end_row).map(|row| {
        let mut cells = Vec::with_capacity(app_state.visible_cols + 1);

        // Add row header
        cells.push(Cell::from(row.to_string()).style(header_style));

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
                Style::default().bg(Color::White).fg(Color::Black)
            } else if app_state.highlight_enabled && app_state.search_results.contains(&(row, col))
            {
                Style::default().bg(Color::Yellow).fg(Color::Black)
            } else {
                Style::default()
            };

            cells.push(Cell::from(content).style(style));
        }

        Row::new(cells)
    });

    // Create table with header and rows
    let table = Table::new(
        // Combine header and data rows
        std::iter::once(header).chain(rows),
    )
    .block(table_block)
    .style(cell_style)
    .widths(&constraints);

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

fn draw_info_panel(f: &mut Frame, app_state: &mut AppState, area: Rect) {
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
            let (vim_mode_str, mode_color) = if let Some(vim_state) = &app_state.vim_state {
                match vim_state.mode {
                    crate::app::VimMode::Normal => ("NORMAL", Color::Green),
                    crate::app::VimMode::Insert => ("INSERT", Color::LightBlue),
                    crate::app::VimMode::Visual => ("VISUAL", Color::Yellow),
                    crate::app::VimMode::Operator(op) => {
                        let op_str = match op {
                            'y' => "YANK",
                            'd' => "DELETE",
                            'c' => "CHANGE",
                            _ => "OPERATOR",
                        };
                        (op_str, Color::LightRed)
                    }
                }
            } else {
                ("VIM", Color::White)
            };

            let title = Line::from(vec![
                Span::raw(" Editing Cell "),
                Span::raw(cell_ref.clone()),
                Span::raw(" - "),
                Span::styled(
                    vim_mode_str,
                    Style::default().fg(mode_color).add_modifier(Modifier::BOLD),
                ),
                Span::raw(" "),
            ]);

            let edit_block = Block::default()
                .borders(Borders::ALL)
                .border_style(Style::default().fg(Color::LightCyan))
                .title(title);

            // Calculate inner area with padding
            let inner_area = edit_block.inner(chunks[0]);
            let padded_area = Rect {
                x: inner_area.x + 1, // Add 1 character padding on the left
                y: inner_area.y,
                width: inner_area.width.saturating_sub(2), // Subtract 2 for left and right padding
                height: inner_area.height,
            };

            f.render_widget(edit_block, chunks[0]);
            f.render_widget(app_state.text_area.widget(), padded_area);
        }
        _ => {
            // Get cell content
            let content = app_state.get_cell_content(row, col);

            let title = format!(" Cell {} Content ", cell_ref);
            let cell_block = Block::default().borders(Borders::ALL).title(title);

            // Create paragraph with cell content
            let cell_paragraph = Paragraph::new(content)
                .block(cell_block)
                .wrap(ratatui::widgets::Wrap { trim: false });

            f.render_widget(cell_paragraph, chunks[0]);
        }
    }

    // Create notification block
    let notification_block = if matches!(app_state.input_mode, InputMode::Editing) {
        Block::default()
            .borders(Borders::ALL)
            .border_style(Style::default().fg(Color::DarkGray))
            .title(Span::styled(
                " Notifications ",
                Style::default().fg(Color::DarkGray),
            ))
    } else {
        Block::default()
            .borders(Borders::ALL)
            .title(" Notifications ")
    };

    // Calculate how many notifications can be shown
    let notification_height = notification_block.inner(chunks[1]).height as usize;

    // Prepare notifications text
    let notifications_text = if app_state.notification_messages.is_empty() {
        String::new()
    } else if app_state.notification_messages.len() <= notification_height {
        app_state.notification_messages.join("\n")
    } else {
        // Show only the most recent notifications that fit
        let start_idx = app_state.notification_messages.len() - notification_height;
        app_state.notification_messages[start_idx..].join("\n")
    };

    let notification_paragraph = Paragraph::new(notifications_text)
        .block(notification_block)
        .wrap(ratatui::widgets::Wrap { trim: false })
        .style(if matches!(app_state.input_mode, InputMode::Editing) {
            Style::default().fg(Color::DarkGray)
        } else {
            Style::default()
        });

    f.render_widget(notification_paragraph, chunks[1]);
}

fn draw_status_bar(f: &mut Frame, app_state: &AppState, area: Rect) {
    match app_state.input_mode {
        InputMode::Normal => {
            let status = "Input :help for operating instructions | hjkl=move [ ]=prev/next-sheet Enter=edit y=copy d=cut p=paste /=search N/n=prev/next-search-result :=command ";

            let status_widget = Paragraph::new(status)
                .style(Style::default())
                .alignment(ratatui::layout::Alignment::Left);

            f.render_widget(status_widget, area);
        }

        InputMode::Editing => {
            let status_widget = Paragraph::new("Press Esc to exit editing mode")
                .style(Style::default().fg(Color::DarkGray))
                .alignment(ratatui::layout::Alignment::Left);

            f.render_widget(status_widget, area);
        }

        InputMode::Command => {
            // Create a styled text with different colors for command and parameters
            let mut spans = vec![Span::styled(":", Style::default())];
            let command_spans = parse_command(&app_state.input_buffer);
            spans.extend(command_spans);

            let text = Line::from(spans);
            let status_widget = Paragraph::new(text)
                .style(Style::default())
                .alignment(ratatui::layout::Alignment::Left);

            f.render_widget(status_widget, area);
        }

        InputMode::SearchForward | InputMode::SearchBackward => {
            // Get search prefix based on mode
            let prefix = if matches!(app_state.input_mode, InputMode::SearchForward) {
                "/"
            } else {
                "?"
            };

            // Split the area for search prefix and search input
            let chunks = Layout::default()
                .direction(Direction::Horizontal)
                .constraints([
                    Constraint::Length(1), // Search prefix
                    Constraint::Min(1),    // Search input
                ])
                .split(area);

            // Render search prefix
            let prefix_widget = Paragraph::new(prefix)
                .style(Style::default())
                .alignment(ratatui::layout::Alignment::Left);

            f.render_widget(prefix_widget, chunks[0]);

            // Render search input with cursor visible
            let mut text_area = app_state.text_area.clone();
            text_area.set_cursor_line_style(Style::default());
            text_area.set_cursor_style(Style::default().add_modifier(Modifier::REVERSED));

            f.render_widget(text_area.widget(), chunks[1]);
        }

        InputMode::Help => {
            // No status bar in help mode
        }
    }
}

fn draw_help_popup(f: &mut Frame, app_state: &mut AppState, area: Rect) {
    // Clear the background
    f.render_widget(Clear, area);

    // Calculate popup dimensions
    let line_count = app_state.help_text.lines().count() as u16;
    let content_height = line_count + 2; // +2 for borders

    let max_line_width = app_state
        .help_text
        .lines()
        .map(|line| line.len() as u16)
        .max()
        .unwrap_or(40);

    let content_width = max_line_width + 4; // +4 for borders and padding

    // Ensure popup fits within screen
    let popup_width = content_width.min(area.width.saturating_sub(4));
    let popup_height = content_height.min(area.height.saturating_sub(4));

    // Center the popup on screen
    let popup_x = (area.width.saturating_sub(popup_width)) / 2;
    let popup_y = (area.height.saturating_sub(popup_height)) / 2;

    let popup_area = Rect::new(popup_x, popup_y, popup_width, popup_height);

    // Calculate scrolling parameters
    let visible_lines = popup_height.saturating_sub(2) as usize; // Subtract 2 for top and bottom borders
    app_state.help_visible_lines = visible_lines;

    let line_count = app_state.help_text.lines().count();
    let max_scroll = line_count.saturating_sub(visible_lines).max(0);

    app_state.help_scroll = app_state.help_scroll.min(max_scroll);

    let mut title = " [ESC/Enter to close] ".to_string();

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

    // Create paragraph with help text
    let help_paragraph = Paragraph::new(app_state.help_text.clone())
        .block(help_block)
        .wrap(ratatui::widgets::Wrap { trim: false })
        .scroll((app_state.help_scroll as u16, 0));

    f.render_widget(help_paragraph, popup_area);
}

fn draw_title_with_tabs(f: &mut Frame, app_state: &AppState, area: Rect) {
    let is_editing = matches!(app_state.input_mode, InputMode::Editing);
    let sheet_names = app_state.workbook.get_sheet_names();
    let current_index = app_state.workbook.get_current_sheet_index();

    let file_name = app_state
        .file_path
        .file_name()
        .and_then(|n| n.to_str())
        .unwrap_or("Untitled");

    let title_content = format!(" {} ", file_name);

    let title_width = title_content
        .chars()
        .fold(0, |acc, c| acc + if c.is_ascii() { 1 } else { 2 }) as u16;

    let available_width = area.width.saturating_sub(title_width) as usize;

    let mut tab_widths = Vec::new();
    let mut total_width = 0;
    let mut visible_tabs = Vec::new();

    for (i, name) in sheet_names.iter().enumerate() {
        let tab_width = name.len();

        if total_width + tab_width <= available_width {
            tab_widths.push(tab_width as u16);
            total_width += tab_width;
            visible_tabs.push(i);
        } else {
            // If current tab isn't visible, make room for it
            if !visible_tabs.contains(&current_index) {
                // Remove tabs from the beginning until there's enough space
                while !visible_tabs.is_empty() && total_width + tab_width > available_width {
                    let removed_width = tab_widths.remove(0) as usize;
                    visible_tabs.remove(0);
                    total_width -= removed_width;
                }

                // Add current tab if there's now enough space
                if total_width + tab_width <= available_width {
                    tab_widths.push(tab_width as u16);
                    visible_tabs.push(current_index);
                }
            }
            break;
        }
    }

    // Limit title width to at most 2/3 of the area
    let max_title_width = (area.width * 2 / 3).min(title_width);

    // Create a two-column layout: title column and tab column
    let horizontal_layout = Layout::default()
        .direction(Direction::Horizontal)
        .constraints([Constraint::Length(max_title_width), Constraint::Min(0)])
        .split(area);

    let title_style = if is_editing {
        Style::default().bg(Color::DarkGray).fg(Color::Gray)
    } else {
        Style::default().bg(Color::DarkGray).fg(Color::White)
    };

    let title_widget = Paragraph::new(title_content).style(title_style);

    f.render_widget(title_widget, horizontal_layout[0]);

    // Create constraints for tab layout
    let mut tab_constraints = Vec::new();
    for &width in &tab_widths {
        tab_constraints.push(Constraint::Length(width));
    }
    tab_constraints.push(Constraint::Min(0)); // Filler space

    let tab_layout = Layout::default()
        .direction(Direction::Horizontal)
        .constraints(tab_constraints)
        .split(horizontal_layout[1]);

    // Render each visible tab
    for (layout_idx, &sheet_idx) in visible_tabs.iter().enumerate() {
        if layout_idx >= tab_layout.len() - 1 {
            break;
        }

        let name = &sheet_names[sheet_idx];
        let is_current = sheet_idx == current_index;

        let style = if is_editing {
            if is_current {
                Style::default().bg(Color::DarkGray).fg(Color::Gray)
            } else {
                Style::default().fg(Color::DarkGray)
            }
        } else if is_current {
            Style::default().bg(Color::DarkGray).fg(Color::White)
        } else {
            Style::default()
        };

        let tab_widget = Paragraph::new(name.to_string())
            .style(style)
            .alignment(ratatui::layout::Alignment::Center);

        f.render_widget(tab_widget, tab_layout[layout_idx]);
    }

    // Show indicator if not all tabs are visible
    if visible_tabs.len() < sheet_names.len() {
        let more_indicator = "...";
        let indicator_style = Style::default().bg(Color::DarkGray).fg(Color::White);
        let indicator_width = more_indicator.len() as u16;

        // Position indicator at the right edge
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
