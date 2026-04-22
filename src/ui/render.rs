use anyhow::Result;
use crossterm::{
    event::{self, Event, KeyEventKind},
    terminal::{disable_raw_mode, enable_raw_mode, EnterAlternateScreen, LeaveAlternateScreen},
    ExecutableCommand,
};
use ratatui::{
    backend::CrosstermBackend,
    layout::{Alignment, Constraint, Direction, Layout, Rect},
    style::{Color, Modifier, Style},
    text::{Line, Span},
    widgets::{Block, Borders, Cell, Clear, Paragraph, Row, Table},
    Frame, Terminal,
};
use std::{io, time::Duration};

use crate::app::AppState;
use crate::app::HelpEntry;
use crate::app::HelpSection;
use crate::app::InputMode;
use crate::app::VimMode;
use crate::app::LEFT_HELP_SECTIONS;
use crate::app::RIGHT_HELP_SECTIONS;
use crate::ui::handlers::handle_key_event;
use crate::utils::cell_reference;
use crate::utils::index_to_col_name;

const HELP_ENTRY_INDENT: u16 = 2;
const HELP_ENTRY_GAP: u16 = 1;

mod theme {
    use ratatui::style::{Color, Style};

    pub const BACKGROUND: Color = Color::Rgb(11, 16, 32);
    pub const SURFACE: Color = Color::Rgb(17, 24, 39);
    pub const SURFACE_MUTED: Color = Color::Rgb(31, 41, 55);
    pub const GRID: Color = Color::Rgb(55, 65, 81);
    pub const TEXT: Color = Color::Rgb(229, 231, 235);
    pub const TEXT_SECONDARY: Color = Color::Rgb(156, 163, 175);
    pub const TEXT_DISABLED: Color = Color::Rgb(107, 114, 128);
    pub const ACCENT: Color = Color::Rgb(56, 189, 248);
    pub const SEARCH: Color = Color::Rgb(250, 204, 21);
    pub const WARNING: Color = Color::Rgb(245, 158, 11);
    pub const SUCCESS: Color = Color::Rgb(34, 197, 94);

    pub fn base() -> Style {
        Style::default().bg(BACKGROUND).fg(TEXT)
    }

    pub fn surface() -> Style {
        Style::default().bg(SURFACE).fg(TEXT)
    }

    pub fn muted() -> Style {
        Style::default().bg(SURFACE_MUTED).fg(TEXT_SECONDARY)
    }
}

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

    // Update row number width based on the maximum row number
    app_state.update_row_number_width();

    // Calculate available width for columns (subtract row numbers and borders)
    let available_width = (area.width as usize).saturating_sub(app_state.row_number_width + 2); // row_number_width + 2 for borders

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
    f.render_widget(Clear, f.size());
    let status_bar_height = status_bar_height(app_state, f.size().width);

    let chunks = Layout::default()
        .direction(Direction::Vertical)
        .constraints([
            Constraint::Length(1),
            Constraint::Min(1),
            Constraint::Length(app_state.info_panel_height as u16),
            Constraint::Length(status_bar_height),
        ])
        .split(f.size());

    draw_title_with_tabs(f, app_state, chunks[0]);

    update_visible_area(app_state, chunks[1]);
    draw_spreadsheet(f, app_state, chunks[1]);
    draw_info_panel(f, app_state, chunks[2]);
    if status_bar_height > 0 {
        draw_status_bar(f, app_state, chunks[3]);
    }

    // If in help mode, draw the help popup over everything else
    if let InputMode::Help = app_state.input_mode {
        draw_help_popup(f, app_state, f.size());
    }

    // If in lazy loading mode or CommandInLazyLoading mode and the current sheet is not loaded, draw the lazy loading overlay
    match app_state.input_mode {
        InputMode::LazyLoading | InputMode::CommandInLazyLoading => {
            let current_index = app_state.workbook.get_current_sheet_index();
            if !app_state.workbook.is_sheet_loaded(current_index) {
                draw_lazy_loading_overlay(f, app_state, chunks[1]);
            } else if matches!(app_state.input_mode, InputMode::LazyLoading) {
                // If the sheet is loaded, switch back to Normal mode
                app_state.input_mode = crate::app::InputMode::Normal;
            }
        }
        _ => {}
    }
}

fn draw_spreadsheet(f: &mut Frame, app_state: &AppState, area: Rect) {
    // Calculate visible row and column ranges
    let start_row = app_state.start_row;
    let end_row = start_row + app_state.visible_rows - 1;
    let start_col = app_state.start_col;
    let end_col = start_col + app_state.visible_cols - 1;

    let mut constraints = Vec::with_capacity(app_state.visible_cols + 1);
    constraints.push(Constraint::Length(app_state.row_number_width as u16)); // Dynamic row header width

    for col in start_col..=end_col {
        constraints.push(Constraint::Length(app_state.get_column_width(col) as u16));
    }

    // Set table style based on current mode
    let is_editing = matches!(app_state.input_mode, InputMode::Editing);
    let table_block = Block::default()
        .style(theme::base())
        .borders(Borders::ALL)
        .border_style(if is_editing {
            Style::default().fg(theme::GRID)
        } else {
            Style::default().fg(theme::ACCENT)
        });
    let header_style = if is_editing {
        Style::default()
            .bg(theme::SURFACE_MUTED)
            .fg(theme::TEXT_DISABLED)
    } else {
        theme::muted()
    };
    let cell_style = if is_editing {
        Style::default()
            .bg(theme::BACKGROUND)
            .fg(theme::TEXT_DISABLED)
    } else {
        theme::base()
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
                Style::default().bg(theme::SEARCH).fg(Color::Black)
            } else {
                cell_style
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
fn parse_command(input: &str) -> Vec<Span<'_>> {
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
        "addsheet",
        "delsheet",
    ];

    let commands_with_params = ["cw", "ej", "eja", "sheet", "dr", "dc", "addsheet"];

    let special_keywords = ["fit", "min", "all", "h", "v", "horizontal", "vertical"];

    // Check if input is a simple command without parameters
    if known_commands.contains(&input) {
        return vec![Span::styled(input, Style::default().fg(theme::WARNING))];
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

        spans.push(Span::styled(cmd, Style::default().fg(theme::WARNING)));

        // Add parameters if they exist
        if parts.len() > 1 {
            spans.push(Span::raw(" "));

            for i in 1..parts.len() {
                // Determine style based on whether it's a special keyword
                let style = if special_keywords.contains(&parts[i]) {
                    Style::default().fg(theme::WARNING)
                } else {
                    Style::default().fg(theme::ACCENT)
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

fn display_width(text: &str) -> u16 {
    text.chars()
        .fold(0, |acc, ch| acc + if ch.is_ascii() { 1 } else { 2 })
}

fn status_bar_height(app_state: &AppState, width: u16) -> u16 {
    let _ = width;
    if matches!(app_state.input_mode, InputMode::Help) {
        0
    } else {
        1
    }
}

fn status_bar_style() -> Style {
    Style::default().bg(Color::Black).fg(theme::TEXT)
}

fn status_badge(label: &'static str, color: Color) -> Span<'static> {
    Span::styled(
        format!(" {label} "),
        Style::default()
            .bg(color)
            .fg(Color::Black)
            .add_modifier(Modifier::BOLD),
    )
}

fn subtle_span(text: impl Into<String>) -> Span<'static> {
    Span::styled(text.into(), Style::default().fg(theme::TEXT_SECONDARY))
}

fn shortcut_key(key: &str) -> Span<'static> {
    Span::styled(
        format!("[{key}]"),
        Style::default()
            .bg(theme::SURFACE_MUTED)
            .fg(theme::ACCENT)
            .add_modifier(Modifier::BOLD),
    )
}

fn shortcut_spans(entries: &[(&str, &str)]) -> Vec<Span<'static>> {
    let mut spans = Vec::new();

    for (index, (key, label)) in entries.iter().enumerate() {
        if index > 0 {
            spans.push(Span::raw("   "));
        }
        spans.push(shortcut_key(key));
        spans.push(Span::raw(" "));
        spans.push(Span::styled(
            (*label).to_string(),
            Style::default().fg(theme::TEXT),
        ));
    }

    spans
}

fn render_single_status_line<'a>(
    f: &mut Frame,
    area: Rect,
    line: Line<'a>,
    alignment: ratatui::layout::Alignment,
) {
    let status_widget = Paragraph::new(line)
        .style(status_bar_style())
        .alignment(alignment);
    f.render_widget(status_widget, area);
}

fn line_display_width(line: &Line<'_>) -> u16 {
    line.spans
        .iter()
        .map(|span| display_width(&span.content))
        .sum()
}

fn render_status_sections<'a, 'b>(
    f: &mut Frame,
    area: Rect,
    left: Line<'a>,
    right: Option<Line<'b>>,
) {
    let Some(right_line) = right else {
        render_single_status_line(f, area, left, ratatui::layout::Alignment::Left);
        return;
    };

    let right_width = line_display_width(&right_line).saturating_add(1);
    if right_width >= area.width {
        render_single_status_line(f, area, right_line, ratatui::layout::Alignment::Right);
        return;
    }

    let sections = Layout::default()
        .direction(Direction::Horizontal)
        .constraints([
            Constraint::Min(area.width.saturating_sub(right_width)),
            Constraint::Length(right_width),
        ])
        .split(area);

    render_single_status_line(f, sections[0], left, ratatui::layout::Alignment::Left);
    render_single_status_line(
        f,
        sections[1],
        right_line,
        ratatui::layout::Alignment::Right,
    );
}

fn draw_info_panel(f: &mut Frame, app_state: &mut AppState, area: Rect) {
    if area.height < 4 {
        if matches!(app_state.input_mode, InputMode::Editing) {
            draw_editing_panel(f, app_state, area);
        } else {
            draw_cell_details(f, app_state, area);
        }
        return;
    }

    let chunks = Layout::default()
        .direction(Direction::Vertical)
        .constraints([Constraint::Percentage(50), Constraint::Percentage(50)])
        .split(area);

    if matches!(app_state.input_mode, InputMode::Editing) {
        draw_editing_panel(f, app_state, chunks[0]);
    } else {
        draw_cell_details(f, app_state, chunks[0]);
    }
    draw_notifications(f, app_state, chunks[1]);
}

fn draw_cell_details(f: &mut Frame, app_state: &AppState, area: Rect) {
    let content = app_state.get_cell_content(app_state.selected_cell.0, app_state.selected_cell.1);
    let cell_ref = cell_reference(app_state.selected_cell);
    let value_type = cell_value_type(&content);
    let length = content.chars().count();

    let title = format!(" Cell {cell_ref}  {value_type}  Len {length} ");
    let block = panel_block(title, theme::TEXT);
    let paragraph = Paragraph::new(content)
        .block(block)
        .style(theme::surface())
        .wrap(ratatui::widgets::Wrap { trim: false });
    f.render_widget(paragraph, area);
}

fn draw_editing_panel(f: &mut Frame, app_state: &AppState, area: Rect) {
    let cell_ref = cell_reference(app_state.selected_cell);
    let mode = app_state.vim_state.as_ref().map(|state| state.mode);
    let input_block = panel_block_line(editing_title_line(cell_ref, mode), theme::ACCENT);
    let inner_area = input_block.inner(area);
    let padded_area = Rect {
        x: inner_area.x.saturating_add(1),
        y: inner_area.y,
        width: inner_area.width.saturating_sub(2),
        height: inner_area.height,
    };

    f.render_widget(input_block, area);
    f.render_widget(app_state.text_area.widget(), padded_area);
}

fn draw_notifications(f: &mut Frame, app_state: &AppState, area: Rect) {
    let lines = if app_state.notification_messages.is_empty() {
        vec![Line::from(Span::styled(
            "No notifications",
            Style::default().fg(theme::TEXT_SECONDARY),
        ))]
    } else {
        app_state
            .notification_messages
            .iter()
            .rev()
            .take(4)
            .enumerate()
            .map(|(index, message)| {
                let color = if index == 0 {
                    theme::TEXT
                } else {
                    theme::TEXT_SECONDARY
                };
                Line::from(Span::styled(message.clone(), Style::default().fg(color)))
            })
            .collect()
    };

    let paragraph = Paragraph::new(lines)
        .block(panel_block(" NOTIFICATIONS ".to_string(), theme::TEXT))
        .style(theme::surface())
        .wrap(ratatui::widgets::Wrap { trim: false });
    f.render_widget(paragraph, area);
}

fn panel_block(title: String, border_color: Color) -> Block<'static> {
    panel_block_line(
        Line::from(Span::styled(
            title,
            Style::default()
                .fg(theme::TEXT)
                .add_modifier(Modifier::BOLD),
        )),
        border_color,
    )
}

fn panel_block_line(title: Line<'static>, border_color: Color) -> Block<'static> {
    Block::default()
        .borders(Borders::ALL)
        .title(title)
        .border_style(Style::default().fg(border_color))
        .style(theme::surface())
}

fn editing_title_line(cell_ref: String, mode: Option<VimMode>) -> Line<'static> {
    let mut spans = vec![
        Span::styled(
            " Editing Cell ",
            Style::default()
                .fg(theme::TEXT)
                .add_modifier(Modifier::BOLD),
        ),
        Span::styled(
            cell_ref,
            Style::default()
                .fg(theme::TEXT)
                .add_modifier(Modifier::BOLD),
        ),
    ];

    if let Some(mode) = mode {
        spans.push(Span::styled(
            " - ",
            Style::default()
                .fg(theme::TEXT)
                .add_modifier(Modifier::BOLD),
        ));
        spans.push(Span::styled(
            mode.to_string(),
            Style::default()
                .fg(vim_mode_color(mode))
                .add_modifier(Modifier::BOLD),
        ));
    }

    spans.push(Span::styled(
        " ",
        Style::default()
            .fg(theme::TEXT)
            .add_modifier(Modifier::BOLD),
    ));

    Line::from(spans)
}

fn vim_mode_color(mode: VimMode) -> Color {
    match mode {
        VimMode::Normal => theme::SUCCESS,
        VimMode::Insert => theme::ACCENT,
        VimMode::Visual => theme::SEARCH,
        VimMode::Operator(_) => theme::WARNING,
    }
}

fn cell_value_type(content: &str) -> &'static str {
    if content.is_empty() {
        "Blank"
    } else if content.starts_with("Formula: ") {
        "Formula"
    } else if content.parse::<f64>().is_ok() {
        "Number"
    } else {
        "String"
    }
}

fn draw_status_bar(f: &mut Frame, app_state: &AppState, area: Rect) {
    match app_state.input_mode {
        InputMode::Normal => {
            let left = Line::from(vec![status_badge("NORMAL", theme::ACCENT)]);
            let right = Line::from(shortcut_spans(&[
                ("Enter", "Edit"),
                (":", "Command"),
                ("/", "Search"),
                (":w", "Save"),
            ]));
            render_status_sections(f, area, left, Some(right));
        }

        InputMode::Editing => {
            let left = Line::from(vec![status_badge("EDIT", theme::SUCCESS)]);
            let right = Line::from(shortcut_spans(&[
                ("Enter", "Save"),
                ("Esc", "Normal"),
                ("i", "Insert"),
                ("v", "Visual"),
            ]));
            render_status_sections(f, area, left, Some(right));
        }

        InputMode::Command | InputMode::CommandInLazyLoading => {
            let mut left_spans = vec![
                status_badge("COMMAND", theme::WARNING),
                Span::raw("  "),
                Span::styled(":", Style::default().fg(theme::TEXT)),
            ];
            left_spans.extend(parse_command(&app_state.input_buffer));
            let right = Line::from(shortcut_spans(&[
                ("Enter", "Run"),
                ("Esc", "Cancel"),
                ("A1", "Jump"),
            ]));
            render_status_sections(f, area, Line::from(left_spans), Some(right));
        }

        InputMode::SearchForward | InputMode::SearchBackward => {
            let prefix = if matches!(app_state.input_mode, InputMode::SearchForward) {
                "/"
            } else {
                "?"
            };
            let query = app_state.text_area.lines().join("\n");
            let left_spans = vec![
                status_badge("SEARCH", theme::SEARCH),
                Span::raw("  "),
                Span::styled(prefix.to_string(), Style::default().fg(theme::TEXT)),
                Span::styled(query, Style::default().fg(theme::TEXT)),
            ];
            let right = Line::from(shortcut_spans(&[
                ("Enter", "Apply"),
                ("Esc", "Cancel"),
                ("n/N", "Navigate"),
            ]));
            render_status_sections(f, area, Line::from(left_spans), Some(right));
        }

        InputMode::Help => {
            // No status bar in help mode
        }

        InputMode::LazyLoading => {
            let left = Line::from(vec![
                status_badge("LAZY", theme::WARNING),
                Span::raw("  "),
                subtle_span("State "),
                Span::styled("not loaded", Style::default().fg(theme::WARNING)),
            ]);
            let right = Line::from(shortcut_spans(&[
                ("Enter", "Load"),
                ("[ ]", "Switch"),
                (":", "Command"),
            ]));
            render_status_sections(f, area, left, Some(right));
        }
    }
}

fn sheet_rows_cols(app_state: &AppState) -> String {
    let sheet = app_state.workbook.get_current_sheet();
    format!("{} x {}", sheet.max_rows, sheet.max_cols)
}

fn draw_lazy_loading_overlay(f: &mut Frame, _app_state: &AppState, area: Rect) {
    // Create a semi-transparent overlay
    let overlay = Block::default()
        .style(theme::surface())
        .borders(Borders::ALL)
        .border_style(Style::default().fg(theme::ACCENT));

    f.render_widget(Clear, area);
    f.render_widget(overlay, area);

    // Calculate center position for the message
    let message = "Sheet not loaded   Enter load   [ ] switch sheet   : command";
    let width = message.len() as u16;
    let x = area.x + (area.width.saturating_sub(width)) / 2;
    let y = area.y + area.height / 2;

    if x < area.width && y < area.height {
        let message_area = Rect {
            x,
            y,
            width: width.min(area.width),
            height: 1,
        };

        let message_widget = Paragraph::new(message).style(
            Style::default()
                .fg(theme::WARNING)
                .add_modifier(Modifier::BOLD),
        );

        f.render_widget(message_widget, message_area);
    }
}

fn draw_help_popup(f: &mut Frame, app_state: &mut AppState, area: Rect) {
    let popup_area = help_popup_area(area);
    let block = Block::default()
        .title(" COMMAND HELP ")
        .title_alignment(Alignment::Center)
        .title_style(
            Style::default()
                .fg(theme::ACCENT)
                .add_modifier(Modifier::BOLD),
        )
        .borders(Borders::ALL)
        .border_style(Style::default().fg(theme::TEXT_SECONDARY))
        .style(theme::surface());
    let inner = block.inner(popup_area);

    f.render_widget(Clear, area);
    f.render_widget(Block::default().style(theme::base()), area);
    f.render_widget(Clear, popup_area);
    f.render_widget(block, popup_area);

    let Some((content_area, divider_area, footer_area)) = help_popup_inner_areas(inner) else {
        return;
    };

    let lines = help_overlay_lines(content_area.width);
    let visible_lines = content_area.height.max(1) as usize;
    app_state.help_visible_lines = visible_lines;
    app_state.help_total_lines = lines.len();
    let max_scroll = lines.len().saturating_sub(visible_lines);
    app_state.help_scroll = app_state.help_scroll.min(max_scroll);

    let help_paragraph = Paragraph::new(lines)
        .style(theme::surface())
        .scroll((app_state.help_scroll as u16, 0));
    f.render_widget(help_paragraph, content_area);

    let divider = Paragraph::new("-".repeat(inner.width as usize)).style(theme::surface());
    f.render_widget(divider, divider_area);
    render_help_footer(
        f,
        footer_area,
        app_state.help_scroll,
        visible_lines,
        max_scroll,
    );
}

fn help_popup_area(area: Rect) -> Rect {
    let popup_width = area.width.saturating_sub(4).min(112).max(48);
    let popup_height = area.height.saturating_sub(2).min(32).max(12);
    let popup_x = area.x + area.width.saturating_sub(popup_width) / 2;
    let popup_y = area.y + area.height.saturating_sub(popup_height) / 2;

    Rect::new(popup_x, popup_y, popup_width, popup_height)
}

fn help_popup_inner_areas(inner: Rect) -> Option<(Rect, Rect, Rect)> {
    if inner.height < 4 || inner.width < 24 {
        return None;
    }

    let footer_height = 2;
    let content_area = Rect {
        x: inner.x.saturating_add(1),
        y: inner.y,
        width: inner.width.saturating_sub(2),
        height: inner.height.saturating_sub(footer_height),
    };
    let divider_area = Rect::new(
        inner.x,
        content_area.y + content_area.height,
        inner.width,
        1,
    );
    let footer_area = Rect::new(inner.x, divider_area.y + 1, inner.width, 1);

    Some((content_area, divider_area, footer_area))
}

fn render_help_footer(
    f: &mut Frame,
    area: Rect,
    scroll: usize,
    visible_lines: usize,
    max_scroll: usize,
) {
    let footer = help_footer_line(scroll, visible_lines, max_scroll);
    let footer_widget = Paragraph::new(footer)
        .style(theme::surface())
        .alignment(Alignment::Center);
    f.render_widget(footer_widget, area);
}

fn help_overlay_lines(width: u16) -> Vec<Line<'static>> {
    if width >= 82 {
        two_column_help_lines(width)
    } else {
        one_column_help_lines(width)
    }
}

fn two_column_help_lines(width: u16) -> Vec<Line<'static>> {
    let gap = 4;
    let column_width = width.saturating_sub(gap) / 2;
    let left = help_column_lines(LEFT_HELP_SECTIONS, column_width);
    let right = help_column_lines(RIGHT_HELP_SECTIONS, column_width);
    let row_count = left.len().max(right.len());
    let mut rows = Vec::with_capacity(row_count);

    for index in 0..row_count {
        let mut line = left.get(index).cloned().unwrap_or_else(Line::default);
        pad_line(&mut line, column_width);
        line.spans.push(Span::raw(" ".repeat(gap as usize)));
        if let Some(right_line) = right.get(index) {
            line.spans.extend(right_line.spans.clone());
        }
        rows.push(line);
    }

    rows
}

fn one_column_help_lines(width: u16) -> Vec<Line<'static>> {
    let mut lines = help_column_lines(LEFT_HELP_SECTIONS, width);
    lines.push(Line::default());
    lines.extend(help_column_lines(RIGHT_HELP_SECTIONS, width));
    lines
}

fn help_column_lines(sections: &[HelpSection], width: u16) -> Vec<Line<'static>> {
    let mut lines = Vec::new();

    for (index, section) in sections.iter().enumerate() {
        if index > 0 {
            lines.push(Line::default());
        }
        lines.push(section_title_line(section.title));
        for entry in section.entries {
            lines.extend(help_entry_lines(entry, width));
        }
    }

    lines
}

fn section_title_line(title: &'static str) -> Line<'static> {
    Line::from(Span::styled(
        title,
        Style::default()
            .fg(theme::WARNING)
            .add_modifier(Modifier::BOLD),
    ))
}

fn help_entry_lines(entry: &HelpEntry, width: u16) -> Vec<Line<'static>> {
    let prefix = help_entry_prefix(entry.keys);
    let prefix_width = spans_display_width(&prefix);
    let description_width = width.saturating_sub(prefix_width + HELP_ENTRY_GAP).max(1);
    let mut chunks = wrap_text(entry.description, description_width);

    if chunks.is_empty() {
        return vec![Line::from(prefix)];
    }

    let first = chunks.remove(0);
    let mut lines = vec![line_with_right_aligned_description(prefix, first, width)];
    for chunk in chunks {
        lines.push(right_aligned_description_line(chunk, width));
    }

    lines
}

fn help_entry_prefix(keys: &str) -> Vec<Span<'static>> {
    let mut spans = vec![Span::raw(" ".repeat(HELP_ENTRY_INDENT as usize))];

    for (index, chip) in key_chips(keys).into_iter().enumerate() {
        if index > 0 {
            spans.push(Span::styled("/", Style::default().fg(theme::TEXT_DISABLED)));
        }
        spans.extend(key_chip_spans(chip));
    }

    spans
}

fn key_chip_spans(label: String) -> Vec<Span<'static>> {
    vec![Span::styled(
        format!(" {label} "),
        Style::default()
            .bg(theme::SURFACE_MUTED)
            .fg(theme::ACCENT)
            .add_modifier(Modifier::BOLD),
    )]
}

fn key_chips(keys: &str) -> Vec<String> {
    keys.split(" / ")
        .flat_map(|group| {
            let group = group.trim();
            if should_split_shortcut_group(group) {
                group.split_whitespace().map(str::to_string).collect()
            } else {
                vec![group.to_string()]
            }
        })
        .collect()
}

fn spans_display_width(spans: &[Span<'_>]) -> u16 {
    spans.iter().map(|span| display_width(&span.content)).sum()
}

fn line_with_right_aligned_description(
    mut spans: Vec<Span<'static>>,
    description: String,
    width: u16,
) -> Line<'static> {
    let prefix_width = spans_display_width(&spans);
    let description_width = display_width(&description);
    let gap = width.saturating_sub(prefix_width + description_width);

    spans.push(Span::raw(" ".repeat(gap as usize)));
    spans.push(description_span(description));

    Line::from(spans)
}

fn right_aligned_description_line(description: String, width: u16) -> Line<'static> {
    let description_width = display_width(&description);
    let gap = width.saturating_sub(description_width);

    Line::from(vec![
        Span::raw(" ".repeat(gap as usize)),
        description_span(description),
    ])
}

fn description_span(description: String) -> Span<'static> {
    Span::styled(description, Style::default().fg(theme::TEXT_SECONDARY))
}

fn wrap_text(text: &str, width: u16) -> Vec<String> {
    if width == 0 {
        return Vec::new();
    }

    let mut lines = Vec::new();
    let mut current = String::new();

    for word in text.split_whitespace() {
        append_wrapped_word(&mut lines, &mut current, word, width);
    }

    if !current.is_empty() {
        lines.push(current);
    }

    lines
}

fn append_wrapped_word(lines: &mut Vec<String>, current: &mut String, word: &str, width: u16) {
    let word_width = display_width(word);
    let current_width = display_width(current);

    if current.is_empty() && word_width <= width {
        current.push_str(word);
    } else if !current.is_empty() && current_width + 1 + word_width <= width {
        current.push(' ');
        current.push_str(word);
    } else {
        if !current.is_empty() {
            lines.push(std::mem::take(current));
        }
        append_word_chunks(lines, current, word, width);
    }
}

fn append_word_chunks(lines: &mut Vec<String>, current: &mut String, word: &str, width: u16) {
    if display_width(word) <= width {
        current.push_str(word);
        return;
    }

    for chunk in split_word_to_width(word, width) {
        if current.is_empty() {
            current.push_str(&chunk);
        } else {
            lines.push(std::mem::take(current));
            current.push_str(&chunk);
        }
    }
}

fn split_word_to_width(word: &str, width: u16) -> Vec<String> {
    let mut chunks = Vec::new();
    let mut current = String::new();
    let mut used = 0;

    for ch in word.chars() {
        let char_width = if ch.is_ascii() { 1 } else { 2 };
        if used + char_width > width && !current.is_empty() {
            chunks.push(std::mem::take(&mut current));
            used = 0;
        }
        current.push(ch);
        used += char_width;
    }

    if !current.is_empty() {
        chunks.push(current);
    }

    chunks
}

fn should_split_shortcut_group(group: &str) -> bool {
    let parts: Vec<&str> = group.split_whitespace().collect();

    parts.len() > 1
        && parts.iter().all(|part| {
            part.chars().count() == 1 && part.chars().all(|ch| ch.is_ascii_alphabetic())
        })
}

fn pad_line(line: &mut Line<'static>, width: u16) {
    let line_width = line_display_width(line);
    if line_width < width {
        line.spans
            .push(Span::raw(" ".repeat((width - line_width) as usize)));
    }
}

fn help_footer_line(scroll: usize, visible_lines: usize, max_scroll: usize) -> Line<'static> {
    let total_pages = if max_scroll == 0 {
        1
    } else {
        (max_scroll + visible_lines) / visible_lines
    };
    let current_page = (scroll / visible_lines).saturating_add(1).min(total_pages);

    Line::from(vec![
        Span::styled("Press ESC or q to close", Style::default().fg(theme::TEXT)),
        Span::styled(
            "  |  j/k scroll  |  ",
            Style::default().fg(theme::TEXT_SECONDARY),
        ),
        Span::styled(
            format!("Page {current_page}/{total_pages}"),
            Style::default().fg(theme::ACCENT),
        ),
    ])
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

    let brand_content = " EXCEL-CLI ";
    let title_content = format!(" {file_name} ");

    let brand_width = display_width(brand_content);
    let title_width = display_width(&title_content);
    let max_title_width = (area.width / 3).min(title_width);

    let mut tab_widths = Vec::new();
    let mut total_width = 0;
    let mut visible_tabs = Vec::new();

    let horizontal_layout = Layout::default()
        .direction(Direction::Horizontal)
        .constraints([
            Constraint::Length(brand_width),
            Constraint::Length(max_title_width),
            Constraint::Min(0),
        ])
        .split(area);

    let title_style = if is_editing {
        Style::default().bg(Color::Black).fg(theme::TEXT_DISABLED)
    } else {
        Style::default().bg(Color::Black).fg(theme::TEXT_SECONDARY)
    };
    let brand_style = Style::default()
        .bg(Color::Black)
        .fg(theme::ACCENT)
        .add_modifier(Modifier::BOLD);

    let brand_widget = Paragraph::new(brand_content).style(brand_style);
    let title_widget = Paragraph::new(title_content).style(title_style);

    f.render_widget(brand_widget, horizontal_layout[0]);
    f.render_widget(title_widget, horizontal_layout[1]);

    let tabs_area = horizontal_layout[2];
    let rows_cols = sheet_rows_cols(app_state);
    let rows_cols_plain = format!("Rows/Cols: {rows_cols}");
    let base_rows_width = display_width(&rows_cols_plain);
    let total_tab_width: u16 = sheet_names.iter().map(|name| display_width(name)).sum();
    let visible_tabs_width = tabs_area.width.saturating_sub(base_rows_width);
    let tabs_overflow = total_tab_width > visible_tabs_width;
    let rows_cols_plain = if tabs_overflow {
        format!("... {rows_cols_plain}")
    } else {
        rows_cols_plain
    };
    let rows_cols_width = display_width(&rows_cols_plain);
    let available_width = tabs_area.width as usize;

    for (i, name) in sheet_names.iter().enumerate() {
        let tab_width = display_width(name) as usize;

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

    // Create constraints for tab layout
    let mut tab_constraints = Vec::new();
    for &width in &tab_widths {
        tab_constraints.push(Constraint::Length(width));
    }
    tab_constraints.push(Constraint::Min(0)); // Filler space

    let tab_layout = Layout::default()
        .direction(Direction::Horizontal)
        .constraints(tab_constraints)
        .split(tabs_area);

    // Render each visible tab
    for (layout_idx, &sheet_idx) in visible_tabs.iter().enumerate() {
        if layout_idx >= tab_layout.len() - 1 {
            break;
        }

        let name = &sheet_names[sheet_idx];
        let is_current = sheet_idx == current_index;

        let style = if is_editing {
            if is_current {
                Style::default().bg(Color::Black).fg(theme::TEXT_DISABLED)
            } else {
                Style::default().bg(Color::Black).fg(theme::TEXT_DISABLED)
            }
        } else if is_current {
            Style::default()
                .bg(Color::Black)
                .fg(theme::ACCENT)
                .add_modifier(Modifier::BOLD)
        } else {
            Style::default().bg(Color::Black).fg(theme::TEXT_SECONDARY)
        };

        let tab_widget = Paragraph::new(name.to_string())
            .style(style)
            .alignment(ratatui::layout::Alignment::Center);

        f.render_widget(tab_widget, tab_layout[layout_idx]);
    }

    let rows_cols_rect = Rect {
        x: tabs_area.x
            + tabs_area
                .width
                .saturating_sub(rows_cols_width.min(tabs_area.width)),
        y: tabs_area.y,
        width: rows_cols_width.min(tabs_area.width),
        height: 1,
    };
    let mut rows_cols_spans = Vec::new();
    if tabs_overflow {
        rows_cols_spans.push(Span::styled(
            "... ",
            Style::default().bg(Color::Black).fg(theme::TEXT_SECONDARY),
        ));
    }
    rows_cols_spans.push(Span::styled(
        "Rows/Cols: ",
        Style::default().bg(Color::Black).fg(theme::TEXT_SECONDARY),
    ));
    rows_cols_spans.push(Span::styled(
        rows_cols,
        Style::default().bg(Color::Black).fg(theme::ACCENT),
    ));

    let rows_cols_widget = Paragraph::new(Line::from(rows_cols_spans))
        .style(Style::default().bg(Color::Black))
        .alignment(ratatui::layout::Alignment::Right);
    f.render_widget(rows_cols_widget, rows_cols_rect);
}

#[cfg(test)]
mod tests {
    use ratatui::{backend::TestBackend, style::Color, Terminal};
    use std::path::PathBuf;

    use super::{theme, ui};
    use crate::app::{AppState, HelpEntry, InputMode};
    use crate::excel::{Cell, Sheet, Workbook};

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
        };
        let app = AppState::new(
            Workbook::from_sheets_for_test(vec![sheet]),
            PathBuf::from("scores.xlsx"),
        )
        .unwrap();
        app
    }

    fn app_with_many_sheets() -> AppState<'static> {
        let make_sheet = |name: &str| Sheet {
            name: name.to_string(),
            data: vec![vec![Cell::empty(); 2]; 2],
            max_rows: 1,
            max_cols: 1,
            is_loaded: true,
        };

        AppState::new(
            Workbook::from_sheets_for_test(vec![
                make_sheet("Alpha"),
                make_sheet("Beta"),
                make_sheet("Gamma"),
                make_sheet("Delta"),
                make_sheet("Epsilon"),
                make_sheet("Zeta"),
            ]),
            PathBuf::from("many.xlsx"),
        )
        .unwrap()
    }

    fn rendered_lines(terminal: &Terminal<TestBackend>) -> Vec<String> {
        let buffer = terminal.backend().buffer();
        let width = buffer.area.width as usize;

        buffer
            .content
            .chunks(width)
            .map(|row| row.iter().map(|cell| cell.symbol.as_str()).collect())
            .collect()
    }

    fn text_fg_at(terminal: &Terminal<TestBackend>, needle: &str) -> Color {
        let lines = rendered_lines(terminal);
        let row = line_index(&lines, needle);
        let col = lines[row]
            .find(needle)
            .unwrap_or_else(|| panic!("expected rendered output to contain {needle}"));
        let offset = needle
            .chars()
            .position(|ch| !ch.is_whitespace())
            .unwrap_or(0);
        let buffer = terminal.backend().buffer();
        let width = buffer.area.width as usize;
        buffer.content[row * width + col + offset].fg
    }

    fn fg_before_needle(terminal: &Terminal<TestBackend>, needle: &str) -> Color {
        let lines = rendered_lines(terminal);
        let row = line_index(&lines, needle);
        let col = lines[row]
            .find(needle)
            .unwrap_or_else(|| panic!("expected rendered output to contain {needle}"));
        let buffer = terminal.backend().buffer();
        let width = buffer.area.width as usize;
        buffer.content[row * width + col.saturating_sub(1)].fg
    }

    fn fg_at(terminal: &Terminal<TestBackend>, row: usize, col: usize) -> Color {
        let buffer = terminal.backend().buffer();
        let width = buffer.area.width as usize;
        buffer.content[row * width + col].fg
    }

    fn bg_at(terminal: &Terminal<TestBackend>, row: usize, col: usize) -> Color {
        let buffer = terminal.backend().buffer();
        let width = buffer.area.width as usize;
        buffer.content[row * width + col].bg
    }

    fn symbol_at(terminal: &Terminal<TestBackend>, row: usize, col: usize) -> String {
        let buffer = terminal.backend().buffer();
        let width = buffer.area.width as usize;
        buffer.content[row * width + col].symbol.clone()
    }

    fn line_index(lines: &[String], needle: &str) -> usize {
        lines
            .iter()
            .position(|line| line.contains(needle))
            .unwrap_or_else(|| panic!("expected rendered output to contain {needle}"))
    }

    fn help_overlay_text(width: u16) -> String {
        super::help_overlay_lines(width)
            .iter()
            .map(|line| {
                line.spans
                    .iter()
                    .map(|span| span.content.as_ref())
                    .collect::<String>()
            })
            .collect::<Vec<_>>()
            .join("\n")
    }

    #[test]
    fn renders_help_overlay_as_structured_command_reference() {
        let backend = TestBackend::new(140, 40);
        let mut terminal = Terminal::new(backend).unwrap();
        let mut app = app_with_sheet();
        app.show_help();

        terminal.draw(|frame| ui(frame, &mut app)).unwrap();

        let rendered = rendered_lines(&terminal).join("\n");

        assert!(matches!(app.input_mode, InputMode::Help));
        assert!(rendered.contains("COMMAND HELP"));
        assert!(rendered.contains("NAVIGATION"));
        assert!(rendered.contains("ACTIONS"));
        assert!(rendered.contains("SEARCH"));
        assert!(rendered.contains("FILE & APP"));
        assert!(rendered.contains("JUMP & SHEETS"));
        assert!(rendered.contains("Press ESC or q to close"));
        assert!(rendered.contains("Page "));
        assert!(!rendered.contains("preview"));
        assert!(!rendered.contains("findings"));
    }

    #[test]
    fn help_overlay_uses_solid_backdrop_to_hide_underlying_sheet() {
        let backend = TestBackend::new(140, 40);
        let mut terminal = Terminal::new(backend).unwrap();
        let mut app = app_with_sheet();
        app.show_help();

        terminal.draw(|frame| ui(frame, &mut app)).unwrap();

        assert_eq!(symbol_at(&terminal, 0, 0), " ");
        assert_eq!(bg_at(&terminal, 0, 0), theme::BACKGROUND);
    }

    #[test]
    fn help_entries_render_grouped_shortcuts_as_individual_chips() {
        let entry = HelpEntry {
            keys: "h j k l / arrows",
            description: "Move cell",
        };

        let line_text = super::help_entry_lines(&entry, 60)[0]
            .spans
            .iter()
            .map(|span| span.content.as_ref())
            .collect::<String>();

        assert!(line_text.contains(" h "));
        assert!(line_text.contains(" j "));
        assert!(line_text.contains(" k "));
        assert!(line_text.contains(" l "));
        assert!(line_text.contains(" arrows "));
        assert!(line_text.contains(" h / j / k / l / arrows "));
        assert!(!line_text.contains("  /  "));
        assert!(!line_text.contains(""));
        assert!(!line_text.contains(""));
        assert!(!line_text.contains("‹"));
        assert!(!line_text.contains("›"));
        assert!(!line_text.contains(" h j k l "));
    }

    #[test]
    fn help_entry_separator_slashes_are_dimmed() {
        let spans = super::help_entry_prefix("h / j");

        assert!(spans.iter().any(
            |span| span.content.as_ref() == "/" && span.style.fg == Some(theme::TEXT_DISABLED)
        ));
    }

    #[test]
    fn help_entry_chips_use_compact_square_background_without_caps() {
        let spans = super::help_entry_prefix("h");
        let text = spans
            .iter()
            .map(|span| span.content.as_ref())
            .collect::<String>();

        assert!(text.contains(" h "));
        assert!(!text.contains(""));
        assert!(!text.contains(""));
        assert!(spans
            .iter()
            .any(|span| span.content.as_ref() == " h "
                && span.style.bg == Some(theme::SURFACE_MUTED)));
    }

    #[test]
    fn help_entry_descriptions_align_to_the_right_edge() {
        let short_entry = HelpEntry {
            keys: "h",
            description: "Move cell",
        };
        let long_entry = HelpEntry {
            keys: "Ctrl+arrows",
            description: "Jump to next non-empty cell",
        };

        let short_line = super::help_entry_lines(&short_entry, 42)[0]
            .spans
            .iter()
            .map(|span| span.content.as_ref())
            .collect::<String>();
        let long_line = super::help_entry_lines(&long_entry, 42)[0]
            .spans
            .iter()
            .map(|span| span.content.as_ref())
            .collect::<String>();

        assert_eq!(super::display_width(&short_line), 42);
        assert_eq!(super::display_width(&long_line), 42);
        assert!(short_line.ends_with("Move cell"));
        assert!(long_line.ends_with("Jump to next non-empty"));
    }

    #[test]
    fn help_entry_keeps_description_on_first_line_for_long_shortcut_groups() {
        let entry = HelpEntry {
            keys: ":noh / :nohlsearch",
            description: "Disable search highlighting",
        };

        let rendered = super::help_entry_lines(&entry, 44)
            .iter()
            .map(|line| {
                line.spans
                    .iter()
                    .map(|span| span.content.as_ref())
                    .collect::<String>()
            })
            .collect::<Vec<_>>();

        assert!(rendered[0].contains(":noh"));
        assert!(rendered[0].contains(":nohlsearch"));
        assert!(rendered[0].contains("Disable search"));
        assert_eq!(super::display_width(&rendered[0]), 44);
        assert!(rendered[0].ends_with("Disable search"));
        assert_eq!(super::display_width(&rendered[1]), 44);
        assert!(rendered[1].ends_with("highlighting"));
    }

    #[test]
    fn help_entry_descriptions_wrap_right_aligned_inside_column_width() {
        let entry = HelpEntry {
            keys: ":sheet <name|index>",
            description: "Switch sheet by exact name or one based index",
        };

        let lines = super::help_entry_lines(&entry, 34);
        let rendered = lines
            .iter()
            .map(|line| {
                line.spans
                    .iter()
                    .map(|span| span.content.as_ref())
                    .collect::<String>()
            })
            .collect::<Vec<_>>();

        let normalized = rendered
            .join(" ")
            .split_whitespace()
            .collect::<Vec<_>>()
            .join(" ");

        assert!(rendered.len() > 1);
        assert!(rendered.iter().all(|line| super::display_width(line) <= 34));
        assert!(normalized.contains("one based index"));
        assert!(rendered.iter().all(|line| {
            super::display_width(line) == 34 || !line.contains(|ch: char| ch.is_alphabetic())
        }));
    }

    #[test]
    fn help_overlay_model_lists_complete_command_reference() {
        let help_text = help_overlay_text(112);

        for required in [
            ":cw fit all",
            ":dr <start> <end>",
            ":dc <start> <end>",
            ":ej <h|v> <rows>",
            ":eja <h|v> <rows>",
            "EDIT MODE",
            "HELP CONTROLS",
        ] {
            assert!(
                help_text.contains(required),
                "expected help overlay to contain {required}"
            );
        }

        assert!(!help_text.contains("preview"));
        assert!(!help_text.contains("findings"));
    }

    #[test]
    fn renders_help_overlay_later_command_sections_when_scrolled() {
        let backend = TestBackend::new(120, 24);
        let mut terminal = Terminal::new(backend).unwrap();
        let mut app = app_with_sheet();
        app.show_help();
        app.help_scroll = 17;

        terminal.draw(|frame| ui(frame, &mut app)).unwrap();

        let mid_page = rendered_lines(&terminal).join("\n");

        assert!(mid_page.contains("ROWS & COLUMNS"));
        assert!(mid_page.contains(":cw fit all"));
        assert!(mid_page.contains(":dr <start> <end>"));
        assert!(mid_page.contains(":dc <start> <end>"));
    }

    #[test]
    fn renders_visual_refresh_shell_with_inspector_and_short_status() {
        let backend = TestBackend::new(140, 32);
        let mut terminal = Terminal::new(backend).unwrap();
        let mut app = app_with_sheet();

        terminal.draw(|frame| ui(frame, &mut app)).unwrap();

        let rendered = terminal
            .backend()
            .buffer()
            .content
            .iter()
            .map(|cell| cell.symbol.as_str())
            .collect::<String>();

        assert!(matches!(app.input_mode, InputMode::Normal));
        assert!(rendered.contains("EXCEL-CLI"));
        assert!(rendered.contains("Cell A1"));
        assert!(rendered.contains("NOTIFICATIONS"));
        assert!(rendered.contains("NORMAL"));
        assert!(rendered.contains("[:w] Save"));
        assert!(!rendered.contains("INSPECTOR"));
        assert!(!rendered.contains("Run Diagnostics"));
        assert!(!rendered.contains("Settings"));
        assert!(!rendered.contains("Execute Script"));
        assert!(!rendered.contains("Findings"));
        assert!(!rendered.contains("Columns"));
        assert!(!rendered.contains("Preview"));
    }

    #[test]
    fn renders_normal_mode_status_bar_as_single_row_on_wide_layout() {
        let backend = TestBackend::new(140, 32);
        let mut terminal = Terminal::new(backend).unwrap();
        let mut app = app_with_sheet();

        terminal.draw(|frame| ui(frame, &mut app)).unwrap();

        let lines = rendered_lines(&terminal);
        let status_row = &lines[lines.len() - 1];
        let title_row = &lines[0];

        assert!(status_row.contains(" NORMAL "));
        assert!(status_row.contains("[Enter] Edit"));
        assert!(status_row.contains("[/] Search"));
        assert!(status_row.contains("[:w] Save"));
        assert!(status_row.trim_end().ends_with("[:w] Save"));
        assert!(!status_row.contains("Rows/Cols"));
        assert!(!status_row.contains("Findings"));
        assert!(!status_row.contains("Columns"));
        assert!(!status_row.contains("Preview"));
        assert!(title_row.contains("Rows/Cols: 2 x 2"));
        assert!(title_row.trim_end().ends_with("Rows/Cols: 2 x 2"));
    }

    #[test]
    fn renders_cell_panel_above_notifications_in_vertical_info_layout() {
        let backend = TestBackend::new(140, 32);
        let mut terminal = Terminal::new(backend).unwrap();
        let mut app = app_with_sheet();

        terminal.draw(|frame| ui(frame, &mut app)).unwrap();

        let lines = rendered_lines(&terminal);
        let cell_row = line_index(&lines, "Cell A1");
        let notifications_row = line_index(&lines, " NOTIFICATIONS ");

        assert!(cell_row < notifications_row);
    }

    #[test]
    fn does_not_render_analysis_tabs_or_inspector_shell() {
        let backend = TestBackend::new(140, 32);
        let mut terminal = Terminal::new(backend).unwrap();
        let mut app = app_with_sheet();

        terminal.draw(|frame| ui(frame, &mut app)).unwrap();

        let lines = rendered_lines(&terminal);
        let full_text = lines.join("\n");

        assert!(!full_text.contains("INSPECTOR"));
        assert!(!full_text.contains("Analysis Panel"));
        assert!(!full_text.contains(" Details  Preview  Findings  Columns "));
        assert!(!full_text.contains("Query Preview"));
        assert!(!full_text.contains("FINDINGS"));
        assert!(!full_text.contains("COLUMNS PROFILE"));
    }

    #[test]
    fn renders_cell_details_with_dynamic_title_and_compact_fields() {
        let backend = TestBackend::new(140, 40);
        let mut terminal = Terminal::new(backend).unwrap();
        let mut app = app_with_sheet();
        app.selected_cell = (2, 2);

        terminal.draw(|frame| ui(frame, &mut app)).unwrap();

        let rendered = terminal
            .backend()
            .buffer()
            .content
            .iter()
            .map(|cell| cell.symbol.as_str())
            .collect::<String>();

        assert!(matches!(app.input_mode, InputMode::Normal));
        assert!(rendered.contains("Cell B2  Number  Len 2"));
        assert!(rendered.contains("10"));
        assert!(!rendered.contains("Type: Number"));
        assert!(!rendered.contains("Length: 2"));
        assert!(!rendered.contains("Content: 10"));
        assert!(rendered.contains("NOTIFICATIONS"));
        assert!(!rendered.contains("SHEET CONTEXT"));
        assert!(!rendered.contains("QUALITY"));
        assert!(!rendered.contains("No findings for active cell"));
    }

    #[test]
    fn renders_cell_panel_title_and_border_with_primary_text_color_by_default() {
        let backend = TestBackend::new(140, 40);
        let mut terminal = Terminal::new(backend).unwrap();
        let mut app = app_with_sheet();

        terminal.draw(|frame| ui(frame, &mut app)).unwrap();

        assert_eq!(
            text_fg_at(&terminal, " Cell A1  String  Len 4 "),
            theme::TEXT
        );
        assert_eq!(
            fg_before_needle(&terminal, " Cell A1  String  Len 4 "),
            theme::TEXT
        );
    }

    #[test]
    fn renders_notifications_panel_when_inspector_moves_below_table() {
        let backend = TestBackend::new(90, 28);
        let mut terminal = Terminal::new(backend).unwrap();
        let mut app = app_with_sheet();
        app.add_notification("Loaded 2 findings".to_string());

        terminal.draw(|frame| ui(frame, &mut app)).unwrap();

        let rendered = terminal
            .backend()
            .buffer()
            .content
            .iter()
            .map(|cell| cell.symbol.as_str())
            .collect::<String>();

        assert!(rendered.contains("Cell A1"));
        assert!(rendered.contains("Loaded 2 findings"));
        assert!(rendered.contains("NOTIFICATIONS"));
    }

    #[test]
    fn renders_notifications_title_and_border_with_primary_text_color() {
        let backend = TestBackend::new(140, 40);
        let mut terminal = Terminal::new(backend).unwrap();
        let mut app = app_with_sheet();
        app.add_notification("Loaded 2 findings".to_string());

        terminal.draw(|frame| ui(frame, &mut app)).unwrap();

        assert_eq!(text_fg_at(&terminal, " NOTIFICATIONS "), theme::TEXT);
        assert_eq!(fg_before_needle(&terminal, " NOTIFICATIONS "), theme::TEXT);
    }

    #[test]
    fn renders_editing_panel_with_vim_mode_in_title_and_without_status_mode() {
        let backend = TestBackend::new(140, 40);
        let mut terminal = Terminal::new(backend).unwrap();
        let mut app = app_with_sheet();
        app.start_editing();

        terminal.draw(|frame| ui(frame, &mut app)).unwrap();

        let lines = rendered_lines(&terminal);
        let full_text = lines.join("\n");
        let status_row = &lines[lines.len() - 1];
        let title_row = &lines[0];

        assert!(full_text.contains("Editing Cell A1"));
        assert!(full_text.contains("NORMAL"));
        assert!(!full_text.contains("TARGET CELL"));
        assert!(!full_text.contains("INPUT BUFFER [EDITING]"));
        assert_eq!(
            fg_at(&terminal, line_index(&lines, " Editing Cell A1 "), 0),
            theme::ACCENT
        );
        assert_eq!(text_fg_at(&terminal, "NORMAL"), theme::SUCCESS);
        assert!(status_row.contains(" EDIT "));
        assert!(status_row.contains("[Enter] Save"));
        assert!(status_row.trim_end().ends_with("[v] Visual"));
        assert!(!status_row.contains("Rows/Cols"));
        assert!(!status_row.contains("Mode "));
        assert!(title_row.contains("Rows/Cols: 2 x 2"));
    }

    #[test]
    fn renders_latest_notification_brighter_than_history() {
        let backend = TestBackend::new(140, 40);
        let mut terminal = Terminal::new(backend).unwrap();
        let mut app = app_with_sheet();
        app.add_notification("older notification".to_string());
        app.add_notification("latest notification".to_string());

        terminal.draw(|frame| ui(frame, &mut app)).unwrap();

        assert_eq!(text_fg_at(&terminal, "latest notification"), theme::TEXT);
        assert_eq!(
            text_fg_at(&terminal, "older notification"),
            theme::TEXT_SECONDARY
        );
    }

    #[test]
    fn removed_analysis_modes_do_not_appear_in_rendered_ui() {
        let backend = TestBackend::new(140, 32);
        let mut terminal = Terminal::new(backend).unwrap();
        let mut app = app_with_sheet();

        terminal.draw(|frame| ui(frame, &mut app)).unwrap();

        let rendered = terminal
            .backend()
            .buffer()
            .content
            .iter()
            .map(|cell| cell.symbol.as_str())
            .collect::<String>();

        assert!(rendered.contains("Cell A1"));
        assert!(rendered.contains("NOTIFICATIONS"));
        assert!(!rendered.contains("Findings"));
        assert!(!rendered.contains("Preview"));
        assert!(!rendered.contains("Columns"));
        assert!(!rendered.contains("COLUMNS PROFILE"));
        assert!(!rendered.contains("SHEET PROFILE"));
    }

    #[test]
    fn renders_rows_cols_in_top_right_with_overflow_hint_when_tabs_exceed_space() {
        let backend = TestBackend::new(60, 20);
        let mut terminal = Terminal::new(backend).unwrap();
        let mut app = app_with_many_sheets();

        terminal.draw(|frame| ui(frame, &mut app)).unwrap();

        let lines = rendered_lines(&terminal);
        let title_row = &lines[0];

        assert!(title_row.contains("Rows/Cols: 1 x 1"));
        assert!(title_row.trim_end().ends_with("... Rows/Cols: 1 x 1"));
        assert!(title_row.contains("Alpha"));
        assert!(!title_row.contains("Zeta"));
    }
}
