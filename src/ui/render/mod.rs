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
    widgets::{Block, Borders, Clear, Paragraph},
    Frame, Terminal,
};
use std::{io, time::Duration};

mod help_overlay;
mod spreadsheet;
mod status;

use help_overlay::draw_help_popup;
use spreadsheet::{draw_spreadsheet, draw_title_with_tabs, update_visible_area};
use status::{draw_status_bar, status_bar_height};

#[cfg(test)]
use help_overlay::{help_entry_lines, help_overlay_lines};

use crate::app::AppState;
use crate::app::InputMode;
use crate::app::VimMode;
use crate::ui::handlers::handle_key_event;
use crate::ui::theme;
use crate::utils::cell_reference;

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

pub(super) fn display_width(text: &str) -> u16 {
    text.chars()
        .fold(0, |acc, ch| acc + if ch.is_ascii() { 1 } else { 2 })
}

pub(super) fn line_display_width(line: &Line<'_>) -> u16 {
    line.spans
        .iter()
        .map(|span| display_width(&span.content))
        .sum()
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

#[cfg(test)]
mod tests;
