use ratatui::{
    layout::{Constraint, Direction, Layout, Rect},
    style::{Color, Modifier, Style},
    text::{Line, Span},
    widgets::Paragraph,
    Frame,
};

use crate::app::{AppState, InputMode};
use crate::ui::theme;

use super::line_display_width;

pub(super) fn status_bar_height(app_state: &AppState, width: u16) -> u16 {
    let _ = width;
    if matches!(app_state.input_mode, InputMode::Help) {
        0
    } else {
        1
    }
}

pub(super) fn draw_status_bar(f: &mut Frame, app_state: &AppState, area: Rect) {
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
