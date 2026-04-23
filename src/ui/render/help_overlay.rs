use ratatui::{
    layout::{Alignment, Rect},
    style::{Modifier, Style},
    text::{Line, Span},
    widgets::{Block, Borders, Clear, Paragraph},
    Frame,
};

use crate::app::{AppState, HelpEntry, HelpSection, LEFT_HELP_SECTIONS, RIGHT_HELP_SECTIONS};
use crate::ui::theme;

use super::{display_width, line_display_width};

const HELP_ENTRY_INDENT: u16 = 2;
const HELP_ENTRY_GAP: u16 = 1;

pub(super) fn draw_help_popup(f: &mut Frame, app_state: &mut AppState, area: Rect) {
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
    let popup_width = area.width.saturating_sub(4).clamp(48, 112);
    let popup_height = area.height.saturating_sub(2).clamp(12, 32);
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

pub(super) fn help_overlay_lines(width: u16) -> Vec<Line<'static>> {
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

pub(super) fn help_entry_lines(entry: &HelpEntry, width: u16) -> Vec<Line<'static>> {
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
