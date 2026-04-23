use ratatui::{
    layout::{Constraint, Direction, Layout, Rect},
    style::{Color, Modifier, Style},
    text::{Line, Span},
    widgets::{Block, Borders, Cell, Paragraph, Row, Table},
    Frame,
};

use crate::app::{AppState, InputMode};
use crate::ui::theme;
use crate::utils::index_to_col_name;

use super::display_width;

/// Update the visible area of the spreadsheet based on the available space
pub(super) fn update_visible_area(app_state: &mut AppState, area: Rect) {
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

pub(super) fn draw_spreadsheet(f: &mut Frame, app_state: &AppState, area: Rect) {
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

pub(super) fn draw_title_with_tabs(f: &mut Frame, app_state: &AppState, area: Rect) {
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
            Style::default().bg(Color::Black).fg(theme::TEXT_DISABLED)
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

fn sheet_rows_cols(app_state: &AppState) -> String {
    let sheet = app_state.workbook.get_current_sheet();
    format!("{} x {}", sheet.max_rows, sheet.max_cols)
}
