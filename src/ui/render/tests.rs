use ratatui::{backend::TestBackend, style::Color, Terminal};
use std::path::PathBuf;

use super::{theme, ui};
use crate::app::{AppState, HelpEntry, InputMode};
use crate::excel::{Cell, FreezePanes, Sheet, Workbook};

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
        freeze_panes: FreezePanes::none(),
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
        freeze_panes: FreezePanes::none(),
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

fn app_with_long_c22_cell() -> AppState<'static> {
    let mut data = vec![vec![Cell::empty(); 5]; 24];
    data[20][1] = Cell::new("分类甲".to_string(), false);
    data[20][2] = Cell::new("示例能源服务股份有限公司".to_string(), false);
    data[20][3] = Cell::new("Example Energy Services Company Limited".to_string(), false);
    data[20][4] = Cell::new(
        "示例省示例市示例区示例路100号示例大厦A座10层、20层、30层".to_string(),
        false,
    );
    data[21][1] = Cell::new("分类甲".to_string(), false);
    data[21][2] = Cell::new("示例一致服务集团股份有限公司".to_string(), false);
    data[21][3] = Cell::new(
        "Example Unified Services Corporation Limited".to_string(),
        false,
    );
    data[21][4] = Cell::new(
        "示例省示例市示例区样例四路15号示例服务大厦".to_string(),
        false,
    );
    data[22][1] = Cell::new("分类甲".to_string(), false);
    data[22][2] = Cell::new("示例跨区域资产服务集团股份有限公司".to_string(), false);
    data[22][3] = Cell::new(
        "Example International Research Operations and Holdings Company Limited".to_string(),
        false,
    );
    data[22][4] = Cell::new(
        "示例省示例市示例区样例南路示例广场45-48楼".to_string(),
        false,
    );

    let sheet = Sheet {
        name: "示例表".to_string(),
        data,
        max_rows: 23,
        max_cols: 4,
        is_loaded: true,
        freeze_panes: FreezePanes::none(),
    };

    AppState::new(
        Workbook::from_sheets_for_test(vec![sheet]),
        PathBuf::from("sample.xlsx"),
    )
    .unwrap()
}

fn app_with_frozen_grid() -> AppState<'static> {
    let mut data = vec![vec![Cell::empty(); 9]; 9];
    for (row_idx, row) in data.iter_mut().enumerate().take(9).skip(1) {
        for (col_idx, cell) in row.iter_mut().enumerate().take(9).skip(1) {
            *cell = Cell::new(format!("R{row_idx}C{col_idx}"), false);
        }
    }

    let sheet = Sheet {
        name: "Frozen".to_string(),
        data,
        max_rows: 8,
        max_cols: 8,
        is_loaded: true,
        freeze_panes: FreezePanes { rows: 1, cols: 1 },
    };

    AppState::new(
        Workbook::from_sheets_for_test(vec![sheet]),
        PathBuf::from("frozen.xlsx"),
    )
    .unwrap()
}

fn rendered_lines(terminal: &Terminal<TestBackend>) -> Vec<String> {
    let buffer = terminal.backend().buffer();
    let width = buffer.area.width as usize;

    buffer
        .content
        .chunks(width)
        .map(|row| row.iter().map(|cell| cell.symbol()).collect())
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

fn text_bg_at(terminal: &Terminal<TestBackend>, needle: &str) -> Color {
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
    buffer.content[row * width + col + offset].bg
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
    buffer.content[row * width + col].symbol().to_string()
}

#[test]
fn auto_fit_all_renders_full_long_cell_content() {
    let backend = TestBackend::new(100, 36);
    let mut terminal = Terminal::new(backend).unwrap();
    let mut app = app_with_long_c22_cell();
    let expected = "Example International Research Operations and Holdings Company Limited";

    terminal.draw(|frame| ui(frame, &mut app)).unwrap();
    app.input_buffer = "cw fit all".to_string();
    app.execute_command();
    app.input_buffer = "C22".to_string();
    app.execute_command();
    terminal.draw(|frame| ui(frame, &mut app)).unwrap();

    let lines = rendered_lines(&terminal);
    let row = lines
        .iter()
        .find(|line| line.contains("│22"))
        .unwrap_or_else(|| panic!("expected row 22 to render:\n{}", lines.join("\n")));

    assert!(row.contains(expected), "{row}");
}

#[test]
fn frozen_panes_keep_top_row_and_left_column_visible_while_scrolled() {
    let backend = TestBackend::new(100, 32);
    let mut terminal = Terminal::new(backend).unwrap();
    let mut app = app_with_frozen_grid();
    app.start_row = 6;
    app.start_col = 6;
    app.selected_cell = (6, 6);

    terminal.draw(|frame| ui(frame, &mut app)).unwrap();

    let rendered = rendered_lines(&terminal).join("\n");
    assert!(rendered.contains("Frozen: B2"), "{rendered}");
    assert!(rendered.contains("R1C1"), "{rendered}");
    assert!(rendered.contains("R1C6"), "{rendered}");
    assert!(rendered.contains("R6C1"), "{rendered}");
    assert!(rendered.contains("R6C6"), "{rendered}");
}

#[test]
fn frozen_panes_style_frozen_regions_while_scrolled() {
    let backend = TestBackend::new(100, 32);
    let mut terminal = Terminal::new(backend).unwrap();
    let mut app = app_with_frozen_grid();
    app.start_row = 6;
    app.start_col = 6;
    app.selected_cell = (8, 8);

    terminal.draw(|frame| ui(frame, &mut app)).unwrap();

    assert_eq!(text_bg_at(&terminal, "R1C1"), theme::FROZEN_BACKGROUND);
    assert_eq!(text_bg_at(&terminal, "R1C6"), theme::FROZEN_BACKGROUND);
    assert_eq!(text_bg_at(&terminal, "R6C1"), theme::FROZEN_BACKGROUND);
    assert_eq!(text_bg_at(&terminal, "R6C6"), theme::BACKGROUND);
}

#[test]
fn selected_and_search_styles_override_frozen_region_style() {
    let backend = TestBackend::new(100, 32);
    let mut terminal = Terminal::new(backend).unwrap();
    let mut app = app_with_frozen_grid();
    app.start_row = 6;
    app.start_col = 6;
    app.selected_cell = (1, 1);
    app.search_results.push((1, 6));

    terminal.draw(|frame| ui(frame, &mut app)).unwrap();

    assert_eq!(text_bg_at(&terminal, "R1C1"), Color::White);
    assert_eq!(text_bg_at(&terminal, "R1C6"), theme::SEARCH);
}

#[test]
fn auto_fit_all_does_not_shrink_visible_fitted_columns() {
    let backend = TestBackend::new(148, 59);
    let mut terminal = Terminal::new(backend).unwrap();
    let mut app = app_with_long_c22_cell();
    let expected = "Example International Research Operations and Holdings Company Limited";

    terminal.draw(|frame| ui(frame, &mut app)).unwrap();
    app.input_buffer = "cw fit all".to_string();
    app.execute_command();
    terminal.draw(|frame| ui(frame, &mut app)).unwrap();

    let lines = rendered_lines(&terminal);
    let row = lines
        .iter()
        .find(|line| line.contains("│22"))
        .unwrap_or_else(|| panic!("expected row 22 to render:\n{}", lines.join("\n")));

    assert!(row.contains(expected), "{row}");
}

#[test]
fn auto_fit_all_shows_partial_next_column_without_shrinking_fitted_columns() {
    let backend = TestBackend::new(148, 59);
    let mut terminal = Terminal::new(backend).unwrap();
    let mut app = app_with_long_c22_cell();
    let full_c_cell = "Example International Research Operations and Holdings Company Limited";
    let partial_d_cell = "示 例 省 示 例 市";

    terminal.draw(|frame| ui(frame, &mut app)).unwrap();
    app.input_buffer = "cw fit all".to_string();
    app.execute_command();
    terminal.draw(|frame| ui(frame, &mut app)).unwrap();

    let lines = rendered_lines(&terminal);
    let row = lines
        .iter()
        .find(|line| line.contains("│22"))
        .unwrap_or_else(|| panic!("expected row 22 to render:\n{}", lines.join("\n")));

    assert!(row.contains(full_c_cell), "{row}");
    assert!(row.contains(partial_d_cell), "{row}");
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
        .map(|cell| cell.symbol())
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
        .map(|cell| cell.symbol())
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
        .map(|cell| cell.symbol())
        .collect::<String>();

    assert!(rendered.contains("Cell A1"));
    assert!(rendered.contains("Loaded 2 findings"));
    assert!(rendered.contains("NOTIFICATIONS"));
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
        .map(|cell| cell.symbol())
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
