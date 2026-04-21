use crate::app::{AppState, InputMode};
use crate::utils::{cell_reference, index_to_col_name};

const SAMPLE_ROWS: usize = 6;
const SAMPLE_COLS: usize = 6;
const MAX_CELL_DISPLAY_CHARS: usize = 24;

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct QueryPreview {
    pub file_path: String,
    pub sheet_name: String,
    pub sheet_index: usize,
    pub selected_cell: String,
    pub used_range: String,
    pub selects: String,
    pub filters: String,
    pub columns: Vec<String>,
    pub rows: Vec<QueryPreviewRow>,
    pub truncated: bool,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct QueryPreviewRow {
    pub row_number: usize,
    pub values: Vec<String>,
}

impl QueryPreview {
    fn from_app(app: &AppState) -> Self {
        let sheet = app.workbook.get_current_sheet();
        let sheet_index = app.workbook.get_current_sheet_index();
        let selected_cell = cell_reference(app.selected_cell);
        let used_range = if sheet.max_rows == 0 || sheet.max_cols == 0 {
            "empty".to_string()
        } else {
            format!("A1:{}{}", index_to_col_name(sheet.max_cols), sheet.max_rows)
        };

        let sample_start_row = if sheet.max_rows == 0 {
            0
        } else {
            app.selected_cell.0.clamp(1, sheet.max_rows)
        };
        let sample_start_col = if sheet.max_cols == 0 { 0 } else { 1 };
        let sample_end_row = (sample_start_row + SAMPLE_ROWS.saturating_sub(1)).min(sheet.max_rows);
        let sample_end_col = (sample_start_col + SAMPLE_COLS.saturating_sub(1)).min(sheet.max_cols);

        let columns = if sample_start_col == 0 {
            Vec::new()
        } else {
            (sample_start_col..=sample_end_col)
                .map(index_to_col_name)
                .collect()
        };

        let rows = if sample_start_row == 0 || sample_start_col == 0 {
            Vec::new()
        } else {
            (sample_start_row..=sample_end_row)
                .map(|row| QueryPreviewRow {
                    row_number: row,
                    values: (sample_start_col..=sample_end_col)
                        .map(|col| {
                            sheet
                                .data
                                .get(row)
                                .and_then(|cells| cells.get(col))
                                .map(|cell| truncate_cell(&cell.value))
                                .unwrap_or_default()
                        })
                        .collect(),
                })
                .collect()
        };

        let truncated = sample_start_row > 1
            || sample_start_col > 1
            || sample_end_row < sheet.max_rows
            || sample_end_col < sheet.max_cols;

        Self {
            file_path: app.file_path.to_string_lossy().to_string(),
            sheet_name: sheet.name.clone(),
            sheet_index: sheet_index + 1,
            selected_cell,
            used_range,
            selects: "all columns".to_string(),
            filters: "none".to_string(),
            columns,
            rows,
            truncated,
        }
    }
}

impl AppState<'_> {
    pub fn show_query_preview(&mut self) {
        let sheet_index = self.workbook.get_current_sheet_index();
        let sheet_name = self.workbook.get_current_sheet_name();

        if self.workbook.is_lazy_loading() && !self.workbook.is_sheet_loaded(sheet_index) {
            if let Err(e) = self.workbook.ensure_sheet_loaded(sheet_index, &sheet_name) {
                self.add_notification(format!("Preview failed: {e}"));
                return;
            }
        }

        self.query_preview = Some(QueryPreview::from_app(self));
        self.input_mode = InputMode::Preview;
    }

    pub fn close_query_preview(&mut self) {
        self.query_preview = None;
        self.input_mode = InputMode::Normal;
    }
}

fn truncate_cell(value: &str) -> String {
    let mut result = String::new();
    for (idx, ch) in value.chars().enumerate() {
        if idx >= MAX_CELL_DISPLAY_CHARS {
            result.push_str("...");
            return result;
        }
        result.push(ch);
    }
    result
}

#[cfg(test)]
mod tests {
    use std::path::PathBuf;

    use crate::app::AppState;
    use crate::excel::{Cell, Sheet, Workbook};

    fn sheet_with_values(name: &str, values: &[&[&str]]) -> Sheet {
        let max_rows = values.len();
        let max_cols = values.iter().map(|row| row.len()).max().unwrap_or(0);
        let mut data = vec![vec![Cell::empty(); max_cols + 1]; max_rows + 1];

        for (row_idx, row) in values.iter().enumerate() {
            for (col_idx, value) in row.iter().enumerate() {
                data[row_idx + 1][col_idx + 1] = Cell::new((*value).to_string(), false);
            }
        }

        Sheet {
            name: name.to_string(),
            data,
            max_rows,
            max_cols,
            is_loaded: true,
        }
    }

    #[test]
    fn preview_snapshots_current_target_and_capped_sample() {
        let workbook = Workbook::from_sheets_for_test(vec![sheet_with_values(
            "Data",
            &[
                &[
                    "Name", "Region", "Sales", "Owner", "Quarter", "Status", "Notes",
                ],
                &["Ada", "West", "10", "Mina", "Q1", "Open", "A"],
                &["Ben", "East", "12", "Noor", "Q1", "Won", "B"],
                &["Cid", "North", "9", "Ira", "Q2", "Open", "C"],
                &["Dee", "South", "7", "Ola", "Q2", "Lost", "D"],
                &["Eli", "West", "8", "Paz", "Q3", "Open", "E"],
                &["Fay", "East", "11", "Uma", "Q3", "Won", "F"],
            ],
        )]);
        let mut app = AppState::new(workbook, PathBuf::from("/tmp/report.xlsx")).unwrap();
        app.selected_cell = (2, 2);

        app.show_query_preview();

        let preview = app.query_preview.as_ref().expect("preview should be set");
        assert_eq!(preview.file_path, "/tmp/report.xlsx");
        assert_eq!(preview.sheet_name, "Data");
        assert_eq!(preview.sheet_index, 1);
        assert_eq!(preview.selected_cell, "B2");
        assert_eq!(preview.used_range, "A1:G7");
        assert_eq!(preview.selects, "all columns");
        assert_eq!(preview.filters, "none");
        assert_eq!(preview.columns, vec!["A", "B", "C", "D", "E", "F"]);
        assert_eq!(preview.rows.len(), 6);
        assert_eq!(preview.rows[0].row_number, 2);
        assert_eq!(
            preview.rows[0].values,
            vec!["Ada", "West", "10", "Mina", "Q1", "Open"]
        );
        assert!(preview.truncated);
    }
}
