use serde_json::Value;
use std::path::Path;

use crate::cli::error::AppError;
use crate::cli::sheet_query::SheetBounds;
use crate::excel::{Sheet, Workbook};
use crate::utils::index_to_col_name;

pub(crate) fn file_format(path: &Path) -> String {
    path.extension()
        .and_then(|extension| extension.to_str())
        .map(str::to_lowercase)
        .unwrap_or_else(|| "unknown".to_string())
}

pub(crate) fn format_range(
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
) -> String {
    format!(
        "{}{}:{}{}",
        index_to_col_name(start_col),
        start_row,
        index_to_col_name(end_col),
        end_row
    )
}

pub(crate) fn format_bounds(bounds: SheetBounds) -> String {
    format_range(
        bounds.start_row,
        bounds.start_col,
        bounds.end_row,
        bounds.end_col,
    )
}

pub(crate) fn value_text(value: &Value) -> String {
    match value {
        Value::Null => String::new(),
        Value::String(value) => value.clone(),
        other => other.to_string(),
    }
}

pub(crate) fn tab_separated_values(values: &[Value]) -> String {
    values.iter().map(value_text).collect::<Vec<_>>().join("\t")
}

pub(crate) fn sheet_by_index<'a>(
    workbook: &'a Workbook,
    sheet_index: usize,
    sheet_name: &str,
) -> Result<&'a Sheet, AppError> {
    workbook
        .get_sheet_by_index(sheet_index)
        .ok_or_else(|| AppError::TargetNotFound {
            message: format!("Sheet '{}' not found", sheet_name),
        })
}
