use anyhow::Context;
use serde_json::{json, Value};

use crate::cli::args::{resolve_sheet_target, InspectCommands};
use crate::cli::envelope;
use crate::cli::error::AppError;
use crate::excel::open_workbook;
use crate::utils::{index_to_col_name, parse_range};

pub fn handle(cmd: InspectCommands) -> Result<Value, AppError> {
    match cmd {
        InspectCommands::Workbook { file, format: _ } => inspect_workbook(file),
        InspectCommands::Sheet {
            file,
            sheet,
            sheet_index,
            format: _,
        } => inspect_sheet(file, sheet, sheet_index),
        InspectCommands::Sample {
            file,
            sheet,
            sheet_index,
            range,
            rows,
            header_row,
            format: _,
        } => inspect_sample(file, sheet, sheet_index, range, rows, header_row),
    }
}

fn file_format(path: &std::path::Path) -> String {
    path.extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_lowercase())
        .unwrap_or_else(|| "unknown".to_string())
}

fn inspect_workbook(file: std::path::PathBuf) -> Result<Value, AppError> {
    let format_str = file_format(&file);
    let path_str = file.to_string_lossy().to_string();

    let workbook = open_workbook(&file, false)
        .map_err(crate::cli::error::anyhow_to_app_error)?;

    let sheets: Vec<Value> = workbook
        .get_sheet_names()
        .iter()
        .enumerate()
        .map(|(index, name)| {
            let is_empty = if let Some(sheet) = workbook.get_sheet_by_index(index) {
                sheet.max_rows == 0 || sheet.max_cols == 0
            } else {
                true
            };
            json!({
                "name": name,
                "index": index,
                "is_empty": is_empty,
                "is_hidden_if_available": false,
            })
        })
        .collect();

    let data = json!({
        "sheet_count": sheets.len(),
        "sheets": sheets,
    });

    Ok(envelope::success_envelope(
        "inspect.workbook",
        &path_str,
        &format_str,
        envelope::target_workbook(),
        json!({}),
        data,
        vec![],
    ))
}

fn inspect_sheet(
    file: std::path::PathBuf,
    sheet: Option<String>,
    sheet_index: Option<usize>,
) -> Result<Value, AppError> {
    let format_str = file_format(&file);
    let path_str = file.to_string_lossy().to_string();

    let mut workbook = open_workbook(&file, false)
        .map_err(crate::cli::error::anyhow_to_app_error)?;

    let index = resolve_sheet_target(&workbook, &sheet, &sheet_index)?;
    let sheet_name = workbook.get_sheet_names()[index].clone();

    workbook
        .ensure_sheet_loaded(index, &sheet_name)
        .map_err(crate::cli::error::anyhow_to_app_error)?;

    let sheet_obj = workbook
        .get_sheet_by_index(index)
        .with_context(|| format!("Sheet '{}' not found", sheet_name))
        .map_err(crate::cli::error::anyhow_to_app_error)?;

    let used_range = workbook
        .get_used_range(index)
        .map_err(crate::cli::error::anyhow_to_app_error)?;

    let non_empty_rows = workbook
        .count_non_empty_rows(index)
        .map_err(crate::cli::error::anyhow_to_app_error)?;
    let non_empty_cols = workbook
        .count_non_empty_cols(index)
        .map_err(crate::cli::error::anyhow_to_app_error)?;

    let (header_candidates, recommended_header_row) = workbook
        .find_header_candidates(index)
        .map_err(crate::cli::error::anyhow_to_app_error)?;

    let data = json!({
        "name": sheet_obj.name,
        "index": index,
        "used_range": used_range,
        "max_rows": sheet_obj.max_rows,
        "max_cols": sheet_obj.max_cols,
        "non_empty_rows": non_empty_rows,
        "non_empty_cols": non_empty_cols,
        "recommended_header_row": recommended_header_row,
        "header_candidates": header_candidates,
    });

    Ok(envelope::success_envelope(
        "inspect.sheet",
        &path_str,
        &format_str,
        envelope::target_sheet(&sheet_name, index),
        json!({}),
        data,
        vec![],
    ))
}

fn inspect_sample(
    file: std::path::PathBuf,
    sheet: Option<String>,
    sheet_index: Option<usize>,
    range: Option<String>,
    rows: Option<usize>,
    header_row: String,
) -> Result<Value, AppError> {
    let format_str = file_format(&file);
    let path_str = file.to_string_lossy().to_string();

    let mut workbook = open_workbook(&file, false)
        .map_err(crate::cli::error::anyhow_to_app_error)?;

    let index = resolve_sheet_target(&workbook, &sheet, &sheet_index)?;
    let sheet_name = workbook.get_sheet_names()[index].clone();

    workbook
        .ensure_sheet_loaded(index, &sheet_name)
        .map_err(crate::cli::error::anyhow_to_app_error)?;

    let sheet_obj = workbook
        .get_sheet_by_index(index)
        .with_context(|| format!("Sheet '{}' not found", sheet_name))
        .map_err(crate::cli::error::anyhow_to_app_error)?;

    let used_range = workbook
        .get_used_range(index)
        .unwrap_or_default();

    // Determine the sample range
    let ((mut start_row, mut start_col), (mut end_row, mut end_col)) = if let Some(ref r) = range {
        parse_range(r)
            .ok_or_else(|| AppError::InvalidQuery {
                message: format!("Invalid range format: {}", r),
            })?
    } else if !used_range.is_empty() {
        parse_range(&used_range).unwrap_or(((1, 1), (1, 1)))
    } else {
        ((1, 1), (1, 1))
    };

    // Clamp to actual bounds
    let max_row = sheet_obj.max_rows.max(1);
    let max_col = sheet_obj.max_cols.max(1);
    start_row = start_row.min(max_row);
    start_col = start_col.min(max_col);
    end_row = end_row.min(max_row);
    end_col = end_col.min(max_col);
    if start_row > end_row {
        std::mem::swap(&mut start_row, &mut end_row);
    }
    if start_col > end_col {
        std::mem::swap(&mut start_col, &mut end_col);
    }

    // Apply row limit
    let row_limit = rows.unwrap_or(10);
    let sample_end_row = (start_row + row_limit.saturating_sub(1)).min(end_row);

    // Resolve header row
    let resolved_header = if header_row == "auto" {
        let (_, recommended) = workbook
            .find_header_candidates(index)
            .map_err(crate::cli::error::anyhow_to_app_error)?;
        recommended
    } else {
        header_row
            .parse::<usize>()
            .ok()
            .filter(|&r| r >= 1 && r <= sheet_obj.max_rows)
    };

    let sample_mode = if resolved_header.is_some() {
        "records"
    } else {
        "rows"
    };

    let range_str = format!(
        "{}{}:{}{}",
        index_to_col_name(start_col),
        start_row,
        index_to_col_name(end_col),
        sample_end_row
    );

    let data = if let Some(header_row_idx) = resolved_header {
        // Build records with headers
        let mut headers = Vec::new();
        if header_row_idx < sheet_obj.data.len() {
            for col in start_col..=end_col {
                let val = if col < sheet_obj.data[header_row_idx].len() {
                    sheet_obj.data[header_row_idx][col].value.clone()
                } else {
                    String::new()
                };
                headers.push(val);
            }
        }

        let mut records = Vec::new();
        for row in start_row..=sample_end_row {
            if row == header_row_idx {
                continue;
            }
            if row >= sheet_obj.data.len() {
                break;
            }
            let mut record = serde_json::Map::new();
            for (col_idx, col) in (start_col..=end_col).enumerate() {
                let key = headers.get(col_idx).cloned().unwrap_or_default();
                let key = if key.is_empty() {
                    format!("col_{}", col_idx + 1)
                } else {
                    key
                };
                let value = if col < sheet_obj.data[row].len() {
                    crate::json_export::process_cell_value(&sheet_obj.data[row][col])
                } else {
                    Value::Null
                };
                record.insert(key, value);
            }
            records.push(Value::Object(record));
        }

        json!({
            "resolved_header_row": header_row_idx,
            "sample_mode": sample_mode,
            "records": records,
        })
    } else {
        // Raw rows
        let mut row_values = Vec::new();
        for row in start_row..=sample_end_row {
            if row >= sheet_obj.data.len() {
                break;
            }
            let mut cols = Vec::new();
            for col in start_col..=end_col {
                let value = if col < sheet_obj.data[row].len() {
                    crate::json_export::process_cell_value(&sheet_obj.data[row][col])
                } else {
                    Value::Null
                };
                cols.push(value);
            }
            row_values.push(Value::Array(cols));
        }

        json!({
            "resolved_header_row": Value::Null,
            "sample_mode": sample_mode,
            "rows": row_values,
        })
    };

    Ok(envelope::success_envelope(
        "inspect.sample",
        &path_str,
        &format_str,
        envelope::target_range(&sheet_name, index, &range_str),
        json!({}),
        data,
        vec![],
    ))
}
