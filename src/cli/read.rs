use anyhow::Context;
use serde_json::{json, Value};

use crate::cli::args::{resolve_sheet_target, ReadCommands};
use crate::cli::envelope;
use crate::cli::error::AppError;
use crate::excel::{open_workbook, CellType};
use crate::utils::{index_to_col_name, parse_cell_reference, parse_range};

pub fn handle(cmd: ReadCommands) -> Result<Value, AppError> {
    match cmd {
        ReadCommands::Cell {
            file,
            sheet,
            sheet_index,
            cell,
            format: _,
        } => read_cell(file, sheet, sheet_index, cell),
        ReadCommands::Range {
            file,
            sheet,
            sheet_index,
            range,
            format: _,
        } => read_range(file, sheet, sheet_index, range),
        ReadCommands::Rows {
            file,
            sheet,
            sheet_index,
            range,
            header_row,
            format: _,
        } => read_rows(file, sheet, sheet_index, range, header_row),
    }
}

fn file_format(path: &std::path::Path) -> String {
    path.extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_lowercase())
        .unwrap_or_else(|| "unknown".to_string())
}

fn read_cell(
    file: std::path::PathBuf,
    sheet: Option<String>,
    sheet_index: Option<usize>,
    cell: String,
) -> Result<Value, AppError> {
    let format_str = file_format(&file);
    let path_str = file.to_string_lossy().to_string();

    let (row, col) = parse_cell_reference(&cell).ok_or_else(|| AppError::InvalidQuery {
        message: format!("Invalid cell reference: {}", cell),
    })?;

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

    let in_bounds = row < sheet_obj.data.len() && col < sheet_obj.data[row].len();
    let (value, cell_type) = if in_bounds {
        let c = &sheet_obj.data[row][col];
        let type_str = match c.cell_type {
            CellType::Text => "text",
            CellType::Number => "number",
            CellType::Date => "date",
            CellType::Boolean => "boolean",
            CellType::Empty => "empty",
        };
        (
            crate::json_export::process_cell_value(c),
            type_str,
        )
    } else {
        (Value::Null, "empty")
    };

    let data = json!({
        "cell": cell.to_ascii_uppercase(),
        "value": value,
        "type": cell_type,
    });

    Ok(envelope::success_envelope(
        "read.cell",
        &path_str,
        &format_str,
        envelope::target_cell(&sheet_name, index, &cell.to_ascii_uppercase()),
        json!({}),
        data,
        vec![],
    ))
}

fn read_range(
    file: std::path::PathBuf,
    sheet: Option<String>,
    sheet_index: Option<usize>,
    range: String,
) -> Result<Value, AppError> {
    let format_str = file_format(&file);
    let path_str = file.to_string_lossy().to_string();

    let ((mut start_row, mut start_col), (mut end_row, mut end_col)) =
        parse_range(&range).ok_or_else(|| AppError::InvalidQuery {
            message: format!("Invalid range format: {}", range),
        })?;

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

    let mut rows = Vec::new();
    for row in start_row..=end_row {
        let mut cols = Vec::new();
        for col in start_col..=end_col {
            let value = if row < sheet_obj.data.len() && col < sheet_obj.data[row].len() {
                crate::json_export::process_cell_value(&sheet_obj.data[row][col])
            } else {
                Value::Null
            };
            cols.push(value);
        }
        rows.push(Value::Array(cols));
    }

    let range_str = format!(
        "{}{}:{}{}",
        index_to_col_name(start_col),
        start_row,
        index_to_col_name(end_col),
        end_row
    );

    let data = json!({
        "range": range_str,
        "rows": rows,
    });

    Ok(envelope::success_envelope(
        "read.range",
        &path_str,
        &format_str,
        envelope::target_range(&sheet_name, index, &range_str),
        json!({}),
        data,
        vec![],
    ))
}

fn read_rows(
    file: std::path::PathBuf,
    sheet: Option<String>,
    sheet_index: Option<usize>,
    range: Option<String>,
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

    // Determine the range
    let ((mut start_row, mut start_col), (mut end_row, mut end_col)) = if let Some(ref r) = range {
        parse_range(r)
            .ok_or_else(|| AppError::InvalidQuery {
                message: format!("Invalid range format: {}", r),
            })?
    } else {
        let used = workbook
            .get_used_range(index)
            .unwrap_or_default();
        if used.is_empty() {
            ((1, 1), (1, 1))
        } else {
            parse_range(&used).unwrap_or(((1, 1), (1, 1)))
        }
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

    let range_str = format!(
        "{}{}:{}{}",
        index_to_col_name(start_col),
        start_row,
        index_to_col_name(end_col),
        end_row
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
        for row in start_row..=end_row {
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
            "mode": "records",
            "records": records,
        })
    } else {
        // Raw rows
        let mut row_values = Vec::new();
        for row in start_row..=end_row {
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
            "mode": "rows",
            "rows": row_values,
        })
    };

    Ok(envelope::success_envelope(
        "read.rows",
        &path_str,
        &format_str,
        envelope::target_range(&sheet_name, index, &range_str),
        json!({}),
        data,
        vec![],
    ))
}
