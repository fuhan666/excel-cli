use regex::Regex;
use serde_json::{json, Value};
use std::path::PathBuf;

use crate::cli::args::{OutputFormat, OutputShape, ReadCommands};
use crate::cli::common::{file_format, format_bounds, sheet_by_index};
use crate::cli::envelope;
use crate::cli::error::AppError;
use crate::cli::sheet_query::{
    load_target_sheet, read_header_values, resolve_bounds, resolve_optional_header_row,
    stable_record_keys,
};
use crate::excel::{open_workbook, CellType, Sheet};
use crate::utils::{index_to_col_name, parse_cell_reference};

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
            select,
            filters,
            limit,
            offset,
            non_empty,
            output_shape,
            format,
        } => read_rows(
            "read.rows",
            false,
            RowReadRequest {
                file,
                sheet,
                sheet_index,
                range,
                header_row,
                select,
                filters,
                limit,
                offset,
                non_empty,
                output_shape,
                format,
            },
        ),
        ReadCommands::Records {
            file,
            sheet,
            sheet_index,
            range,
            header_row,
            select,
            filters,
            limit,
            offset,
            non_empty,
            output_shape,
            format,
        } => read_rows(
            "read.records",
            true,
            RowReadRequest {
                file,
                sheet,
                sheet_index,
                range,
                header_row,
                select,
                filters,
                limit,
                offset,
                non_empty,
                output_shape,
                format,
            },
        ),
    }
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

    let mut workbook =
        open_workbook(&file, false).map_err(crate::cli::error::anyhow_to_app_error)?;

    let resolved_sheet = load_target_sheet(&workbook, &sheet, &sheet_index)?;

    workbook
        .ensure_sheet_loaded(resolved_sheet.index, &resolved_sheet.name)
        .map_err(crate::cli::error::anyhow_to_app_error)?;

    let sheet_obj = sheet_by_index(&workbook, resolved_sheet.index, &resolved_sheet.name)?;

    let cell_ref = cell.to_ascii_uppercase();
    let in_bounds = row < sheet_obj.data.len() && col < sheet_obj.data[row].len();
    let (value, cell_type, formula) = if in_bounds {
        let c = &sheet_obj.data[row][col];
        let formula =
            workbook.formula_for_cell(resolved_sheet.index, &resolved_sheet.name, &cell_ref);
        let type_str = if c.is_formula || formula.is_some() {
            "formula"
        } else {
            match c.cell_type {
                CellType::Text => "text",
                CellType::Number => "number",
                CellType::Date => "date",
                CellType::Boolean => "boolean",
                CellType::Empty => "empty",
            }
        };
        (crate::json_export::process_cell_value(c), type_str, formula)
    } else {
        (Value::Null, "empty", None)
    };

    let mut data = serde_json::Map::new();
    data.insert("cell".to_string(), json!(cell_ref));
    data.insert("value".to_string(), value);
    data.insert("type".to_string(), json!(cell_type));
    if let Some(formula) = formula {
        data.insert("formula".to_string(), json!(formula));
    }

    Ok(envelope::success_envelope(
        "read.cell",
        &path_str,
        &format_str,
        envelope::target_cell(&resolved_sheet.name, resolved_sheet.index, &cell_ref),
        json!({}),
        Value::Object(data),
        vec![],
    ))
}

#[derive(Clone, Copy)]
enum FilterOp {
    Eq,
    Ne,
    Gt,
    Gte,
    Lt,
    Lte,
    Contains,
    Regex,
    IsNull,
    NotNull,
}

struct FilterSpec {
    raw: String,
    col_idx: usize,
    op: FilterOp,
    value: String,
    numeric_value: Option<f64>,
    regex: Option<Regex>,
}

struct RowReadRequest {
    file: PathBuf,
    sheet: Option<String>,
    sheet_index: Option<usize>,
    range: Option<String>,
    header_row: String,
    select: Option<String>,
    filters: Vec<String>,
    limit: Option<usize>,
    offset: Option<usize>,
    non_empty: bool,
    output_shape: OutputShape,
    format: OutputFormat,
}

#[derive(Clone, Copy)]
struct RowBounds {
    start_row: usize,
    end_row: usize,
    start_col: usize,
    end_col: usize,
}

struct RowOutput {
    values: Vec<Value>,
    row_count: usize,
    truncated: bool,
}

#[derive(Clone, Copy)]
struct RowOutputFormat<'a> {
    selected_indices: &'a [usize],
    columns: &'a [String],
    output_shape: OutputShape,
}

struct RowCollectRequest<'a> {
    sheet: &'a Sheet,
    bounds: RowBounds,
    output_format: RowOutputFormat<'a>,
    filters: &'a [FilterSpec],
    non_empty: bool,
    offset: usize,
    limit: Option<usize>,
}

fn invalid_query(message: impl Into<String>) -> AppError {
    AppError::InvalidQuery {
        message: message.into(),
    }
}

fn sheet_row_values(sheet: &Sheet, row: usize, bounds: RowBounds) -> Option<Vec<Value>> {
    if row >= sheet.data.len() {
        return None;
    }

    let values = (bounds.start_col..=bounds.end_col)
        .map(|col| {
            if col < sheet.data[row].len() {
                crate::json_export::process_cell_value(&sheet.data[row][col])
            } else {
                Value::Null
            }
        })
        .collect();

    Some(values)
}

fn row_passes_filters(row: &[Value], filters: &[FilterSpec], non_empty: bool) -> bool {
    if non_empty && row.iter().all(is_empty_cell) {
        return false;
    }

    filters.iter().all(|filter| filter_matches(row, filter))
}

fn output_row(row: &[Value], output_format: RowOutputFormat<'_>) -> Value {
    if matches!(
        output_format.output_shape,
        OutputShape::Records | OutputShape::Jsonl
    ) {
        let mut record = serde_json::Map::new();
        for idx in output_format.selected_indices {
            let value = row.get(*idx).cloned().unwrap_or(Value::Null);
            record.insert(output_format.columns[*idx].clone(), value);
        }
        return Value::Object(record);
    }

    Value::Array(
        output_format
            .selected_indices
            .iter()
            .map(|idx| row.get(*idx).cloned().unwrap_or(Value::Null))
            .collect(),
    )
}

fn collect_row_output(request: RowCollectRequest<'_>) -> RowOutput {
    let mut values = Vec::new();
    let mut skipped = 0usize;
    let mut truncated = false;

    for row_idx in request.bounds.start_row..=request.bounds.end_row {
        let Some(row) = sheet_row_values(request.sheet, row_idx, request.bounds) else {
            break;
        };
        if !row_passes_filters(&row, request.filters, request.non_empty) {
            continue;
        }
        if skipped < request.offset {
            skipped += 1;
            continue;
        }
        if request.limit.is_some_and(|size| values.len() >= size) {
            truncated = true;
            break;
        }

        values.push(output_row(&row, request.output_format));
    }

    let row_count = values.len();
    RowOutput {
        values,
        row_count,
        truncated,
    }
}

fn parse_selected_columns(
    select: Option<String>,
    columns: &[String],
) -> Result<Vec<usize>, AppError> {
    let Some(select) = select else {
        return Ok((0..columns.len()).collect());
    };

    let mut selected = Vec::new();
    for field in select.split(',').map(str::trim) {
        if field.is_empty() {
            return Err(invalid_query("Selected column names cannot be empty"));
        }
        let col_idx = columns
            .iter()
            .position(|column| column == field)
            .ok_or_else(|| invalid_query(format!("Unknown selected column: {field}")))?;
        selected.push(col_idx);
    }

    Ok(selected)
}

fn parse_filters(filters: Vec<String>, columns: &[String]) -> Result<Vec<FilterSpec>, AppError> {
    filters
        .into_iter()
        .map(|raw| {
            let mut parts = raw.splitn(3, ':');
            let field = parts.next().unwrap_or_default().trim();
            let op = parts.next().unwrap_or_default().trim();
            let value = parts.next().ok_or_else(|| {
                invalid_query(format!("Invalid filter '{raw}'; expected field:op:value"))
            })?;
            let value = value.to_string();

            if field.is_empty() {
                return Err(invalid_query(format!(
                    "Invalid filter '{raw}'; field is empty"
                )));
            }

            let col_idx = columns
                .iter()
                .position(|column| column == field)
                .ok_or_else(|| invalid_query(format!("Unknown filter column: {field}")))?;

            let op = match op {
                "eq" => FilterOp::Eq,
                "ne" => FilterOp::Ne,
                "gt" => FilterOp::Gt,
                "gte" => FilterOp::Gte,
                "lt" => FilterOp::Lt,
                "lte" => FilterOp::Lte,
                "contains" => FilterOp::Contains,
                "regex" => FilterOp::Regex,
                "isnull" => FilterOp::IsNull,
                "notnull" => FilterOp::NotNull,
                "" => {
                    return Err(invalid_query(format!(
                        "Invalid filter '{raw}'; operator is empty"
                    )))
                }
                _ => return Err(invalid_query(format!("Unknown filter operator: {op}"))),
            };

            let numeric_value = if matches!(
                op,
                FilterOp::Gt | FilterOp::Gte | FilterOp::Lt | FilterOp::Lte
            ) {
                Some(value.trim().parse::<f64>().map_err(|_| {
                    invalid_query(format!("Numeric filter value is invalid in '{raw}'"))
                })?)
            } else {
                None
            };

            let regex = if matches!(op, FilterOp::Regex) {
                Some(
                    Regex::new(&value)
                        .map_err(|err| invalid_query(format!("Invalid regex filter: {err}")))?,
                )
            } else {
                None
            };

            Ok(FilterSpec {
                raw,
                col_idx,
                op,
                value,
                numeric_value,
                regex,
            })
        })
        .collect()
}

fn value_as_filter_text(value: &Value) -> String {
    match value {
        Value::Null => String::new(),
        Value::String(value) => value.clone(),
        Value::Number(value) => value.to_string(),
        Value::Bool(value) => value.to_string(),
        other => other.to_string(),
    }
}

fn value_as_number(value: &Value) -> Option<f64> {
    match value {
        Value::Number(number) => number.as_f64(),
        Value::String(value) => value.trim().parse::<f64>().ok(),
        _ => None,
    }
}

fn is_empty_cell(value: &Value) -> bool {
    match value {
        Value::Null => true,
        Value::String(value) => value.trim().is_empty(),
        _ => false,
    }
}

fn compare_numeric<F>(cell: &Value, filter_value: f64, compare: F) -> bool
where
    F: Fn(f64, f64) -> bool,
{
    let Some(left) = value_as_number(cell) else {
        return false;
    };
    compare(left, filter_value)
}

fn filter_matches(row: &[Value], filter: &FilterSpec) -> bool {
    let Some(cell) = row.get(filter.col_idx) else {
        return false;
    };

    match filter.op {
        FilterOp::Eq => {
            if let (Some(left), Ok(right)) =
                (value_as_number(cell), filter.value.trim().parse::<f64>())
            {
                (left - right).abs() < f64::EPSILON
            } else {
                value_as_filter_text(cell) == filter.value
            }
        }
        FilterOp::Ne => {
            if let (Some(left), Ok(right)) =
                (value_as_number(cell), filter.value.trim().parse::<f64>())
            {
                (left - right).abs() >= f64::EPSILON
            } else {
                value_as_filter_text(cell) != filter.value
            }
        }
        FilterOp::Gt => {
            compare_numeric(cell, filter.numeric_value.unwrap_or_default(), |a, b| a > b)
        }
        FilterOp::Gte => compare_numeric(cell, filter.numeric_value.unwrap_or_default(), |a, b| {
            a >= b
        }),
        FilterOp::Lt => {
            compare_numeric(cell, filter.numeric_value.unwrap_or_default(), |a, b| a < b)
        }
        FilterOp::Lte => compare_numeric(cell, filter.numeric_value.unwrap_or_default(), |a, b| {
            a <= b
        }),
        FilterOp::Contains => value_as_filter_text(cell).contains(&filter.value),
        FilterOp::Regex => filter
            .regex
            .as_ref()
            .is_some_and(|regex| regex.is_match(&value_as_filter_text(cell))),
        FilterOp::IsNull => is_empty_cell(cell),
        FilterOp::NotNull => !is_empty_cell(cell),
    }
}

fn read_range(
    file: std::path::PathBuf,
    sheet: Option<String>,
    sheet_index: Option<usize>,
    range: String,
) -> Result<Value, AppError> {
    let format_str = file_format(&file);
    let path_str = file.to_string_lossy().to_string();

    let mut workbook =
        open_workbook(&file, false).map_err(crate::cli::error::anyhow_to_app_error)?;

    let resolved_sheet = load_target_sheet(&workbook, &sheet, &sheet_index)?;

    workbook
        .ensure_sheet_loaded(resolved_sheet.index, &resolved_sheet.name)
        .map_err(crate::cli::error::anyhow_to_app_error)?;

    let sheet_obj = sheet_by_index(&workbook, resolved_sheet.index, &resolved_sheet.name)?;
    let bounds = resolve_bounds(&workbook, sheet_obj, resolved_sheet.index, Some(&range))?;

    let mut rows = Vec::new();
    for row in bounds.start_row..=bounds.end_row {
        let mut cols = Vec::new();
        for col in bounds.start_col..=bounds.end_col {
            let value = if row < sheet_obj.data.len() && col < sheet_obj.data[row].len() {
                crate::json_export::process_cell_value(&sheet_obj.data[row][col])
            } else {
                Value::Null
            };
            cols.push(value);
        }
        rows.push(Value::Array(cols));
    }

    let range_str = format_bounds(bounds);

    let data = json!({
        "range": range_str,
        "rows": rows,
    });

    Ok(envelope::success_envelope(
        "read.range",
        &path_str,
        &format_str,
        envelope::target_range(&resolved_sheet.name, resolved_sheet.index, &range_str),
        json!({}),
        data,
        vec![],
    ))
}

fn read_rows(
    command: &'static str,
    command_requires_header: bool,
    request: RowReadRequest,
) -> Result<Value, AppError> {
    let RowReadRequest {
        file,
        sheet,
        sheet_index,
        range,
        header_row,
        select,
        filters,
        limit,
        offset,
        non_empty,
        output_shape,
        format,
    } = request;

    if output_shape == OutputShape::Jsonl && matches!(format, OutputFormat::Text) {
        return Err(AppError::InvalidArgs {
            message: "--output-shape jsonl cannot be combined with --format text".to_string(),
        });
    }

    let format_str = file_format(&file);
    let path_str = file.to_string_lossy().to_string();

    let mut workbook =
        open_workbook(&file, false).map_err(crate::cli::error::anyhow_to_app_error)?;

    let resolved_sheet = load_target_sheet(&workbook, &sheet, &sheet_index)?;

    workbook
        .ensure_sheet_loaded(resolved_sheet.index, &resolved_sheet.name)
        .map_err(crate::cli::error::anyhow_to_app_error)?;

    let sheet_obj = sheet_by_index(&workbook, resolved_sheet.index, &resolved_sheet.name)?;
    let requested_bounds =
        resolve_bounds(&workbook, sheet_obj, resolved_sheet.index, range.as_deref())?;
    let resolved_header =
        resolve_optional_header_row(&workbook, sheet_obj, resolved_sheet.index, &header_row)?;

    let range_str = format_bounds(requested_bounds);

    if resolved_header.is_none()
        && (command_requires_header
            || matches!(output_shape, OutputShape::Records | OutputShape::Jsonl))
    {
        return Err(invalid_query(
            "A resolved header row is required for records or jsonl output",
        ));
    }

    let (has_header, columns, data_start_row) = if let Some(header_row_idx) = resolved_header {
        let headers = read_header_values(sheet_obj, header_row_idx, requested_bounds);
        let columns = stable_record_keys(&headers, requested_bounds.start_col);
        let data_start_row = requested_bounds
            .start_row
            .max(header_row_idx.saturating_add(1));
        (true, columns, data_start_row)
    } else {
        let columns: Vec<String> = (requested_bounds.start_col..=requested_bounds.end_col)
            .map(|col| format!("col_{}", index_to_col_name(col)))
            .collect();
        (false, columns, requested_bounds.start_row)
    };

    let selected_indices = parse_selected_columns(select, &columns)?;
    let parsed_filters = parse_filters(filters, &columns)?;
    let applied_filters: Vec<String> = parsed_filters
        .iter()
        .map(|filter| filter.raw.clone())
        .collect();
    let selected_columns: Vec<String> = selected_indices
        .iter()
        .map(|idx| columns[*idx].clone())
        .collect();

    let row_output = collect_row_output(RowCollectRequest {
        sheet: sheet_obj,
        bounds: RowBounds {
            start_row: data_start_row,
            end_row: requested_bounds.end_row,
            start_col: requested_bounds.start_col,
            end_col: requested_bounds.end_col,
        },
        output_format: RowOutputFormat {
            selected_indices: &selected_indices,
            columns: &columns,
            output_shape,
        },
        filters: &parsed_filters,
        non_empty,
        offset: offset.unwrap_or(0),
        limit,
    });

    let row_count = row_output.row_count;
    let truncated = row_output.truncated;

    let data = if matches!(output_shape, OutputShape::Records | OutputShape::Jsonl) {
        json!({
            "resolved_header_row": resolved_header.unwrap(),
            "mode": output_shape.as_str(),
            "records": row_output.values,
        })
    } else {
        json!({
            "resolved_header_row": if has_header {
                resolved_header.map(Value::from).unwrap_or(Value::Null)
            } else {
                Value::Null
            },
            "mode": "rows",
            "rows": row_output.values,
        })
    };

    let meta = json!({
        "applied_filters": applied_filters,
        "selected_columns": selected_columns,
        "row_count": row_count,
        "truncated": truncated,
        "output_shape": output_shape.as_str(),
    });

    Ok(envelope::success_envelope(
        command,
        &path_str,
        &format_str,
        envelope::target_range(&resolved_sheet.name, resolved_sheet.index, &range_str),
        meta,
        data,
        vec![],
    ))
}
