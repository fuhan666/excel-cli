use anyhow::Context;
use quick_xml::events::Event;
use regex::Regex;
use serde_json::{json, Value};
use std::collections::HashMap;
use std::fs::File;
use std::io::{Read, Seek};
use std::path::Path;
use zip::ZipArchive;

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
            select,
            filters,
            limit,
            offset,
            non_empty,
            format: _,
        } => read_rows(
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
        ),
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

    let mut workbook =
        open_workbook(&file, false).map_err(crate::cli::error::anyhow_to_app_error)?;

    let index = resolve_sheet_target(&workbook, &sheet, &sheet_index)?;
    let sheet_name = workbook.get_sheet_names()[index].clone();

    workbook
        .ensure_sheet_loaded(index, &sheet_name)
        .map_err(crate::cli::error::anyhow_to_app_error)?;

    let sheet_obj = workbook
        .get_sheet_by_index(index)
        .with_context(|| format!("Sheet '{}' not found", sheet_name))
        .map_err(crate::cli::error::anyhow_to_app_error)?;

    let cell_ref = cell.to_ascii_uppercase();
    let in_bounds = row < sheet_obj.data.len() && col < sheet_obj.data[row].len();
    let (value, cell_type, formula) = if in_bounds {
        let c = &sheet_obj.data[row][col];
        let formula = c
            .formula
            .clone()
            .or_else(|| lookup_formula_in_xlsx(&file, &sheet_name, &cell_ref));
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
        envelope::target_cell(&sheet_name, index, &cell_ref),
        json!({}),
        Value::Object(data),
        vec![],
    ))
}

fn stable_record_keys(headers: &[String], start_col: usize) -> Vec<String> {
    let mut counts = HashMap::new();

    headers
        .iter()
        .enumerate()
        .map(|(offset, header)| {
            let base = if header.trim().is_empty() {
                format!("col_{}", index_to_col_name(start_col + offset))
            } else {
                header.trim().to_string()
            };

            let count = counts.entry(base.clone()).or_insert(0usize);
            *count += 1;
            if *count == 1 {
                base
            } else {
                format!("{base}_{count}")
            }
        })
        .collect()
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

fn invalid_query(message: impl Into<String>) -> AppError {
    AppError::InvalidQuery {
        message: message.into(),
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

fn read_zip_entry<R: Read + Seek>(archive: &mut ZipArchive<R>, entry_name: &str) -> Option<String> {
    let mut entry = archive.by_name(entry_name).ok()?;
    let mut contents = String::new();
    entry.read_to_string(&mut contents).ok()?;
    Some(contents)
}

fn attr_value(
    reader: &quick_xml::Reader<&[u8]>,
    event: &quick_xml::events::BytesStart<'_>,
    key: &[u8],
) -> Option<String> {
    for attr in event.attributes().flatten() {
        if attr.key.as_ref() == key {
            return attr
                .decode_and_unescape_value(reader.decoder())
                .ok()
                .map(|value| value.into_owned());
        }
    }
    None
}

fn resolve_xlsx_sheet_path<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    sheet_name: &str,
) -> Option<String> {
    let workbook_xml = read_zip_entry(archive, "xl/workbook.xml")?;
    let mut workbook_reader = quick_xml::Reader::from_str(&workbook_xml);
    workbook_reader.config_mut().trim_text(true);
    let mut workbook_buf = Vec::new();
    let mut relationship_id = None;

    loop {
        match workbook_reader.read_event_into(&mut workbook_buf).ok()? {
            Event::Start(event) | Event::Empty(event) if event.name().as_ref() == b"sheet" => {
                let name = attr_value(&workbook_reader, &event, b"name");
                if name.as_deref() == Some(sheet_name) {
                    relationship_id = attr_value(&workbook_reader, &event, b"r:id");
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        workbook_buf.clear();
    }

    let relationship_id = relationship_id?;
    let rels_xml = read_zip_entry(archive, "xl/_rels/workbook.xml.rels")?;
    let mut rels_reader = quick_xml::Reader::from_str(&rels_xml);
    rels_reader.config_mut().trim_text(true);
    let mut rels_buf = Vec::new();

    loop {
        match rels_reader.read_event_into(&mut rels_buf).ok()? {
            Event::Start(event) | Event::Empty(event)
                if event.name().as_ref() == b"Relationship" =>
            {
                let id = attr_value(&rels_reader, &event, b"Id");
                if id.as_deref() == Some(relationship_id.as_str()) {
                    let target = attr_value(&rels_reader, &event, b"Target")?;
                    return Some(if target.starts_with('/') {
                        target.trim_start_matches('/').to_string()
                    } else {
                        format!("xl/{target}")
                    });
                }
            }
            Event::Eof => break,
            _ => {}
        }
        rels_buf.clear();
    }

    None
}

fn lookup_formula_in_xlsx(file: &Path, sheet_name: &str, cell_ref: &str) -> Option<String> {
    let extension = file
        .extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| ext.to_ascii_lowercase())?;
    if extension != "xlsx" && extension != "xlsm" {
        return None;
    }

    let archive_file = File::open(file).ok()?;
    let mut archive = ZipArchive::new(archive_file).ok()?;
    let sheet_path = resolve_xlsx_sheet_path(&mut archive, sheet_name)?;
    let sheet_xml = read_zip_entry(&mut archive, &sheet_path)?;
    let target_ref = cell_ref.to_ascii_uppercase();

    let mut reader = quick_xml::Reader::from_str(&sheet_xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut current_cell = None;

    loop {
        match reader.read_event_into(&mut buf).ok()? {
            Event::Start(event) if event.name().as_ref() == b"c" => {
                current_cell = attr_value(&reader, &event, b"r")
                    .map(|reference| reference.to_ascii_uppercase());
            }
            Event::End(event) if event.name().as_ref() == b"c" => {
                current_cell = None;
            }
            Event::Start(event) if event.name().as_ref() == b"f" => {
                let mut formula = String::new();
                let end_tag = event.name().as_ref().to_vec();
                let mut inner_buf = Vec::new();
                loop {
                    match reader.read_event_into(&mut inner_buf).ok()? {
                        Event::Text(text) => formula.push_str(&text.unescape().ok()?),
                        Event::End(end_event)
                            if end_event.name().as_ref() == end_tag.as_slice() =>
                        {
                            break;
                        }
                        Event::Eof => return None,
                        _ => {}
                    }
                    inner_buf.clear();
                }

                if current_cell.as_deref() == Some(target_ref.as_str()) && !formula.is_empty() {
                    return Some(if formula.starts_with('=') {
                        formula
                    } else {
                        format!("={formula}")
                    });
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    None
}

fn read_range(
    file: std::path::PathBuf,
    sheet: Option<String>,
    sheet_index: Option<usize>,
    range: String,
) -> Result<Value, AppError> {
    let format_str = file_format(&file);
    let path_str = file.to_string_lossy().to_string();

    let ((mut start_row, mut start_col), (mut end_row, mut end_col)) = parse_range(&range)
        .ok_or_else(|| AppError::InvalidQuery {
            message: format!("Invalid range format: {}", range),
        })?;

    let mut workbook =
        open_workbook(&file, false).map_err(crate::cli::error::anyhow_to_app_error)?;

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
    select: Option<String>,
    filters: Vec<String>,
    limit: Option<usize>,
    offset: Option<usize>,
    non_empty: bool,
) -> Result<Value, AppError> {
    let format_str = file_format(&file);
    let path_str = file.to_string_lossy().to_string();

    let mut workbook =
        open_workbook(&file, false).map_err(crate::cli::error::anyhow_to_app_error)?;

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
        parse_range(r).ok_or_else(|| AppError::InvalidQuery {
            message: format!("Invalid range format: {}", r),
        })?
    } else {
        let used = workbook.get_used_range(index).unwrap_or_default();
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

    let (mode, columns, row_values) = if let Some(header_row_idx) = resolved_header {
        let mut headers = Vec::new();
        for col in start_col..=end_col {
            let val = if header_row_idx < sheet_obj.data.len()
                && col < sheet_obj.data[header_row_idx].len()
            {
                sheet_obj.data[header_row_idx][col].value.clone()
            } else {
                String::new()
            };
            headers.push(val);
        }
        let columns = stable_record_keys(&headers, start_col);

        let mut row_values = Vec::new();
        let data_start_row = start_row.max(header_row_idx.saturating_add(1));
        for row in data_start_row..=end_row {
            if row >= sheet_obj.data.len() {
                break;
            }
            let mut values = Vec::new();
            for col in start_col..=end_col {
                let value = if col < sheet_obj.data[row].len() {
                    crate::json_export::process_cell_value(&sheet_obj.data[row][col])
                } else {
                    Value::Null
                };
                values.push(value);
            }
            row_values.push(values);
        }

        ("records", columns, row_values)
    } else {
        let columns: Vec<String> = (start_col..=end_col)
            .map(|col| format!("col_{}", index_to_col_name(col)))
            .collect();
        let mut row_values = Vec::new();
        for row in start_row..=end_row {
            if row >= sheet_obj.data.len() {
                break;
            }
            let mut values = Vec::new();
            for col in start_col..=end_col {
                let value = if col < sheet_obj.data[row].len() {
                    crate::json_export::process_cell_value(&sheet_obj.data[row][col])
                } else {
                    Value::Null
                };
                values.push(value);
            }
            row_values.push(values);
        }

        ("rows", columns, row_values)
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

    let mut filtered_rows: Vec<Vec<Value>> = row_values;
    if non_empty {
        filtered_rows.retain(|row| row.iter().any(|cell| !is_empty_cell(cell)));
    }
    filtered_rows.retain(|row| {
        parsed_filters
            .iter()
            .all(|filter| filter_matches(row, filter))
    });

    let offset = offset.unwrap_or(0);
    let rows_after_offset: Vec<Vec<Value>> = filtered_rows.into_iter().skip(offset).collect();
    let truncated = limit.is_some_and(|size| rows_after_offset.len() > size);
    let returned_rows: Vec<Vec<Value>> = if let Some(size) = limit {
        rows_after_offset.into_iter().take(size).collect()
    } else {
        rows_after_offset
    };

    let row_count = returned_rows.len();

    let data = if mode == "records" {
        let records: Vec<Value> = returned_rows
            .into_iter()
            .map(|row| {
                let mut record = serde_json::Map::new();
                for idx in &selected_indices {
                    let value = row.get(*idx).cloned().unwrap_or(Value::Null);
                    record.insert(columns[*idx].clone(), value);
                }
                Value::Object(record)
            })
            .collect();

        json!({
            "resolved_header_row": resolved_header.unwrap(),
            "mode": "records",
            "records": records,
        })
    } else {
        let rows: Vec<Value> = returned_rows
            .into_iter()
            .map(|row| {
                Value::Array(
                    selected_indices
                        .iter()
                        .map(|idx| row.get(*idx).cloned().unwrap_or(Value::Null))
                        .collect(),
                )
            })
            .collect();

        json!({
            "resolved_header_row": Value::Null,
            "mode": "rows",
            "rows": rows,
        })
    };

    let meta = json!({
        "applied_filters": applied_filters,
        "selected_columns": selected_columns,
        "row_count": row_count,
        "truncated": truncated,
    });

    Ok(envelope::success_envelope(
        "read.rows",
        &path_str,
        &format_str,
        envelope::target_range(&sheet_name, index, &range_str),
        meta,
        data,
        vec![],
    ))
}
