use anyhow::Context;
use quick_xml::events::Event;
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
        let record_keys = stable_record_keys(&headers, start_col);

        let mut records = Vec::new();
        let data_start_row = start_row.max(header_row_idx.saturating_add(1));
        for row in data_start_row..=end_row {
            if row >= sheet_obj.data.len() {
                break;
            }
            let mut record = serde_json::Map::new();
            for (col_idx, col) in (start_col..=end_col).enumerate() {
                let key = record_keys
                    .get(col_idx)
                    .cloned()
                    .unwrap_or_else(|| format!("col_{}", index_to_col_name(col)));
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
