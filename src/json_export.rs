use anyhow::{Context, Result};
use chrono::{Duration, NaiveDate, NaiveDateTime};
use indexmap::IndexMap;
use serde::Serialize;
use serde_json::{json, Value};
use std::collections::HashMap;
use std::fs::File;
use std::io::Write;
use std::path::Path;

use crate::excel::{CellType, DataTypeInfo, Sheet, Workbook};

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum HeaderDirection {
    Horizontal,
    Vertical,
}

impl HeaderDirection {
    pub fn from_str(s: &str) -> Option<Self> {
        match s.to_lowercase().as_str() {
            "h" | "horizontal" => Some(HeaderDirection::Horizontal),
            "v" | "vertical" => Some(HeaderDirection::Vertical),
            _ => None,
        }
    }
}

// Convert Excel date number to ISO date string
fn excel_date_to_iso_string(excel_date: f64) -> String {
    let days = if excel_date > 59.0 {
        excel_date - 1.0
    } else {
        excel_date
    };

    let base_date = NaiveDate::from_ymd_opt(1900, 1, 1).unwrap();
    let whole_days = days.trunc() as i64;
    let fractional_day = days.fract();

    let date = base_date + Duration::days(whole_days - 1); // Subtract 1 because Excel day 1 is 1900-01-01

    if fractional_day > 0.0 {
        let seconds_in_day = 24.0 * 60.0 * 60.0;
        let seconds = (fractional_day * seconds_in_day).round() as u32;

        let hours = seconds / 3600;
        let minutes = (seconds % 3600) / 60;
        let secs = seconds % 60;

        let datetime = NaiveDateTime::new(
            date,
            chrono::NaiveTime::from_hms_opt(hours, minutes, secs).unwrap(),
        );

        datetime.format("%Y-%m-%dT%H:%M:%S").to_string()
    } else {
        date.format("%Y-%m-%d").to_string()
    }
}

// Process cell value based on its type
fn process_cell_value(cell: &crate::excel::Cell) -> Value {
    if cell.value.is_empty() {
        return Value::Null;
    }

    if let Some(original_type) = &cell.original_type {
        match original_type {
            DataTypeInfo::Float(f) => {
                if f.fract() == 0.0 {
                    json!(f.trunc() as i64)
                } else {
                    json!(f)
                }
            }
            DataTypeInfo::Int(i) => json!(i),
            DataTypeInfo::DateTime(dt) => {
                if *dt >= 0.0 {
                    json!(excel_date_to_iso_string(*dt))
                } else {
                    json!(cell.value)
                }
            }
            DataTypeInfo::DateTimeIso(s) => json!(s),
            DataTypeInfo::Bool(b) => json!(b),
            DataTypeInfo::Empty => Value::Null,
            _ => json!(cell.value),
        }
    } else {
        match cell.cell_type {
            CellType::Number => {
                if let Ok(num) = cell.value.parse::<f64>() {
                    if num.fract() == 0.0 {
                        json!(num.trunc() as i64)
                    } else {
                        json!(num)
                    }
                } else {
                    json!(cell.value)
                }
            }
            CellType::Boolean => {
                if cell.value.to_lowercase() == "true" {
                    json!(true)
                } else if cell.value.to_lowercase() == "false" {
                    json!(false)
                } else {
                    json!(cell.value)
                }
            }
            CellType::Date => {
                if let Ok(excel_date) = cell.value.parse::<f64>() {
                    if excel_date >= 0.0 {
                        json!(excel_date_to_iso_string(excel_date))
                    } else {
                        json!(cell.value)
                    }
                } else {
                    json!(cell.value)
                }
            }
            CellType::Empty => Value::Null,
            _ => json!(cell.value), // Text, etc.
        }
    }
}

// Process JSON export with horizontal headers
fn export_horizontal_json(sheet: &Sheet, header_rows: usize, path: &Path) -> Result<()> {
    if header_rows == 0 || header_rows >= sheet.data.len() {
        anyhow::bail!("Invalid header rows: {}", header_rows);
    }

    let headers = extract_horizontal_headers(sheet, header_rows)?;

    let mut ordered_headers: Vec<(usize, String)> = headers
        .iter()
        .map(|(col_idx, header)| (*col_idx, header.clone()))
        .collect();
    ordered_headers.sort_by_key(|(col_idx, _)| *col_idx);

    let mut json_data = Vec::new();

    for row_idx in (header_rows + 1)..sheet.data.len() {
        let mut row_data = IndexMap::new();

        for (col_idx, header) in &ordered_headers {
            let cell = &sheet.data[row_idx][*col_idx];

            if !header.is_empty() {
                let json_value = process_cell_value(cell);
                row_data.insert(header.clone(), json_value);
            }
        }

        if !row_data.is_empty() {
            json_data.push(row_data);
        }
    }

    write_json_to_file(&json_data, path)
}

// Process JSON export with vertical headers
fn export_vertical_json(sheet: &Sheet, header_cols: usize, path: &Path) -> Result<()> {
    if header_cols == 0 || header_cols >= sheet.data[0].len() {
        anyhow::bail!("Invalid header columns: {}", header_cols);
    }

    let headers = extract_vertical_headers(sheet, header_cols)?;

    let mut ordered_headers: Vec<(usize, String)> = headers
        .iter()
        .map(|(row_idx, header)| (*row_idx, header.clone()))
        .collect();
    ordered_headers.sort_by_key(|(row_idx, _)| *row_idx);

    let mut json_data = Vec::new();

    for col_idx in (header_cols + 1)..sheet.data[0].len() {
        let mut obj = IndexMap::new();

        for (row_idx, header) in &ordered_headers {
            let cell = &sheet.data[*row_idx][col_idx];

            if !header.is_empty() {
                let json_value = process_cell_value(cell);
                obj.insert(header.clone(), json_value);
            }
        }

        if !obj.is_empty() {
            json_data.push(obj);
        }
    }

    write_json_to_file(&json_data, path)
}

// Extract horizontal headers
fn extract_horizontal_headers(sheet: &Sheet, header_rows: usize) -> Result<HashMap<usize, String>> {
    let mut headers = HashMap::new();
    let mut last_values_by_row: HashMap<usize, String> = HashMap::new();

    for col_idx in 1..sheet.data[0].len() {
        let mut header_parts = Vec::new();

        for row_idx in 1..=header_rows {
            if row_idx < sheet.data.len() && col_idx < sheet.data[row_idx].len() {
                let cell_value = &sheet.data[row_idx][col_idx].value;

                if cell_value.is_empty() {
                    if let Some(last_value) = last_values_by_row.get(&row_idx) {
                        header_parts.push(last_value.clone());
                    } else {
                        if row_idx > 1 {
                            let prev_row_idx = row_idx - 1;
                            let prev_header_parts_len = header_parts.len();

                            if prev_header_parts_len > 0 && prev_row_idx >= 1 {
                                header_parts.push(header_parts[prev_header_parts_len - 1].clone());
                            }
                        }
                    }
                } else {
                    last_values_by_row.insert(row_idx, cell_value.clone());
                    header_parts.push(cell_value.clone());
                }
            }
        }

        let header = header_parts.join("-");

        if !header.is_empty() {
            headers.insert(col_idx, header);
        }
    }

    Ok(headers)
}

// Extract vertical headers
fn extract_vertical_headers(sheet: &Sheet, header_cols: usize) -> Result<HashMap<usize, String>> {
    let mut headers = HashMap::new();
    let mut last_values_by_col: HashMap<usize, String> = HashMap::new();

    for row_idx in 1..sheet.data.len() {
        let mut header_parts = Vec::new();

        for col_idx in 1..=header_cols {
            if col_idx < sheet.data[0].len() && row_idx < sheet.data.len() {
                let cell_value = &sheet.data[row_idx][col_idx].value;

                if cell_value.is_empty() {
                    if let Some(last_value) = last_values_by_col.get(&col_idx) {
                        header_parts.push(last_value.clone());
                    } else {
                        if col_idx > 1 {
                            let prev_col_idx = col_idx - 1;
                            let prev_header_parts_len = header_parts.len();

                            if prev_header_parts_len > 0 && prev_col_idx >= 1 {
                                header_parts.push(header_parts[prev_header_parts_len - 1].clone());
                            }
                        }
                    }
                } else {
                    last_values_by_col.insert(col_idx, cell_value.clone());
                    header_parts.push(cell_value.clone());
                }
            }
        }

        let header = header_parts.join("-");

        if !header.is_empty() {
            headers.insert(row_idx, header);
        }
    }

    Ok(headers)
}

// Serialize data to JSON string
pub fn serialize_to_json<T: Serialize>(data: &T) -> Result<String> {
    serde_json::to_string_pretty(data).context("Failed to serialize data to JSON")
}

// Write data to JSON file
fn write_json_to_file<T: Serialize>(data: &T, path: &Path) -> Result<()> {
    let mut file =
        File::create(path).with_context(|| format!("Failed to create file: {}", path.display()))?;

    let json_string = serialize_to_json(data)?;

    file.write_all(json_string.as_bytes())
        .with_context(|| format!("Failed to write to file: {}", path.display()))?;

    Ok(())
}

// Export JSON file for a single sheet
pub fn export_json(
    sheet: &Sheet,
    direction: HeaderDirection,
    header_count: usize,
    path: &Path,
) -> Result<()> {
    match direction {
        HeaderDirection::Horizontal => export_horizontal_json(sheet, header_count, path),
        HeaderDirection::Vertical => export_vertical_json(sheet, header_count, path),
    }
}

// Type for sheet data that preserves field order
type OrderedSheetData = Vec<IndexMap<String, Value>>;

// Process a single sheet for all-sheets export
fn process_sheet_for_json(
    sheet: &Sheet,
    direction: HeaderDirection,
    header_count: usize,
) -> Result<OrderedSheetData> {
    match direction {
        HeaderDirection::Horizontal => {
            if header_count == 0 || header_count >= sheet.data.len() {
                anyhow::bail!("Invalid header rows: {}", header_count);
            }

            let headers = extract_horizontal_headers(sheet, header_count)?;

            let mut ordered_headers: Vec<(usize, String)> = headers
                .iter()
                .map(|(col_idx, header)| (*col_idx, header.clone()))
                .collect();
            ordered_headers.sort_by_key(|(col_idx, _)| *col_idx);

            let mut sheet_data = Vec::new();

            for row_idx in (header_count + 1)..sheet.data.len() {
                let mut row_data = IndexMap::new();

                for (col_idx, header) in &ordered_headers {
                    let cell = &sheet.data[row_idx][*col_idx];

                    if !header.is_empty() {
                        let json_value = process_cell_value(cell);
                        row_data.insert(header.clone(), json_value);
                    }
                }

                if !row_data.is_empty() {
                    sheet_data.push(row_data);
                }
            }

            Ok(sheet_data)
        }
        HeaderDirection::Vertical => {
            if header_count == 0 || header_count >= sheet.data[0].len() {
                anyhow::bail!("Invalid header columns: {}", header_count);
            }

            let headers = extract_vertical_headers(sheet, header_count)?;

            let mut ordered_headers: Vec<(usize, String)> = headers
                .iter()
                .map(|(row_idx, header)| (*row_idx, header.clone()))
                .collect();
            ordered_headers.sort_by_key(|(row_idx, _)| *row_idx);

            let mut sheet_data = Vec::new();

            for col_idx in (header_count + 1)..sheet.data[0].len() {
                let mut obj = IndexMap::new();

                for (row_idx, header) in &ordered_headers {
                    let cell = &sheet.data[*row_idx][col_idx];

                    if !header.is_empty() {
                        let json_value = process_cell_value(cell);
                        obj.insert(header.clone(), json_value);
                    }
                }

                if !obj.is_empty() {
                    sheet_data.push(obj);
                }
            }

            Ok(sheet_data)
        }
    }
}

// Generate JSON data for all sheets
pub fn generate_all_sheets_json(
    workbook: &Workbook,
    direction: HeaderDirection,
    header_count: usize,
) -> Result<IndexMap<String, OrderedSheetData>> {
    let mut all_sheets = IndexMap::new();
    let sheet_names = workbook.get_sheet_names();

    // Process each sheet
    for (index, sheet_name) in sheet_names.iter().enumerate() {
        // We need to temporarily switch to each sheet to process it
        let mut wb_clone = workbook.clone();
        wb_clone.switch_sheet(index)?;

        let sheet = wb_clone.get_current_sheet();
        let sheet_data = process_sheet_for_json(sheet, direction, header_count)?;

        // Add the sheet data to our collection with the sheet name as the key
        all_sheets.insert(sheet_name.clone(), sheet_data);
    }

    Ok(all_sheets)
}

// Export all sheets to a single JSON file
pub fn export_all_sheets_json(
    workbook: &Workbook,
    direction: HeaderDirection,
    header_count: usize,
    path: &Path,
) -> Result<()> {
    let all_sheets = generate_all_sheets_json(workbook, direction, header_count)?;

    // Write the combined data to file
    write_json_to_file(&all_sheets, path)
}
