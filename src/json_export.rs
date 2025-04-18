use anyhow::{Context, Result};
use chrono::{Duration, NaiveDate, NaiveDateTime};
use indexmap::IndexMap;
use serde::Serialize;
use serde_json::{json, Value};
use std::collections::HashMap;
use std::fs::File;
use std::io::Write;
use std::path::Path;

use crate::excel::{CellType, DataTypeInfo, Sheet};

/// Header direction
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum HeaderDirection {
    /// Horizontal headers (headers in top rows)
    Horizontal,
    /// Vertical headers (headers in left columns)
    Vertical,
}

impl HeaderDirection {
    /// Parse header direction from string
    pub fn from_str(s: &str) -> Option<Self> {
        match s.to_lowercase().as_str() {
            "h" | "horizontal" => Some(HeaderDirection::Horizontal),
            "v" | "vertical" => Some(HeaderDirection::Vertical),
            _ => None,
        }
    }
}

/// Convert Excel date number to ISO date string
/// Excel dates are stored as the number of days since 1900-01-00 (yes, day 0)
/// with a few quirks (like the non-existent leap day in 1900)
fn excel_date_to_iso_string(excel_date: f64) -> String {
    // Excel has a quirk where it thinks 1900 is a leap year (it's not)
    // This affects dates after February 28, 1900
    let days = if excel_date > 59.0 {
        // Adjust for the non-existent leap day (February 29, 1900)
        excel_date - 1.0
    } else {
        excel_date
    };

    // Excel dates start from 1900-01-01, which is day 1 (not 0)
    let base_date = NaiveDate::from_ymd_opt(1900, 1, 1).unwrap();

    // Calculate the whole days and fractional part (for time)
    let whole_days = days.trunc() as i64;
    let fractional_day = days.fract();

    // Add the days to the base date
    let date = base_date + Duration::days(whole_days - 1); // Subtract 1 because Excel day 1 is 1900-01-01

    // Calculate time from fractional part (if any)
    if fractional_day > 0.0 {
        // Convert fractional day to seconds
        let seconds_in_day = 24.0 * 60.0 * 60.0;
        let seconds = (fractional_day * seconds_in_day).round() as u32;

        // Create a datetime with the time component
        let hours = seconds / 3600;
        let minutes = (seconds % 3600) / 60;
        let secs = seconds % 60;

        let datetime = NaiveDateTime::new(
            date,
            chrono::NaiveTime::from_hms_opt(hours, minutes, secs).unwrap(),
        );

        // Return ISO format
        datetime.format("%Y-%m-%dT%H:%M:%S").to_string()
    } else {
        // Just return the date in ISO format
        date.format("%Y-%m-%d").to_string()
    }
}

/// Process cell value based on its type
fn process_cell_value(cell: &crate::excel::Cell) -> Value {
    // Handle empty cells - return null instead of empty string
    if cell.value.is_empty() {
        return Value::Null;
    }

    // First check if we have original type information from calamine
    if let Some(original_type) = &cell.original_type {
        match original_type {
            // For numeric types, return as JSON number
            DataTypeInfo::Float(f) => {
                // Check if it's actually an integer (no decimal part)
                if f.fract() == 0.0 {
                    // Convert to integer to avoid decimal point in output
                    json!(f.trunc() as i64)
                } else {
                    json!(f)
                }
            }
            DataTypeInfo::Int(i) => json!(i),

            // For DateTime, convert to ISO string format
            DataTypeInfo::DateTime(dt) => {
                // Try to convert the Excel date number to ISO string
                if *dt >= 0.0 {
                    json!(excel_date_to_iso_string(*dt))
                } else {
                    // Fallback to the displayed value if conversion fails
                    json!(cell.value)
                }
            }
            DataTypeInfo::DateTimeIso(s) => json!(s),

            // For Boolean, return as JSON boolean
            DataTypeInfo::Bool(b) => json!(b),

            // For Empty, return null
            DataTypeInfo::Empty => Value::Null,

            // For all other types, return as string
            _ => json!(cell.value),
        }
    } else {
        // Fallback to using cell_type if original_type is not available
        match cell.cell_type {
            CellType::Number => {
                // Try to parse as number
                if let Ok(num) = cell.value.parse::<f64>() {
                    // Check if it's actually an integer (no decimal part)
                    if num.fract() == 0.0 {
                        // Convert to integer to avoid decimal point in output
                        json!(num.trunc() as i64)
                    } else {
                        json!(num)
                    }
                } else {
                    json!(cell.value)
                }
            }
            CellType::Boolean => {
                // Try to parse as boolean
                if cell.value.to_lowercase() == "true" {
                    json!(true)
                } else if cell.value.to_lowercase() == "false" {
                    json!(false)
                } else {
                    json!(cell.value)
                }
            }
            CellType::Date => {
                // Try to convert to ISO date string if it's a number
                if let Ok(excel_date) = cell.value.parse::<f64>() {
                    if excel_date >= 0.0 {
                        json!(excel_date_to_iso_string(excel_date))
                    } else {
                        json!(cell.value)
                    }
                } else {
                    // Keep as is if it's not a valid number
                    json!(cell.value)
                }
            }
            CellType::Empty => Value::Null,
            _ => json!(cell.value), // Text, etc.
        }
    }
}

/// Process JSON export with horizontal headers
fn export_horizontal_json(sheet: &Sheet, header_rows: usize, path: &Path) -> Result<()> {
    // Validate header row count
    if header_rows == 0 || header_rows >= sheet.data.len() {
        anyhow::bail!("Invalid header rows: {}", header_rows);
    }

    // Extract headers
    let headers = extract_horizontal_headers(sheet, header_rows)?;

    // Create ordered header list sorted by column index
    let mut ordered_headers: Vec<(usize, String)> = headers
        .iter()
        .map(|(col_idx, header)| (*col_idx, header.clone()))
        .collect();
    ordered_headers.sort_by_key(|(col_idx, _)| *col_idx);

    // Create JSON array
    let mut json_data = Vec::new();

    // Process data rows after header rows
    for row_idx in (header_rows + 1)..sheet.data.len() {
        let mut row_data = IndexMap::new();

        // Iterate through headers in column index order
        for (col_idx, header) in &ordered_headers {
            // Get cell
            let cell = &sheet.data[row_idx][*col_idx];

            // Only add data when header is not empty
            if !header.is_empty() {
                let json_value = process_cell_value(cell);
                row_data.insert(header.clone(), json_value);
            }
        }

        // Only add to JSON array when row data is not empty
        if !row_data.is_empty() {
            json_data.push(row_data);
        }
    }

    // Write data to JSON file
    write_json_to_file(&json_data, path)
}

/// Process JSON export with vertical headers
fn export_vertical_json(sheet: &Sheet, header_cols: usize, path: &Path) -> Result<()> {
    // Validate header column count
    if header_cols == 0 || header_cols >= sheet.data[0].len() {
        anyhow::bail!("Invalid header columns: {}", header_cols);
    }

    // Extract headers
    let headers = extract_vertical_headers(sheet, header_cols)?;

    // Create ordered header list sorted by row index
    let mut ordered_headers: Vec<(usize, String)> = headers
        .iter()
        .map(|(row_idx, header)| (*row_idx, header.clone()))
        .collect();
    ordered_headers.sort_by_key(|(row_idx, _)| *row_idx);

    // Create JSON object array
    let mut json_data = Vec::new();

    // For each column (starting after header columns)
    for col_idx in (header_cols + 1)..sheet.data[0].len() {
        let mut obj = IndexMap::new();

        // Iterate through headers in row index order
        for (row_idx, header) in &ordered_headers {
            // Get cell
            let cell = &sheet.data[*row_idx][col_idx];

            // Only add data when header is not empty
            if !header.is_empty() {
                let json_value = process_cell_value(cell);
                obj.insert(header.clone(), json_value);
            }
        }

        // Only add to JSON array when object is not empty
        if !obj.is_empty() {
            json_data.push(obj);
        }
    }

    // Write data to JSON file
    write_json_to_file(&json_data, path)
}

/// Extract horizontal headers (handling multi-row headers and merged cells)
fn extract_horizontal_headers(sheet: &Sheet, header_rows: usize) -> Result<HashMap<usize, String>> {
    let mut headers = HashMap::new();

    // Store the last non-empty value for each row to handle merged cells in the same row
    let mut last_values_by_row: HashMap<usize, String> = HashMap::new();

    // For each column
    for col_idx in 1..sheet.data[0].len() {
        let mut header_parts = Vec::new();

        // For each header row
        for row_idx in 1..=header_rows {
            if row_idx < sheet.data.len() && col_idx < sheet.data[row_idx].len() {
                let cell_value = &sheet.data[row_idx][col_idx].value;

                // Handle empty values caused by merged cells
                if cell_value.is_empty() {
                    // Try to get value from previous column in the same row (horizontal merged cells)
                    if let Some(last_value) = last_values_by_row.get(&row_idx) {
                        header_parts.push(last_value.clone());
                    } else {
                        // If no value from previous column in same row, try using value from same column in previous row (vertical merged cells)
                        if row_idx > 1 {
                            let prev_row_idx = row_idx - 1;
                            let prev_header_parts_len = header_parts.len();

                            // If previous row has been processed and has a value
                            if prev_header_parts_len > 0 && prev_row_idx >= 1 {
                                // Use value from previous row
                                header_parts.push(header_parts[prev_header_parts_len - 1].clone());
                            }
                        }
                    }
                } else {
                    // Record non-empty value
                    last_values_by_row.insert(row_idx, cell_value.clone());
                    header_parts.push(cell_value.clone());
                }
            }
        }

        // Join multi-row headers with "-"
        let header = header_parts.join("-");

        // Only add when header is not empty
        if !header.is_empty() {
            headers.insert(col_idx, header);
        }
    }

    Ok(headers)
}

/// Extract vertical headers (handling multi-column headers and merged cells)
fn extract_vertical_headers(sheet: &Sheet, header_cols: usize) -> Result<HashMap<usize, String>> {
    let mut headers = HashMap::new();

    // Store the last non-empty value for each column to handle merged cells in the same column
    let mut last_values_by_col: HashMap<usize, String> = HashMap::new();

    // For each row
    for row_idx in 1..sheet.data.len() {
        let mut header_parts = Vec::new();

        // For each header column
        for col_idx in 1..=header_cols {
            if col_idx < sheet.data[0].len() && row_idx < sheet.data.len() {
                let cell_value = &sheet.data[row_idx][col_idx].value;

                // Handle empty values caused by merged cells
                if cell_value.is_empty() {
                    // Try to get value from previous row in the same column (vertical merged cells)
                    if let Some(last_value) = last_values_by_col.get(&col_idx) {
                        header_parts.push(last_value.clone());
                    } else {
                        // If no value from previous row in same column, try using value from previous column in same row (horizontal merged cells)
                        if col_idx > 1 {
                            let prev_col_idx = col_idx - 1;
                            let prev_header_parts_len = header_parts.len();

                            // If previous column has been processed and has a value
                            if prev_header_parts_len > 0 && prev_col_idx >= 1 {
                                // Use value from previous column
                                header_parts.push(header_parts[prev_header_parts_len - 1].clone());
                            }
                        }
                    }
                } else {
                    // Record non-empty value
                    last_values_by_col.insert(col_idx, cell_value.clone());
                    header_parts.push(cell_value.clone());
                }
            }
        }

        // Join multi-column headers with "-"
        let header = header_parts.join("-");

        // Only add when header is not empty
        if !header.is_empty() {
            headers.insert(row_idx, header);
        }
    }

    Ok(headers)
}

/// Write data to JSON file
fn write_json_to_file<T: Serialize>(data: &T, path: &Path) -> Result<()> {
    // Create file
    let mut file =
        File::create(path).with_context(|| format!("Failed to create file: {}", path.display()))?;

    // Serialize to formatted JSON
    let json_string =
        serde_json::to_string_pretty(data).context("Failed to serialize data to JSON")?;

    // Write to file
    file.write_all(json_string.as_bytes())
        .with_context(|| format!("Failed to write to file: {}", path.display()))?;

    Ok(())
}

/// Export JSON file
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
