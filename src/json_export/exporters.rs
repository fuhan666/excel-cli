use anyhow::{Context, Result};
use indexmap::IndexMap;
use serde::Serialize;

use std::fs::File;
use std::io::Write;
use std::path::Path;

use crate::excel::{Sheet, Workbook};
use crate::json_export::converters::process_cell_value;
use crate::json_export::extractors::{extract_horizontal_headers, extract_vertical_headers};
use crate::json_export::types::{HeaderDirection, OrderedSheetData};

pub fn serialize_to_json<T: Serialize>(data: &T) -> Result<String> {
    serde_json::to_string_pretty(data).context("Failed to serialize data to JSON")
}

fn write_json_to_file<T: Serialize>(data: &T, path: &Path) -> Result<()> {
    let mut file =
        File::create(path).with_context(|| format!("Failed to create file: {}", path.display()))?;

    let json_string = serialize_to_json(data)?;

    file.write_all(json_string.as_bytes())
        .with_context(|| format!("Failed to write to file: {}", path.display()))?;

    Ok(())
}

// Process a single sheet for all-sheets export
pub fn process_sheet_for_json(
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

            let row_count = sheet.data.len().saturating_sub(header_count + 1);
            let mut sheet_data = Vec::with_capacity(row_count);

            let mut ordered_headers: Vec<(usize, &String)> = headers
                .iter()
                .map(|(col_idx, header)| (*col_idx, header))
                .collect();
            ordered_headers.sort_by_key(|(col_idx, _)| *col_idx);

            // Process each data row
            for row_idx in (header_count + 1)..sheet.data.len() {
                let mut row_data = IndexMap::with_capacity(ordered_headers.len());

                for (col_idx, header) in &ordered_headers {
                    if row_idx < sheet.data.len() && *col_idx < sheet.data[row_idx].len() {
                        let cell = &sheet.data[row_idx][*col_idx];

                        if !header.is_empty() {
                            let json_value = process_cell_value(cell);
                            row_data.insert((*header).clone(), json_value);
                        }
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

            let col_count = sheet.data[0].len().saturating_sub(header_count + 1);
            let mut sheet_data = Vec::with_capacity(col_count);

            let mut ordered_headers: Vec<(usize, &String)> = headers
                .iter()
                .map(|(row_idx, header)| (*row_idx, header))
                .collect();
            ordered_headers.sort_by_key(|(row_idx, _)| *row_idx);

            // Process each data column
            for col_idx in (header_count + 1)..sheet.data[0].len() {
                let mut obj = IndexMap::with_capacity(ordered_headers.len());

                for (row_idx, header) in &ordered_headers {
                    if *row_idx < sheet.data.len() && col_idx < sheet.data[*row_idx].len() {
                        let cell = &sheet.data[*row_idx][col_idx];

                        if !header.is_empty() {
                            let json_value = process_cell_value(cell);
                            obj.insert((*header).clone(), json_value);
                        }
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

// Export JSON file for a single sheet
pub fn export_json(
    sheet: &Sheet,
    direction: HeaderDirection,
    header_count: usize,
    path: &Path,
) -> Result<()> {
    let sheet_data = process_sheet_for_json(sheet, direction, header_count)?;
    write_json_to_file(&sheet_data, path)
}

pub fn generate_all_sheets_json(
    workbook: &Workbook,
    direction: HeaderDirection,
    header_count: usize,
) -> Result<IndexMap<String, OrderedSheetData>> {
    let sheet_names = workbook.get_sheet_names();

    let mut all_sheets = IndexMap::with_capacity(sheet_names.len());

    let current_sheet_index = workbook.get_current_sheet_index();

    // Process each sheet
    for (index, sheet_name) in sheet_names.iter().enumerate() {
        let sheet_data = if index == current_sheet_index {
            process_sheet_for_json(workbook.get_current_sheet(), direction, header_count)?
        } else {
            // Need to switch sheets - create a clone and process
            let mut wb_clone = workbook.clone();
            wb_clone.switch_sheet(index)?;
            process_sheet_for_json(wb_clone.get_current_sheet(), direction, header_count)?
        };

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

    write_json_to_file(&all_sheets, path)
}
