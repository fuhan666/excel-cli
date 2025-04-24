use anyhow::Result;
use std::collections::HashMap;

use crate::excel::Sheet;

pub fn extract_horizontal_headers(
    sheet: &Sheet,
    header_rows: usize,
) -> Result<HashMap<usize, String>> {
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

pub fn extract_vertical_headers(
    sheet: &Sheet,
    header_cols: usize,
) -> Result<HashMap<usize, String>> {
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
