use std::collections::HashMap;

use crate::cli::args::resolve_sheet_target;
use crate::cli::error::{anyhow_to_app_error, AppError};
use crate::excel::{Cell, Sheet, Workbook};
use crate::utils::{index_to_col_name, parse_range};

pub(crate) struct ResolvedSheet {
    pub(crate) index: usize,
    pub(crate) name: String,
}

#[derive(Clone, Copy)]
pub(crate) struct SheetBounds {
    pub(crate) start_row: usize,
    pub(crate) end_row: usize,
    pub(crate) start_col: usize,
    pub(crate) end_col: usize,
}

pub(crate) fn load_target_sheet(
    workbook: &Workbook,
    sheet: &Option<String>,
    sheet_index: &Option<usize>,
) -> Result<ResolvedSheet, AppError> {
    let index = resolve_sheet_target(workbook, sheet, sheet_index)?;
    let name = workbook
        .get_sheet_names()
        .get(index)
        .cloned()
        .ok_or_else(|| AppError::TargetNotFound {
            message: format!("Sheet index {} not found", index),
        })?;

    Ok(ResolvedSheet { index, name })
}

pub(crate) fn resolve_bounds(
    workbook: &Workbook,
    sheet: &Sheet,
    sheet_index: usize,
    range: Option<&str>,
) -> Result<SheetBounds, AppError> {
    let ((mut start_row, mut start_col), (mut end_row, mut end_col)) = if let Some(range) = range {
        parse_range(range).ok_or_else(|| AppError::InvalidQuery {
            message: format!("Invalid range format: {}", range),
        })?
    } else {
        let used_range = workbook.get_used_range(sheet_index).unwrap_or_default();
        if used_range.is_empty() {
            ((1, 1), (1, 1))
        } else {
            parse_range(&used_range).unwrap_or(((1, 1), (1, 1)))
        }
    };

    let max_row = sheet.max_rows.max(1);
    let max_col = sheet.max_cols.max(1);
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

    Ok(SheetBounds {
        start_row,
        end_row,
        start_col,
        end_col,
    })
}

pub(crate) fn resolve_header_row(
    workbook: &Workbook,
    sheet: &Sheet,
    sheet_index: usize,
    header_row: &str,
) -> Result<Option<usize>, AppError> {
    if header_row == "auto" {
        let (_, recommended) = workbook
            .find_header_candidates(sheet_index)
            .map_err(anyhow_to_app_error)?;
        return Ok(recommended);
    }

    let row = header_row
        .parse::<usize>()
        .map_err(|_| AppError::InvalidQuery {
            message: format!("Invalid header row: {}", header_row),
        })?;

    if row < 1 || row > sheet.max_rows {
        return Err(AppError::InvalidQuery {
            message: format!(
                "Header row {} is outside the used row range 1..={}",
                row, sheet.max_rows
            ),
        });
    }

    Ok(Some(row))
}

pub(crate) fn resolve_optional_header_row(
    workbook: &Workbook,
    sheet: &Sheet,
    sheet_index: usize,
    header_row: &str,
) -> Result<Option<usize>, AppError> {
    if header_row == "auto" {
        let (_, recommended) = workbook
            .find_header_candidates(sheet_index)
            .map_err(anyhow_to_app_error)?;
        return Ok(recommended);
    }

    Ok(header_row
        .parse::<usize>()
        .ok()
        .filter(|row| *row >= 1 && *row <= sheet.max_rows))
}

pub(crate) fn cell_at(sheet: &Sheet, row: usize, col: usize) -> Option<&Cell> {
    sheet.data.get(row).and_then(|row_data| row_data.get(col))
}

pub(crate) fn cell_has_formula(cell: &Cell) -> bool {
    cell.is_formula || cell.formula.is_some()
}

pub(crate) fn cell_is_present(cell: Option<&Cell>) -> bool {
    cell.map(|cell| !cell.value.trim().is_empty() || cell_has_formula(cell))
        .unwrap_or(false)
}

pub(crate) fn header_value(sheet: &Sheet, row: usize, col: usize) -> String {
    cell_at(sheet, row, col)
        .filter(|cell| !cell_has_formula(cell))
        .map(|cell| cell.value.trim().to_string())
        .unwrap_or_default()
}

pub(crate) fn stable_record_keys(headers: &[String], start_col: usize) -> Vec<String> {
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

pub(crate) fn read_header_values(
    sheet: &Sheet,
    header_row: usize,
    bounds: SheetBounds,
) -> Vec<String> {
    (bounds.start_col..=bounds.end_col)
        .map(|col| {
            if header_row < sheet.data.len() && col < sheet.data[header_row].len() {
                sheet.data[header_row][col].value.clone()
            } else {
                String::new()
            }
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::excel::{CellType, Sheet};

    fn sheet_with_values(name: &str, values: &[&[&str]]) -> Sheet {
        let max_rows = values.len();
        let max_cols = values.iter().map(|row| row.len()).max().unwrap_or(0);
        let mut data = vec![vec![Cell::empty(); max_cols + 1]; max_rows + 1];

        for (row_idx, row) in values.iter().enumerate() {
            for (col_idx, value) in row.iter().enumerate() {
                data[row_idx + 1][col_idx + 1] = Cell::new((*value).to_string(), false);
            }
        }

        Sheet {
            name: name.to_string(),
            data,
            max_rows,
            max_cols,
            is_loaded: true,
        }
    }

    #[test]
    fn resolve_bounds_clamps_and_normalizes_explicit_ranges() {
        let workbook = Workbook::from_sheets_for_test(vec![sheet_with_values(
            "Orders",
            &[&["order_id", "customer"], &["1001", "Alice"]],
        )]);
        let sheet = workbook.get_sheet_by_index(0).unwrap();

        let bounds = resolve_bounds(&workbook, sheet, 0, Some("D5:B2")).unwrap();

        assert_eq!(bounds.start_row, 2);
        assert_eq!(bounds.end_row, 2);
        assert_eq!(bounds.start_col, 2);
        assert_eq!(bounds.end_col, 2);
    }

    #[test]
    fn resolve_bounds_falls_back_to_a1_for_empty_used_ranges() {
        let mut sheet = Sheet::blank("Empty".to_string());
        sheet.max_rows = 0;
        sheet.max_cols = 0;
        let workbook = Workbook::from_sheets_for_test(vec![sheet]);
        let sheet = workbook.get_sheet_by_index(0).unwrap();

        let bounds = resolve_bounds(&workbook, sheet, 0, None).unwrap();

        assert_eq!(bounds.start_row, 1);
        assert_eq!(bounds.end_row, 1);
        assert_eq!(bounds.start_col, 1);
        assert_eq!(bounds.end_col, 1);
    }

    #[test]
    fn resolve_header_row_rejects_non_numeric_and_out_of_range_values() {
        let workbook = Workbook::from_sheets_for_test(vec![sheet_with_values(
            "Orders",
            &[&["order_id", "customer"], &["1001", "Alice"]],
        )]);
        let sheet = workbook.get_sheet_by_index(0).unwrap();

        let invalid = resolve_header_row(&workbook, sheet, 0, "header").unwrap_err();
        assert_eq!(invalid.code(), "invalid_query");

        let out_of_range = resolve_header_row(&workbook, sheet, 0, "9").unwrap_err();
        assert_eq!(out_of_range.code(), "invalid_query");
    }

    #[test]
    fn resolve_optional_header_row_preserves_lenient_record_resolution() {
        let workbook = Workbook::from_sheets_for_test(vec![sheet_with_values(
            "Orders",
            &[&["order_id", "customer"], &["1001", "Alice"]],
        )]);
        let sheet = workbook.get_sheet_by_index(0).unwrap();

        assert_eq!(
            resolve_optional_header_row(&workbook, sheet, 0, "header").unwrap(),
            None
        );
        assert_eq!(
            resolve_optional_header_row(&workbook, sheet, 0, "9").unwrap(),
            None
        );
        assert_eq!(
            resolve_optional_header_row(&workbook, sheet, 0, "1").unwrap(),
            Some(1)
        );
    }

    #[test]
    fn header_and_cell_helpers_preserve_existing_formula_semantics() {
        let mut sheet = sheet_with_values("Orders", &[&["order_id", ""], &["1001", "Alice"]]);
        sheet.data[1][2] = Cell {
            value: "total".to_string(),
            formula: Some("=UPPER(\"total\")".to_string()),
            is_formula: false,
            cell_type: CellType::Text,
            original_type: None,
        };
        sheet.data[2][2] = Cell {
            value: String::new(),
            formula: Some("=A2".to_string()),
            is_formula: false,
            cell_type: CellType::Text,
            original_type: None,
        };

        assert_eq!(header_value(&sheet, 1, 1), "order_id");
        assert_eq!(header_value(&sheet, 1, 2), "");
        assert!(cell_has_formula(cell_at(&sheet, 1, 2).unwrap()));
        assert!(cell_is_present(cell_at(&sheet, 2, 2)));
    }

    #[test]
    fn record_helpers_generate_stable_column_names_and_header_values() {
        let sheet = sheet_with_values(
            "Orders",
            &[
                &["order_id", "customer", "customer", ""],
                &["1001", "Alice", "VIP", "true"],
            ],
        );
        let bounds = SheetBounds {
            start_row: 1,
            end_row: 2,
            start_col: 1,
            end_col: 4,
        };

        let headers = read_header_values(&sheet, 1, bounds);
        let columns = stable_record_keys(&headers, bounds.start_col);

        assert_eq!(headers, vec!["order_id", "customer", "customer", ""]);
        assert_eq!(columns, vec!["order_id", "customer", "customer_2", "col_D"]);
    }
}
