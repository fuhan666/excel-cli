use anyhow::Result;
use chrono::Local;
use rust_xlsxwriter::{Format, Workbook as XlsxWorkbook, Worksheet};
use std::path::{Path, PathBuf};

use super::Workbook;
use crate::excel::{Cell, CellType, Sheet};

impl Workbook {
    pub fn save(&mut self) -> Result<()> {
        if !self.is_modified {
            println!("No changes to save.");
            return Ok(());
        }

        self.ensure_all_sheets_loaded()?;

        let mut workbook = XlsxWorkbook::new();
        let new_filepath = timestamped_save_path(&self.file_path);
        let number_format = Format::new().set_num_format("General");
        let date_format = Format::new().set_num_format("yyyy-mm-dd");

        for sheet in &self.sheets {
            write_sheet(&mut workbook, sheet, &number_format, &date_format)?;
        }

        workbook.save(&new_filepath)?;
        self.is_modified = false;

        Ok(())
    }
}

fn timestamped_save_path(file_path: &str) -> PathBuf {
    let timestamp = Local::now().format("%Y%m%d_%H%M%S").to_string();
    let path = Path::new(file_path);
    let file_stem = path.file_stem().and_then(|s| s.to_str()).unwrap_or("sheet");
    let extension = path.extension().and_then(|s| s.to_str()).unwrap_or("xlsx");
    let parent_dir = path.parent().unwrap_or_else(|| Path::new(""));
    parent_dir.join(format!("{file_stem}_{timestamp}.{extension}"))
}

fn write_sheet(
    workbook: &mut XlsxWorkbook,
    sheet: &Sheet,
    number_format: &Format,
    date_format: &Format,
) -> Result<()> {
    let worksheet = workbook.add_worksheet().set_name(&sheet.name)?;

    if sheet.freeze_panes.is_frozen() {
        worksheet.set_freeze_panes(
            sheet.freeze_panes.rows as u32,
            sheet.freeze_panes.cols as u16,
        )?;
    }

    for col in 0..sheet.max_cols {
        worksheet.set_column_width(col as u16, 15)?;
    }

    for row in 1..sheet.data.len() {
        if row > sheet.max_rows {
            continue;
        }

        for col in 1..sheet.data[0].len() {
            if col > sheet.max_cols {
                continue;
            }

            let cell = &sheet.data[row][col];
            if cell.value.is_empty() {
                continue;
            }

            let row_idx = (row - 1) as u32;
            let col_idx = (col - 1) as u16;
            write_cell(
                worksheet,
                cell,
                row_idx,
                col_idx,
                number_format,
                date_format,
            )?;
        }
    }

    Ok(())
}

fn write_cell(
    worksheet: &mut Worksheet,
    cell: &Cell,
    row_idx: u32,
    col_idx: u16,
    number_format: &Format,
    date_format: &Format,
) -> Result<()> {
    if cell.is_formula {
        let formula_text = cell.formula.as_deref().unwrap_or(cell.value.as_str());
        let formula = rust_xlsxwriter::Formula::new(formula_text);
        worksheet.write_formula(row_idx, col_idx, formula)?;
        if !cell.value.is_empty() && cell.value != formula_text {
            worksheet.set_formula_result(row_idx, col_idx, &cell.value);
        }
        return Ok(());
    }

    match cell.cell_type {
        CellType::Number => {
            if let Ok(num) = cell.value.parse::<f64>() {
                worksheet.write_number_with_format(row_idx, col_idx, num, number_format)?;
            } else {
                worksheet.write_string(row_idx, col_idx, &cell.value)?;
            }
        }
        CellType::Date => {
            worksheet.write_string_with_format(row_idx, col_idx, &cell.value, date_format)?;
        }
        CellType::Boolean => {
            if let Ok(b) = cell.value.parse::<bool>() {
                worksheet.write_boolean(row_idx, col_idx, b)?;
            } else {
                worksheet.write_string(row_idx, col_idx, &cell.value)?;
            }
        }
        CellType::Text => {
            worksheet.write_string(row_idx, col_idx, &cell.value)?;
        }
        CellType::Empty => {}
    }

    Ok(())
}
