use anyhow::{Context, Result};
use calamine::{DataType, Reader, open_workbook_auto};
use chrono::Local;
use rust_xlsxwriter::{Format, Workbook as XlsxWorkbook};
use std::path::Path;

use crate::excel::{Cell, CellType, DataTypeInfo, Sheet};

#[derive(Clone)]
pub struct Workbook {
    sheets: Vec<Sheet>,
    current_sheet_index: usize,
    file_path: String,
    is_modified: bool,
}

pub fn open_workbook<P: AsRef<Path>>(path: P) -> Result<Workbook> {
    let path_str = path.as_ref().to_string_lossy().to_string();

    // Open workbook directly from path
    let mut workbook = open_workbook_auto(&path).context("Unable to parse Excel file")?;

    let sheet_names = workbook.sheet_names().to_vec();
    let mut sheets = Vec::new();

    for name in &sheet_names {
        let range = workbook
            .worksheet_range(name)
            .context(format!("Unable to read worksheet: {}", name))?;
        let sheet = create_sheet_from_range(name, range?);
        sheets.push(sheet);
    }

    if sheets.is_empty() {
        anyhow::bail!("No worksheets found in file");
    }

    Ok(Workbook {
        sheets,
        current_sheet_index: 0,
        file_path: path_str,
        is_modified: false,
    })
}

fn create_sheet_from_range(name: &str, range: calamine::Range<DataType>) -> Sheet {
    let height = range.height();
    let width = range.width();

    let mut data = vec![vec![Cell::empty(); width + 1]; height + 1];

    for (row_idx, row) in range.rows().enumerate() {
        for (col_idx, cell) in row.iter().enumerate() {
            if let DataType::Empty = cell {
                continue;
            }

            // Extract value, cell_type, and original_type from the DataType
            let (value, cell_type, original_type) = match cell {
                DataType::Empty => (String::new(), CellType::Empty, Some(DataTypeInfo::Empty)),
                DataType::String(s) => {
                    let mut value = String::with_capacity(s.len());
                    value.push_str(s);
                    (value, CellType::Text, Some(DataTypeInfo::String))
                }
                DataType::Float(f) => {
                    let value = if *f == (*f as i64) as f64 && f.abs() < 1e10 {
                        (*f as i64).to_string()
                    } else {
                        f.to_string()
                    };
                    (value, CellType::Number, Some(DataTypeInfo::Float(*f)))
                }
                DataType::Int(i) => (i.to_string(), CellType::Number, Some(DataTypeInfo::Int(*i))),
                DataType::Bool(b) => (
                    if *b {
                        "TRUE".to_string()
                    } else {
                        "FALSE".to_string()
                    },
                    CellType::Boolean,
                    Some(DataTypeInfo::Bool(*b)),
                ),
                DataType::Error(e) => {
                    // Pre-allocate with capacity for error message
                    let mut value = String::with_capacity(15);
                    value.push_str("Error: ");
                    value.push_str(&format!("{:?}", e));
                    (value, CellType::Text, Some(DataTypeInfo::Error))
                }
                DataType::DateTime(dt) => (
                    dt.to_string(),
                    CellType::Date,
                    Some(DataTypeInfo::DateTime(*dt)),
                ),
                DataType::Duration(d) => (
                    d.to_string(),
                    CellType::Text,
                    Some(DataTypeInfo::Duration(*d)),
                ),
                DataType::DateTimeIso(s) => {
                    let value = s.to_string();
                    (
                        value.clone(),
                        CellType::Date,
                        Some(DataTypeInfo::DateTimeIso(value)),
                    )
                }
                DataType::DurationIso(s) => {
                    let value = s.to_string();
                    (
                        value.clone(),
                        CellType::Text,
                        Some(DataTypeInfo::DurationIso(value)),
                    )
                }
            };

            let is_formula = !value.is_empty() && value.starts_with('=');

            data[row_idx + 1][col_idx + 1] =
                Cell::new_with_type(value, is_formula, cell_type, original_type);
        }
    }

    Sheet {
        name: name.to_string(),
        data,
        max_rows: height,
        max_cols: width,
    }
}

impl Workbook {
    pub fn get_current_sheet(&self) -> &Sheet {
        &self.sheets[self.current_sheet_index]
    }

    pub fn get_sheet_by_index(&self, index: usize) -> Option<&Sheet> {
        self.sheets.get(index)
    }

    pub fn set_cell_value(&mut self, row: usize, col: usize, value: String) -> Result<()> {
        if row >= self.sheets[self.current_sheet_index].data.len()
            || col >= self.sheets[self.current_sheet_index].data[0].len()
        {
            anyhow::bail!("Cell coordinates out of range");
        }

        let is_formula = value.starts_with('=');

        // Use Cell::new which handles type detection
        self.sheets[self.current_sheet_index].data[row][col] = Cell::new(value, is_formula);
        self.is_modified = true;

        Ok(())
    }

    pub fn get_sheet_names(&self) -> Vec<String> {
        let mut names = Vec::with_capacity(self.sheets.len());
        for sheet in &self.sheets {
            names.push(sheet.name.clone());
        }
        names
    }

    pub fn get_sheet_names_ref(&self) -> Vec<&str> {
        self.sheets.iter().map(|s| s.name.as_str()).collect()
    }

    pub fn get_current_sheet_name(&self) -> String {
        self.sheets[self.current_sheet_index].name.clone()
    }

    pub fn get_current_sheet_name_ref(&self) -> &str {
        &self.sheets[self.current_sheet_index].name
    }

    pub fn get_current_sheet_index(&self) -> usize {
        self.current_sheet_index
    }

    pub fn switch_sheet(&mut self, index: usize) -> Result<()> {
        if index >= self.sheets.len() {
            anyhow::bail!("Sheet index out of range");
        }

        self.current_sheet_index = index;
        Ok(())
    }

    pub fn delete_current_sheet(&mut self) -> Result<()> {
        // Prevent deleting the last sheet
        if self.sheets.len() <= 1 {
            anyhow::bail!("Cannot delete the last sheet");
        }

        self.sheets.remove(self.current_sheet_index);
        self.is_modified = true;

        // Adjust current_sheet_index
        if self.current_sheet_index >= self.sheets.len() {
            self.current_sheet_index = self.sheets.len() - 1;
        }

        Ok(())
    }

    pub fn delete_row(&mut self, row: usize) -> Result<()> {
        let sheet = &mut self.sheets[self.current_sheet_index];

        if row < 1 || row > sheet.max_rows {
            anyhow::bail!("Row index out of range");
        }

        sheet.data.remove(row);

        sheet.max_rows = sheet.max_rows.saturating_sub(1);

        self.is_modified = true;
        Ok(())
    }

    // Delete a range of rows from the current sheet
    pub fn delete_rows(&mut self, start_row: usize, end_row: usize) -> Result<()> {
        let sheet = &mut self.sheets[self.current_sheet_index];

        if start_row < 1
            || start_row > sheet.max_rows
            || end_row < start_row
            || end_row > sheet.max_rows
        {
            anyhow::bail!("Row range out of bounds");
        }

        let rows_to_remove = end_row - start_row + 1;

        for row in (start_row..=end_row).rev() {
            sheet.data.remove(row);
        }

        sheet.max_rows = sheet.max_rows.saturating_sub(rows_to_remove);

        self.is_modified = true;
        Ok(())
    }

    pub fn delete_column(&mut self, col: usize) -> Result<()> {
        let sheet = &mut self.sheets[self.current_sheet_index];

        if col < 1 || col > sheet.max_cols {
            anyhow::bail!("Column index out of range");
        }

        for row in sheet.data.iter_mut() {
            if col < row.len() {
                row.remove(col);
            }
        }

        sheet.max_cols = sheet.max_cols.saturating_sub(1);

        self.is_modified = true;
        Ok(())
    }

    // Delete a range of columns from the current sheet
    pub fn delete_columns(&mut self, start_col: usize, end_col: usize) -> Result<()> {
        let sheet = &mut self.sheets[self.current_sheet_index];

        if start_col < 1
            || start_col > sheet.max_cols
            || end_col < start_col
            || end_col > sheet.max_cols
        {
            anyhow::bail!("Column range out of bounds");
        }

        let cols_to_remove = end_col - start_col + 1;

        for row in sheet.data.iter_mut() {
            for col in (start_col..=end_col).rev() {
                if col < row.len() {
                    row.remove(col);
                }
            }
        }

        sheet.max_cols = sheet.max_cols.saturating_sub(cols_to_remove);

        self.is_modified = true;
        Ok(())
    }

    pub fn is_modified(&self) -> bool {
        self.is_modified
    }

    pub fn get_file_path(&self) -> &str {
        &self.file_path
    }

    pub fn save(&mut self) -> Result<()> {
        if !self.is_modified {
            println!("No changes to save.");
            return Ok(());
        }

        // Create a new workbook with rust_xlsxwriter
        let mut workbook = XlsxWorkbook::new();

        let now = Local::now();
        let timestamp = now.format("%Y%m%d_%H%M%S").to_string();
        let path = Path::new(&self.file_path);
        let file_stem = path.file_stem().and_then(|s| s.to_str()).unwrap_or("sheet");
        let extension = path.extension().and_then(|s| s.to_str()).unwrap_or("xlsx");
        let parent_dir = path.parent().unwrap_or_else(|| Path::new(""));
        let new_filename = format!("{}_{}.{}", file_stem, timestamp, extension);
        let new_filepath = parent_dir.join(new_filename);

        // Create formats
        let number_format = Format::new().set_num_format("General");
        let date_format = Format::new().set_num_format("yyyy-mm-dd");

        // Process each sheet
        for sheet in &self.sheets {
            let mut worksheet = workbook.add_worksheet().set_name(&sheet.name)?;

            // Set column widths
            for col in 0..sheet.max_cols {
                worksheet.set_column_width(col as u16, 15)?;
            }

            // Write cell data
            for row in 1..sheet.data.len() {
                if row <= sheet.max_rows {
                    for col in 1..sheet.data[0].len() {
                        if col <= sheet.max_cols {
                            let cell = &sheet.data[row][col];

                            // Skip empty cells
                            if cell.value.is_empty() {
                                continue;
                            }

                            let row_idx = (row - 1) as u32;
                            let col_idx = (col - 1) as u16;

                            // Write cell based on its type
                            match cell.cell_type {
                                CellType::Number => {
                                    if let Ok(num) = cell.value.parse::<f64>() {
                                        worksheet.write_number_with_format(
                                            row_idx,
                                            col_idx,
                                            num,
                                            &number_format,
                                        )?;
                                    } else {
                                        worksheet.write_string(row_idx, col_idx, &cell.value)?;
                                    }
                                }
                                CellType::Date => {
                                    worksheet.write_string_with_format(
                                        row_idx,
                                        col_idx,
                                        &cell.value,
                                        &date_format,
                                    )?;
                                }
                                CellType::Boolean => {
                                    if let Ok(b) = cell.value.parse::<bool>() {
                                        worksheet.write_boolean(row_idx, col_idx, b)?;
                                    } else {
                                        worksheet.write_string(row_idx, col_idx, &cell.value)?;
                                    }
                                }
                                CellType::Text => {
                                    if cell.is_formula {
                                        let formula = rust_xlsxwriter::Formula::new(&cell.value);
                                        worksheet.write_formula(row_idx, col_idx, formula)?;
                                    } else {
                                        worksheet.write_string(row_idx, col_idx, &cell.value)?;
                                    }
                                }
                                CellType::Empty => {}
                            }
                        }
                    }
                }
            }
        }

        workbook.save(&new_filepath)?;
        self.is_modified = false;

        Ok(())
    }
}
