use anyhow::{Context, Result};
use calamine::{open_workbook_auto, Data, Reader, Xls, Xlsx};
use chrono::Local;
use rust_xlsxwriter::{Format, Workbook as XlsxWorkbook};
use std::collections::HashSet;
use std::fs::File;
use std::io::BufReader;
use std::path::Path;

use crate::excel::{Cell, CellType, DataTypeInfo, Sheet};

pub enum CalamineWorkbook {
    Xlsx(Xlsx<BufReader<File>>),
    Xls(Xls<BufReader<File>>),
    None,
}

impl Clone for CalamineWorkbook {
    fn clone(&self) -> Self {
        CalamineWorkbook::None
    }
}

pub struct Workbook {
    sheets: Vec<Sheet>,
    current_sheet_index: usize,
    file_path: String,
    is_modified: bool,
    calamine_workbook: CalamineWorkbook,
    lazy_loading: bool,
    loaded_sheets: HashSet<usize>, // Track which sheets have been loaded
}

impl Clone for Workbook {
    fn clone(&self) -> Self {
        Workbook {
            sheets: self.sheets.clone(),
            current_sheet_index: self.current_sheet_index,
            file_path: self.file_path.clone(),
            is_modified: self.is_modified,
            calamine_workbook: CalamineWorkbook::None,
            lazy_loading: false,
            loaded_sheets: self.loaded_sheets.clone(),
        }
    }
}

pub fn open_workbook<P: AsRef<Path>>(path: P, enable_lazy_loading: bool) -> Result<Workbook> {
    let path_str = path.as_ref().to_string_lossy().to_string();
    let path_ref = path.as_ref();

    // Determine if the file format supports lazy loading
    let extension = path_ref
        .extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| ext.to_lowercase());

    // Only enable lazy loading if both the flag is set AND the format supports it
    let supports_lazy_loading =
        enable_lazy_loading && matches!(extension.as_deref(), Some("xlsx") | Some("xlsm"));

    // Open workbook directly from path
    let mut workbook = open_workbook_auto(&path)
        .with_context(|| format!("Unable to parse Excel file: {}", path_str))?;

    let sheet_names = workbook.sheet_names().to_vec();

    // Pre-allocate with the right capacity
    let mut sheets = Vec::with_capacity(sheet_names.len());

    // Store the original calamine workbook for lazy loading if enabled
    let mut calamine_workbook = CalamineWorkbook::None;

    if supports_lazy_loading {
        // For formats that support lazy loading, keep the original workbook
        // and only load sheet metadata
        for name in &sheet_names {
            // Create a minimal sheet with just the name
            let sheet = Sheet {
                name: name.to_string(),
                data: vec![vec![Cell::empty(); 1]; 1],
                max_rows: 0,
                max_cols: 0,
                is_loaded: false,
            };

            sheets.push(sheet);
        }

        // Try to reopen the file to get a fresh reader for lazy loading
        if let Ok(file) = File::open(&path) {
            let reader = BufReader::new(file);

            // Try to open as XLSX first
            if let Ok(xlsx_workbook) = Xlsx::new(reader) {
                calamine_workbook = CalamineWorkbook::Xlsx(xlsx_workbook);
            } else {
                // If not XLSX, try to open as XLS
                if let Ok(file) = File::open(&path) {
                    let reader = BufReader::new(file);
                    if let Ok(xls_workbook) = Xls::new(reader) {
                        calamine_workbook = CalamineWorkbook::Xls(xls_workbook);
                    }
                }
            }
        }
    } else {
        // For formats that don't support lazy loading or if lazy loading is disabled,
        for name in &sheet_names {
            let range = workbook
                .worksheet_range(name)
                .with_context(|| format!("Unable to read worksheet: {}", name))?;

            let mut sheet = create_sheet_from_range(name, range);
            sheet.is_loaded = true;
            sheets.push(sheet);
        }
    }

    if sheets.is_empty() {
        anyhow::bail!("No worksheets found in file");
    }

    let mut loaded_sheets = HashSet::new();

    if !supports_lazy_loading {
        for i in 0..sheets.len() {
            loaded_sheets.insert(i);
        }
    }

    Ok(Workbook {
        sheets,
        current_sheet_index: 0,
        file_path: path_str,
        is_modified: false,
        calamine_workbook,
        lazy_loading: supports_lazy_loading,
        loaded_sheets,
    })
}

fn create_sheet_from_range(name: &str, range: calamine::Range<Data>) -> Sheet {
    let (height, width) = range.get_size();

    // Create a data grid with empty cells, adding 1 to dimensions for 1-based indexing
    let mut data = vec![vec![Cell::empty(); width + 1]; height + 1];

    // Process only non-empty cells
    for (row_idx, col_idx, cell) in range.used_cells() {
        // Extract value, cell_type, and original_type from the Data
        let (value, cell_type, original_type) = match cell {
            Data::Empty => (String::new(), CellType::Empty, Some(DataTypeInfo::Empty)),

            Data::String(s) => {
                let value = s.clone();
                (value, CellType::Text, Some(DataTypeInfo::String))
            }

            Data::Float(f) => {
                let value = if *f == (*f as i64) as f64 && f.abs() < 1e10 {
                    (*f as i64).to_string()
                } else {
                    f.to_string()
                };
                (value, CellType::Number, Some(DataTypeInfo::Float(*f)))
            }

            Data::Int(i) => (i.to_string(), CellType::Number, Some(DataTypeInfo::Int(*i))),

            Data::Bool(b) => (
                if *b {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
                },
                CellType::Boolean,
                Some(DataTypeInfo::Bool(*b)),
            ),

            Data::Error(e) => {
                let mut value = String::with_capacity(15);
                value.push_str("Error: ");
                value.push_str(&format!("{:?}", e));
                (value, CellType::Text, Some(DataTypeInfo::Error))
            }

            Data::DateTime(dt) => (
                dt.to_string(),
                CellType::Date,
                Some(DataTypeInfo::DateTime(dt.as_f64())),
            ),

            Data::DateTimeIso(s) => {
                let value = s.clone();
                (
                    value.clone(),
                    CellType::Date,
                    Some(DataTypeInfo::DateTimeIso(value)),
                )
            }

            Data::DurationIso(s) => {
                let value = s.clone();
                (
                    value.clone(),
                    CellType::Text,
                    Some(DataTypeInfo::DurationIso(value)),
                )
            }
        };

        let is_formula = !value.is_empty() && value.starts_with('=');

        // Store the cell in data grid (using 1-based indexing)
        data[row_idx + 1][col_idx + 1] =
            Cell::new_with_type(value, is_formula, cell_type, original_type);
    }

    Sheet {
        name: name.to_string(),
        data,
        max_rows: height,
        max_cols: width,
        is_loaded: true,
    }
}

impl Workbook {
    pub fn get_current_sheet(&self) -> &Sheet {
        &self.sheets[self.current_sheet_index]
    }

    pub fn get_current_sheet_mut(&mut self) -> &mut Sheet {
        &mut self.sheets[self.current_sheet_index]
    }

    pub fn ensure_sheet_loaded(&mut self, sheet_index: usize, sheet_name: &str) -> Result<()> {
        if !self.lazy_loading || self.sheets[sheet_index].is_loaded {
            return Ok(());
        }

        // Load the sheet data from the calamine workbook
        match &mut self.calamine_workbook {
            CalamineWorkbook::Xlsx(xlsx) => {
                if let Ok(range) = xlsx.worksheet_range(sheet_name) {
                    // Replace the placeholder sheet with a fully loaded one
                    let mut sheet = create_sheet_from_range(sheet_name, range);

                    // Preserve the original name in case it was customized
                    let original_name = self.sheets[sheet_index].name.clone();
                    sheet.name = original_name;

                    self.sheets[sheet_index] = sheet;

                    // Mark the sheet as loaded
                    self.loaded_sheets.insert(sheet_index);
                }
            }
            CalamineWorkbook::Xls(xls) => {
                if let Ok(range) = xls.worksheet_range(sheet_name) {
                    // Replace the placeholder sheet with a fully loaded one
                    let mut sheet = create_sheet_from_range(sheet_name, range);

                    // Preserve the original name in case it was customized
                    let original_name = self.sheets[sheet_index].name.clone();
                    sheet.name = original_name;

                    self.sheets[sheet_index] = sheet;

                    // Mark the sheet as loaded
                    self.loaded_sheets.insert(sheet_index);
                }
            }
            CalamineWorkbook::None => {
                return Err(anyhow::anyhow!("Cannot load sheet: no workbook available"));
            }
        }

        Ok(())
    }

    pub fn get_sheet_by_index(&self, index: usize) -> Option<&Sheet> {
        self.sheets.get(index)
    }

    pub fn ensure_cell_exists(&mut self, row: usize, col: usize) {
        let sheet = &mut self.sheets[self.current_sheet_index];

        // Expand rows if needed
        if row >= sheet.data.len() {
            let default_row_len = if sheet.data.is_empty() {
                col + 1
            } else {
                sheet.data[0].len()
            };
            let rows_to_add = row + 1 - sheet.data.len();

            sheet
                .data
                .extend(vec![vec![Cell::empty(); default_row_len]; rows_to_add]);
            sheet.max_rows = sheet.max_rows.max(row);
        }

        // Expand columns if needed
        if col >= sheet.data[0].len() {
            for row_data in &mut sheet.data {
                row_data.resize_with(col + 1, Cell::empty);
            }

            sheet.max_cols = sheet.max_cols.max(col);
        }
    }

    pub fn set_cell_value(&mut self, row: usize, col: usize, value: String) -> Result<()> {
        self.ensure_cell_exists(row, col);

        let sheet = &mut self.sheets[self.current_sheet_index];
        let current_value = &sheet.data[row][col].value;

        // Only set modified flag if value actually changes
        if current_value != &value {
            let is_formula = value.starts_with('=');
            sheet.data[row][col] = Cell::new(value, is_formula);

            // Update max_cols if needed
            if col > sheet.max_cols && !sheet.data[row][col].value.is_empty() {
                sheet.max_cols = col;
            }

            self.is_modified = true;
        }

        Ok(())
    }

    pub fn get_sheet_names(&self) -> Vec<String> {
        let mut names = Vec::with_capacity(self.sheets.len());
        for sheet in &self.sheets {
            names.push(sheet.name.clone());
        }
        names
    }

    pub fn get_current_sheet_name(&self) -> String {
        self.sheets[self.current_sheet_index].name.clone()
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

        // If row is less than 1, return early with success
        if row < 1 {
            return Ok(());
        }

        // If row is outside the max range, return early with success
        if row > sheet.max_rows {
            return Ok(());
        }

        // Only remove the row if it exists in the data
        if row < sheet.data.len() {
            sheet.data.remove(row);
            self.recalculate_max_cols();
            self.is_modified = true;
        }

        Ok(())
    }

    // Delete a range of rows from the current sheet
    pub fn delete_rows(&mut self, start_row: usize, end_row: usize) -> Result<()> {
        let sheet = &mut self.sheets[self.current_sheet_index];

        // If start_row is less than 1 or start_row > end_row, return early with success
        if start_row < 1 || start_row > end_row {
            return Ok(());
        }

        // If the entire range is outside max_rows, return early with success
        if start_row > sheet.max_rows {
            return Ok(());
        }

        // Adjust end_row to not exceed the data length
        let adjusted_end_row = end_row.min(sheet.data.len() - 1);

        // If start_row is valid but end_row exceeds max_rows, adjust end_row to max_rows
        let effective_end_row = if end_row > sheet.max_rows {
            sheet.max_rows
        } else {
            adjusted_end_row
        };

        // Only proceed if there are rows to delete
        if start_row <= effective_end_row && start_row < sheet.data.len() {
            // Remove rows in reverse order to avoid index shifting issues
            for row in (start_row..=effective_end_row).rev() {
                if row < sheet.data.len() {
                    sheet.data.remove(row);
                }
            }

            self.recalculate_max_cols();
            self.is_modified = true;
        }

        Ok(())
    }

    pub fn delete_column(&mut self, col: usize) -> Result<()> {
        let sheet = &mut self.sheets[self.current_sheet_index];

        // If column is less than 1, return early with success
        if col < 1 {
            return Ok(());
        }

        // If column is outside the max range, return early with success
        if col > sheet.max_cols {
            return Ok(());
        }

        let mut has_data = false;
        for row in &sheet.data {
            if col < row.len() && !row[col].value.is_empty() {
                has_data = true;
                break;
            }
        }

        for row in sheet.data.iter_mut() {
            if col < row.len() {
                row.remove(col);
            }
        }

        self.recalculate_max_cols();
        self.recalculate_max_rows();

        if has_data {
            self.is_modified = true;
        }

        Ok(())
    }

    // Delete a range of columns from the current sheet
    pub fn delete_columns(&mut self, start_col: usize, end_col: usize) -> Result<()> {
        let sheet = &mut self.sheets[self.current_sheet_index];

        // If start_col is less than 1 or start_col > end_col, return early with success
        if start_col < 1 || start_col > end_col {
            return Ok(());
        }

        // If the entire range is outside max_cols, return early with success
        if start_col > sheet.max_cols {
            return Ok(());
        }

        // If start_col is valid but end_col exceeds max_cols, adjust end_col to max_cols
        let effective_end_col = end_col.min(sheet.max_cols);

        let mut has_data = false;
        for row in &sheet.data {
            for col in start_col..=effective_end_col {
                if col < row.len() && !row[col].value.is_empty() {
                    has_data = true;
                    break;
                }
            }
            if has_data {
                break;
            }
        }

        for row in sheet.data.iter_mut() {
            for col in (start_col..=effective_end_col).rev() {
                if col < row.len() {
                    row.remove(col);
                }
            }
        }

        self.recalculate_max_cols();
        self.recalculate_max_rows();

        if has_data {
            self.is_modified = true;
        }

        Ok(())
    }

    pub fn is_modified(&self) -> bool {
        self.is_modified
    }

    pub fn set_modified(&mut self, modified: bool) {
        self.is_modified = modified;
    }

    pub fn get_file_path(&self) -> &str {
        &self.file_path
    }

    pub fn is_lazy_loading(&self) -> bool {
        self.lazy_loading
    }

    pub fn is_sheet_loaded(&self, sheet_index: usize) -> bool {
        if !self.lazy_loading || sheet_index >= self.sheets.len() {
            return true;
        }

        self.sheets[sheet_index].is_loaded
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
            let worksheet = workbook.add_worksheet().set_name(&sheet.name)?;

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

    pub fn insert_sheet_at_index(&mut self, sheet: Sheet, index: usize) -> Result<()> {
        if index > self.sheets.len() {
            anyhow::bail!(
                "Cannot insert sheet at index {}: index out of bounds (max index: {})",
                index,
                self.sheets.len()
            );
        }
        self.sheets.insert(index, sheet);
        self.is_modified = true;
        Ok(())
    }

    pub fn recalculate_max_cols(&mut self) {
        let sheet = &mut self.sheets[self.current_sheet_index];

        // Find maximum non-empty column across all rows
        let actual_max_col = sheet
            .data
            .iter()
            .map(|row| {
                // Find last non-empty cell in this row
                row.iter()
                    .enumerate()
                    .rev()
                    .find(|(_, cell)| !cell.value.is_empty())
                    .map(|(idx, _)| idx)
                    .unwrap_or(0)
            })
            .max()
            .unwrap_or(0);

        sheet.max_cols = actual_max_col.max(1);
    }

    pub fn recalculate_max_rows(&mut self) {
        let sheet = &mut self.sheets[self.current_sheet_index];

        // Find last row with any non-empty cells
        let actual_max_row = sheet
            .data
            .iter()
            .enumerate()
            .rev()
            .find(|(_, row)| row.iter().any(|cell| !cell.value.is_empty()))
            .map(|(idx, _)| idx)
            .unwrap_or(0);

        sheet.max_rows = actual_max_row.max(1);
    }
}
