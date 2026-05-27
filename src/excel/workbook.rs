use anyhow::{Context, Result};
use calamine::{open_workbook_auto, Reader, Xls, Xlsx};
use std::collections::HashSet;
use std::fs::File;
use std::io::BufReader;
use std::panic::{catch_unwind, AssertUnwindSafe};
use std::path::Path;

use crate::excel::{Cell, CellType, FreezePanes, Sheet};
use crate::utils::{index_to_col_name, parse_cell_reference};

mod formula_lookup;
mod freeze_panes;
mod save;
mod sheet_parse;

use formula_lookup::lookup_formula_in_xlsx;
use freeze_panes::lookup_freeze_panes_in_xlsx;
use sheet_parse::create_sheet_from_range;

pub enum CalamineWorkbook {
    Xlsx(Box<Xlsx<BufReader<File>>>),
    Xls(Box<Xls<BufReader<File>>>),
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

    let hook = std::panic::take_hook();
    std::panic::set_hook(Box::new(|_| {}));
    let result = catch_unwind(AssertUnwindSafe(|| {
        open_workbook_impl(path.as_ref(), enable_lazy_loading)
    }));
    std::panic::set_hook(hook);

    match result {
        Ok(inner) => inner,
        Err(_) => anyhow::bail!(
            "Unable to parse Excel file: {} (parser panic: malformed workbook data)",
            path_str
        ),
    }
}

fn open_workbook_impl<P: AsRef<Path>>(path: P, enable_lazy_loading: bool) -> Result<Workbook> {
    let path_str = path.as_ref().to_string_lossy().to_string();
    let path_ref = path.as_ref();

    // Determine if the file format supports lazy loading
    let extension = path_ref
        .extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| ext.to_lowercase());

    // Only enable lazy loading if both the flag is set AND the format supports it
    let supports_lazy_loading =
        enable_lazy_loading && matches!(extension.as_deref(), Some("xlsx" | "xlsm"));

    // Open workbook directly from path
    let mut workbook = open_workbook_auto(&path)
        .with_context(|| format!("Unable to parse Excel file: {}", path_str))?;

    let sheet_names = workbook.sheet_names().to_vec();
    let freeze_panes_by_name = sheet_names
        .iter()
        .map(|name| {
            (
                name.clone(),
                lookup_freeze_panes_in_xlsx(path_ref, name).unwrap_or_default(),
            )
        })
        .collect::<std::collections::HashMap<_, _>>();

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
                freeze_panes: freeze_panes_by_name.get(name).cloned().unwrap_or_default(),
            };

            sheets.push(sheet);
        }

        // Try to reopen the file to get a fresh reader for lazy loading
        if let Ok(file) = File::open(&path) {
            let reader = BufReader::new(file);

            // Try to open as XLSX first
            if let Ok(xlsx_workbook) = Xlsx::new(reader) {
                calamine_workbook = CalamineWorkbook::Xlsx(Box::new(xlsx_workbook));
            } else {
                // If not XLSX, try to open as XLS
                if let Ok(file) = File::open(&path) {
                    let reader = BufReader::new(file);
                    if let Ok(xls_workbook) = Xls::new(reader) {
                        calamine_workbook = CalamineWorkbook::Xls(Box::new(xls_workbook));
                    }
                }
            }
        }
    } else {
        // For formats that don't support lazy loading or if lazy loading is disabled,
        for name in &sheet_names {
            let range = workbook.worksheet_range(name).with_context(|| {
                format!(
                    "Unable to parse Excel file: {} (unable to read worksheet: {})",
                    path_str, name
                )
            })?;

            let formula_range = workbook.worksheet_formula(name).ok();
            let mut sheet = create_sheet_from_range(name, range, formula_range);
            sheet.is_loaded = true;
            sheet.freeze_panes = freeze_panes_by_name.get(name).cloned().unwrap_or_default();
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

fn shrink_freeze_rows(freeze_panes: &mut FreezePanes, start_row: usize, end_row: usize) -> bool {
    if freeze_panes.rows == 0 || start_row > end_row || start_row > freeze_panes.rows {
        return false;
    }

    let affected_end = end_row.min(freeze_panes.rows);
    let deleted_frozen_rows = affected_end - start_row + 1;
    freeze_panes.rows = freeze_panes.rows.saturating_sub(deleted_frozen_rows);
    true
}

fn shrink_freeze_cols(freeze_panes: &mut FreezePanes, start_col: usize, end_col: usize) -> bool {
    if freeze_panes.cols == 0 || start_col > end_col || start_col > freeze_panes.cols {
        return false;
    }

    let affected_end = end_col.min(freeze_panes.cols);
    let deleted_frozen_cols = affected_end - start_col + 1;
    freeze_panes.cols = freeze_panes.cols.saturating_sub(deleted_frozen_cols);
    true
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
        let hook = std::panic::take_hook();
        std::panic::set_hook(Box::new(|_| {}));
        match &mut self.calamine_workbook {
            CalamineWorkbook::Xlsx(xlsx) => {
                let result = catch_unwind(AssertUnwindSafe(|| xlsx.worksheet_range(sheet_name)));
                std::panic::set_hook(hook);
                match result {
                    Ok(Ok(range)) => {
                        let formula_range = xlsx.worksheet_formula(sheet_name).ok();
                        let freeze_panes = self.sheets[sheet_index].freeze_panes.clone();
                        let mut sheet = create_sheet_from_range(sheet_name, range, formula_range);
                        let original_name = self.sheets[sheet_index].name.clone();
                        sheet.name = original_name;
                        sheet.freeze_panes = freeze_panes;
                        self.sheets[sheet_index] = sheet;
                        self.loaded_sheets.insert(sheet_index);
                    }
                    Ok(Err(_)) => {
                        // Non-panic parse error: leave sheet unloaded and continue
                    }
                    Err(_) => {
                        self.calamine_workbook = CalamineWorkbook::None;
                        return Err(anyhow::anyhow!(
                            "Unable to read worksheet '{}': parser panic: malformed workbook data",
                            sheet_name
                        ));
                    }
                }
            }
            CalamineWorkbook::Xls(xls) => {
                let result = catch_unwind(AssertUnwindSafe(|| xls.worksheet_range(sheet_name)));
                std::panic::set_hook(hook);
                match result {
                    Ok(Ok(range)) => {
                        let formula_range = xls.worksheet_formula(sheet_name).ok();
                        let freeze_panes = self.sheets[sheet_index].freeze_panes.clone();
                        let mut sheet = create_sheet_from_range(sheet_name, range, formula_range);
                        let original_name = self.sheets[sheet_index].name.clone();
                        sheet.name = original_name;
                        sheet.freeze_panes = freeze_panes;
                        self.sheets[sheet_index] = sheet;
                        self.loaded_sheets.insert(sheet_index);
                    }
                    Ok(Err(_)) => {
                        // Non-panic parse error: leave sheet unloaded and continue
                    }
                    Err(_) => {
                        self.calamine_workbook = CalamineWorkbook::None;
                        return Err(anyhow::anyhow!(
                            "Unable to read worksheet '{}': parser panic: malformed workbook data",
                            sheet_name
                        ));
                    }
                }
            }
            CalamineWorkbook::None => {
                std::panic::set_hook(hook);
                return Err(anyhow::anyhow!("Cannot load sheet: no workbook available"));
            }
        }

        Ok(())
    }

    pub fn get_sheet_by_index(&self, index: usize) -> Option<&Sheet> {
        self.sheets.get(index)
    }

    pub fn get_sheet_by_name(&self, name: &str) -> Option<&Sheet> {
        self.sheets.iter().find(|s| s.name == name)
    }

    pub(crate) fn formula_for_cell(
        &self,
        sheet_index: usize,
        sheet_name: &str,
        cell_ref: &str,
    ) -> Option<String> {
        let (row, col) = parse_cell_reference(cell_ref)?;
        let loaded_formula = self
            .sheets
            .get(sheet_index)
            .and_then(|sheet| sheet.data.get(row))
            .and_then(|cells| cells.get(col))
            .and_then(|cell| cell.formula.clone());

        loaded_formula
            .or_else(|| lookup_formula_in_xlsx(Path::new(&self.file_path), sheet_name, cell_ref))
    }

    /// Resolve a sheet specifier (name or 0-based index) to a sheet index.
    pub fn resolve_sheet(&self, spec: &str) -> Result<usize> {
        // Try parsing as 0-based index first
        if let Ok(index) = spec.parse::<usize>() {
            if index < self.sheets.len() {
                return Ok(index);
            }
        }

        // Try matching by exact name
        if let Some(index) = self.sheets.iter().position(|s| s.name == spec) {
            return Ok(index);
        }

        anyhow::bail!("Sheet '{}' not found", spec)
    }

    /// Resolve sheet by exact name.
    pub fn resolve_sheet_by_name(&self, name: &str) -> Result<usize> {
        self.sheets
            .iter()
            .position(|s| s.name == name)
            .ok_or_else(|| anyhow::anyhow!("Sheet '{}' not found", name))
    }

    /// Resolve sheet by 0-based index.
    pub fn resolve_sheet_by_index(&self, index: usize) -> Result<usize> {
        if index < self.sheets.len() {
            Ok(index)
        } else {
            anyhow::bail!(
                "Sheet index {} out of range (max: {})",
                index,
                self.sheets.len().saturating_sub(1)
            )
        }
    }

    /// Compute non-empty row count for a sheet.
    pub fn count_non_empty_rows(&self, sheet_index: usize) -> Result<usize> {
        let sheet = self
            .sheets
            .get(sheet_index)
            .ok_or_else(|| anyhow::anyhow!("Sheet index out of range"))?;

        let mut count = 0;
        for row in 1..=sheet.max_rows {
            if row < sheet.data.len() {
                let has_data = sheet.data[row].iter().any(|cell| !cell.value.is_empty());
                if has_data {
                    count += 1;
                }
            }
        }
        Ok(count)
    }

    /// Compute non-empty column count for a sheet.
    pub fn count_non_empty_cols(&self, sheet_index: usize) -> Result<usize> {
        let sheet = self
            .sheets
            .get(sheet_index)
            .ok_or_else(|| anyhow::anyhow!("Sheet index out of range"))?;

        let mut count = 0;
        for col in 1..=sheet.max_cols {
            let has_data = sheet.data.iter().any(|row| {
                if col < row.len() {
                    !row[col].value.is_empty()
                } else {
                    false
                }
            });
            if has_data {
                count += 1;
            }
        }
        Ok(count)
    }

    /// Find header row candidates for a sheet.
    /// Returns (candidates[], recommended_header_row).
    pub fn find_header_candidates(
        &self,
        sheet_index: usize,
    ) -> Result<(Vec<usize>, Option<usize>)> {
        let sheet = self
            .sheets
            .get(sheet_index)
            .ok_or_else(|| anyhow::anyhow!("Sheet index out of range"))?;

        if sheet.max_rows == 0 {
            return Ok((vec![], None));
        }

        let mut candidates = Vec::new();
        let max_scan = sheet.max_rows.min(20);

        for row in 1..=max_scan {
            if row >= sheet.data.len() {
                break;
            }
            let row_data = &sheet.data[row];
            let non_empty_count = row_data
                .iter()
                .take(sheet.max_cols + 1)
                .filter(|c| !c.value.is_empty())
                .count();
            let text_count = row_data
                .iter()
                .take(sheet.max_cols + 1)
                .filter(|c| {
                    !c.value.is_empty()
                        && (c.cell_type == CellType::Text || c.cell_type == CellType::Boolean)
                })
                .count();
            let total_cols = sheet.max_cols.max(1);

            let non_empty_ratio = non_empty_count as f64 / total_cols as f64;
            let text_ratio = if non_empty_count > 0 {
                text_count as f64 / non_empty_count as f64
            } else {
                0.0
            };

            // A good header row has decent coverage and mostly text values
            if non_empty_ratio >= 0.3 && text_ratio >= 0.5 {
                candidates.push(row);
            }
        }

        let recommended = candidates.first().copied();
        Ok((candidates, recommended))
    }

    /// Get the used range of a sheet in A1 notation (e.g., "A1:H2048").
    /// Returns an empty string if the sheet has no data.
    pub fn get_used_range(&self, sheet_index: usize) -> Result<String> {
        let sheet = self
            .sheets
            .get(sheet_index)
            .ok_or_else(|| anyhow::anyhow!("Sheet index out of range"))?;

        if sheet.max_rows == 0 || sheet.max_cols == 0 {
            return Ok(String::new());
        }

        let end_col = index_to_col_name(sheet.max_cols);
        Ok(format!("A1:{}{}", end_col, sheet.max_rows))
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

    pub fn set_freeze_panes(&mut self, rows: usize, cols: usize) {
        let sheet = &mut self.sheets[self.current_sheet_index];

        if sheet.freeze_panes.rows != rows || sheet.freeze_panes.cols != cols {
            sheet.freeze_panes = FreezePanes { rows, cols };
            self.is_modified = true;
        }
    }

    pub fn clear_freeze_panes(&mut self) {
        self.set_freeze_panes(0, 0);
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

    pub fn add_sheet(&mut self, name: &str, index: usize) -> Result<String> {
        let sheet_name = name.trim();

        self.validate_sheet_name(sheet_name)?;
        self.insert_sheet_at_index(Sheet::blank(sheet_name.to_string()), index)?;

        Ok(sheet_name.to_string())
    }

    pub fn delete_current_sheet(&mut self) -> Result<()> {
        self.delete_sheet_at_index(self.current_sheet_index)
    }

    pub fn delete_sheet_at_index(&mut self, index: usize) -> Result<()> {
        // Prevent deleting the last sheet
        if self.sheets.len() <= 1 {
            anyhow::bail!("Cannot delete the last sheet");
        }

        if index >= self.sheets.len() {
            anyhow::bail!("Sheet index out of range");
        }

        self.sheets.remove(index);
        self.is_modified = true;

        if index < self.current_sheet_index {
            self.current_sheet_index = self.current_sheet_index.saturating_sub(1);
        } else if self.current_sheet_index >= self.sheets.len() {
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

        let freeze_changed = shrink_freeze_rows(&mut sheet.freeze_panes, row, row);

        // Only remove the row if it exists in the data
        if row < sheet.data.len() {
            sheet.data.remove(row);
            self.recalculate_max_cols();
            self.is_modified = true;
        }

        if freeze_changed {
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

        let freeze_changed =
            shrink_freeze_rows(&mut sheet.freeze_panes, start_row, effective_end_row);

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

        if freeze_changed {
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

        let freeze_changed = shrink_freeze_cols(&mut sheet.freeze_panes, col, col);
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

        if has_data || freeze_changed {
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

        let freeze_changed =
            shrink_freeze_cols(&mut sheet.freeze_panes, start_col, effective_end_col);
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

        if has_data || freeze_changed {
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

    pub fn insert_sheet_at_index(&mut self, sheet: Sheet, index: usize) -> Result<()> {
        if index > self.sheets.len() {
            anyhow::bail!(
                "Cannot insert sheet at index {}: index out of bounds (max index: {})",
                index,
                self.sheets.len()
            );
        }

        if index <= self.current_sheet_index {
            self.current_sheet_index += 1;
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

    fn ensure_all_sheets_loaded(&mut self) -> Result<()> {
        if !self.lazy_loading {
            return Ok(());
        }

        let pending_sheets: Vec<(usize, String)> = self
            .sheets
            .iter()
            .enumerate()
            .filter(|(_, sheet)| !sheet.is_loaded)
            .map(|(index, sheet)| (index, sheet.name.clone()))
            .collect();

        for (index, name) in pending_sheets {
            self.ensure_sheet_loaded(index, &name)?;
        }

        Ok(())
    }

    fn validate_sheet_name(&self, name: &str) -> Result<()> {
        if name.is_empty() {
            anyhow::bail!("Sheet name cannot be empty");
        }

        if name.chars().count() > 31 {
            anyhow::bail!("Sheet name cannot exceed 31 characters");
        }

        if name.starts_with('\'') || name.ends_with('\'') {
            anyhow::bail!("Sheet name cannot start or end with apostrophes");
        }

        if name
            .chars()
            .any(|c| matches!(c, '[' | ']' | ':' | '*' | '?' | '/' | '\\'))
        {
            anyhow::bail!("Sheet name cannot contain any of these characters: [ ] : * ? / \\");
        }

        if self
            .sheets
            .iter()
            .any(|sheet| sheet.name.eq_ignore_ascii_case(name))
        {
            anyhow::bail!("Sheet '{}' already exists", name);
        }

        Ok(())
    }

    #[cfg(test)]
    pub(crate) fn from_sheets_for_test(sheets: Vec<Sheet>) -> Self {
        let loaded_sheets = (0..sheets.len()).collect();

        Self {
            sheets,
            current_sheet_index: 0,
            file_path: "test.xlsx".to_string(),
            is_modified: false,
            calamine_workbook: CalamineWorkbook::None,
            lazy_loading: false,
            loaded_sheets,
        }
    }
}

#[cfg(test)]
mod tests;
