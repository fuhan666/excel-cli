use std::fs::File;
use std::io::Read;
use std::path::{Path, PathBuf};

use super::{open_workbook, Workbook};
use crate::excel::{Cell, FreezePanes, Sheet};

fn blank_sheet(name: &str) -> Sheet {
    Sheet::blank(name.to_string())
}

fn temp_path(name: &str) -> PathBuf {
    std::env::temp_dir().join(name)
}

fn create_formula_workbook(path: &Path) {
    use rust_xlsxwriter::Workbook as XlsxWorkbook;

    let mut workbook = XlsxWorkbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("TypedCells").unwrap();
    sheet.write_string(0, 0, "text_value").unwrap();
    sheet.write_string(0, 1, "number_value").unwrap();
    sheet.write_string(0, 2, "date_value").unwrap();
    sheet.write_string(0, 3, "boolean_value").unwrap();
    sheet.write_string(0, 4, "formula_value").unwrap();
    sheet.write_string(1, 0, "hello").unwrap();
    sheet.write_number(1, 1, 42.5).unwrap();
    sheet.write_boolean(1, 3, true).unwrap();
    sheet.write_formula(1, 4, "=B2*2").unwrap();
    sheet.set_formula_result(1, 4, "85");
    workbook.save(path).unwrap();
}

fn create_freeze_workbook(path: &Path) {
    use rust_xlsxwriter::Workbook as XlsxWorkbook;

    let mut workbook = XlsxWorkbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Frozen").unwrap();
    sheet.set_freeze_panes(1, 1).unwrap();
    sheet.write_string(0, 0, "name").unwrap();
    sheet.write_string(1, 1, "Ada").unwrap();
    workbook.save(path).unwrap();
}

fn worksheet_xml(path: &Path, sheet_entry: &str) -> String {
    let archive_file = File::open(path).unwrap();
    let mut archive = zip::ZipArchive::new(archive_file).unwrap();
    let mut entry = archive.by_name(sheet_entry).unwrap();
    let mut xml = String::new();
    entry.read_to_string(&mut xml).unwrap();
    xml
}

fn remove_temp_outputs(prefix: &str) {
    for entry in std::fs::read_dir(std::env::temp_dir()).unwrap() {
        let entry = entry.unwrap();
        let file_name = entry.file_name();
        let file_name = file_name.to_string_lossy();
        if file_name.starts_with(prefix) && file_name.ends_with(".xlsx") {
            let _ = std::fs::remove_file(entry.path());
        }
    }
}

fn find_temp_output(prefix: &str) -> PathBuf {
    std::fs::read_dir(std::env::temp_dir())
        .unwrap()
        .filter_map(Result::ok)
        .map(|entry| entry.path())
        .find(|path| {
            path.file_name()
                .and_then(|name| name.to_str())
                .is_some_and(|name| name.starts_with(prefix) && name.ends_with(".xlsx"))
        })
        .unwrap_or_else(|| panic!("expected saved workbook with prefix {prefix}"))
}

#[test]
fn adds_blank_sheet_after_current_sheet() {
    let mut workbook =
        Workbook::from_sheets_for_test(vec![blank_sheet("Sheet1"), blank_sheet("Sheet2")]);

    let sheet_name = workbook.add_sheet("Added", 1).unwrap();

    assert_eq!(sheet_name, "Added");
    assert_eq!(
        workbook.get_sheet_names(),
        vec!["Sheet1", "Added", "Sheet2"]
    );

    let added_sheet = workbook.get_sheet_by_index(1).unwrap();
    assert_eq!(added_sheet.name, "Added");
    assert_eq!(added_sheet.max_rows, 1);
    assert_eq!(added_sheet.max_cols, 1);
    assert!(added_sheet.is_loaded);
    assert_eq!(added_sheet.data.len(), 2);
    assert_eq!(added_sheet.data[1].len(), 2);
}

#[test]
fn rejects_duplicate_sheet_names_case_insensitively() {
    let mut workbook = Workbook::from_sheets_for_test(vec![blank_sheet("Summary")]);

    let error = workbook.add_sheet("summary", 1).unwrap_err().to_string();

    assert!(error.contains("already exists"));
}

#[test]
fn rejects_invalid_sheet_names() {
    let mut workbook = Workbook::from_sheets_for_test(vec![blank_sheet("Sheet1")]);

    assert!(workbook.add_sheet("", 1).is_err());
    assert!(workbook.add_sheet("Bad/Name", 1).is_err());
    assert!(workbook.add_sheet("'quoted", 1).is_err());
    assert!(workbook
        .add_sheet("this-sheet-name-is-definitely-too-long", 1)
        .is_err());
}

#[test]
fn counts_sheet_name_length_by_characters() {
    let mut workbook = Workbook::from_sheets_for_test(vec![blank_sheet("Sheet1")]);
    let valid_name = "表".repeat(31);
    let invalid_name = "表".repeat(32);

    assert!(workbook.add_sheet(&valid_name, 1).is_ok());
    assert!(workbook.add_sheet(&invalid_name, 2).is_err());
}

#[test]
fn resolves_sheet_by_index_and_name() {
    let workbook = Workbook::from_sheets_for_test(vec![
        blank_sheet("Sheet1"),
        blank_sheet("Orders"),
        blank_sheet("客户"),
    ]);

    assert_eq!(workbook.resolve_sheet("0").unwrap(), 0);
    assert_eq!(workbook.resolve_sheet("2").unwrap(), 2);
    assert_eq!(workbook.resolve_sheet("Sheet1").unwrap(), 0);
    assert_eq!(workbook.resolve_sheet("Orders").unwrap(), 1);
    assert_eq!(workbook.resolve_sheet("客户").unwrap(), 2);

    assert!(workbook.resolve_sheet("99").is_err());
    assert!(workbook.resolve_sheet("Missing").is_err());
}

#[test]
fn computes_used_range_for_sheet() {
    let mut sheet = Sheet::blank("Test".to_string());
    sheet.max_rows = 10;
    sheet.max_cols = 5;
    let workbook = Workbook::from_sheets_for_test(vec![sheet]);

    assert_eq!(workbook.get_used_range(0).unwrap(), "A1:E10");
    assert!(workbook.get_used_range(99).is_err());
}

#[test]
fn empty_sheet_has_no_used_range() {
    let mut sheet = Sheet::blank("Empty".to_string());
    sheet.max_rows = 0;
    sheet.max_cols = 0;
    let workbook = Workbook::from_sheets_for_test(vec![sheet]);
    assert_eq!(workbook.get_used_range(0).unwrap(), "");
}

#[test]
fn formula_for_cell_falls_back_to_xlsx_archive_metadata() {
    let path = temp_path("excel_cli_workbook_formula_lookup.xlsx");
    create_formula_workbook(&path);

    let workbook = open_workbook(&path, true).unwrap();
    let sheet_index = workbook.resolve_sheet_by_name("TypedCells").unwrap();

    let sheet = workbook.get_sheet_by_index(sheet_index).unwrap();
    assert!(!sheet.is_loaded);
    assert_eq!(
        workbook.formula_for_cell(sheet_index, "TypedCells", "E2"),
        Some("=B2*2".to_string())
    );
}

#[test]
fn open_workbook_restores_xlsx_freeze_panes_metadata() {
    let path = temp_path("excel_cli_workbook_freeze_lookup.xlsx");
    create_freeze_workbook(&path);

    let workbook = open_workbook(&path, false).unwrap();
    let sheet = workbook.get_current_sheet();

    assert_eq!(sheet.freeze_panes.rows, 1);
    assert_eq!(sheet.freeze_panes.cols, 1);
    assert_eq!(sheet.freeze_panes.split_cell_ref(), "B2");
}

#[test]
fn lazy_loaded_sheet_preserves_freeze_panes_after_loading() {
    let path = temp_path("excel_cli_workbook_lazy_freeze_lookup.xlsx");
    create_freeze_workbook(&path);

    let mut workbook = open_workbook(&path, true).unwrap();
    let sheet = workbook.get_current_sheet();
    assert!(!sheet.is_loaded);
    assert_eq!(sheet.freeze_panes.split_cell_ref(), "B2");

    workbook.ensure_sheet_loaded(0, "Frozen").unwrap();
    let sheet = workbook.get_current_sheet();
    assert!(sheet.is_loaded);
    assert_eq!(sheet.freeze_panes.split_cell_ref(), "B2");
}

#[test]
fn save_writes_freeze_panes_to_xlsx_xml() {
    let prefix = "excel_cli_freeze_save_";
    remove_temp_outputs(prefix);

    let mut sheet = Sheet::blank("Frozen".to_string());
    sheet.freeze_panes = FreezePanes { rows: 1, cols: 1 };
    let mut workbook = Workbook::from_sheets_for_test(vec![sheet]);
    workbook.file_path = temp_path(&format!("{prefix}source.xlsx"))
        .to_string_lossy()
        .to_string();
    workbook.set_modified(true);

    workbook.save().unwrap();

    let saved_path = find_temp_output(prefix);
    let xml = worksheet_xml(&saved_path, "xl/worksheets/sheet1.xml");
    assert!(xml.contains(r#"xSplit="1""#), "{xml}");
    assert!(xml.contains(r#"ySplit="1""#), "{xml}");
    assert!(xml.contains(r#"topLeftCell="B2""#), "{xml}");
    assert!(xml.contains(r#"state="frozen""#), "{xml}");
}

#[test]
fn deleting_rows_and_columns_shrinks_freeze_panes() {
    let mut sheet = Sheet::blank("Frozen".to_string());
    sheet.data = vec![vec![Cell::empty(); 5]; 5];
    sheet.data[4][4] = Cell::new("keep bounds".to_string(), false);
    sheet.max_rows = 4;
    sheet.max_cols = 4;
    sheet.freeze_panes = FreezePanes { rows: 2, cols: 2 };
    let mut workbook = Workbook::from_sheets_for_test(vec![sheet]);

    workbook.delete_row(1).unwrap();
    assert_eq!(workbook.get_current_sheet().freeze_panes.rows, 1);
    assert_eq!(workbook.get_current_sheet().freeze_panes.cols, 2);

    workbook.delete_column(2).unwrap();
    assert_eq!(workbook.get_current_sheet().freeze_panes.rows, 1);
    assert_eq!(workbook.get_current_sheet().freeze_panes.cols, 1);

    workbook.delete_row(10).unwrap();
    workbook.delete_column(10).unwrap();
    assert_eq!(workbook.get_current_sheet().freeze_panes.rows, 1);
    assert_eq!(workbook.get_current_sheet().freeze_panes.cols, 1);
}
