use std::path::{Path, PathBuf};

use super::{open_workbook, Workbook};
use crate::excel::Sheet;

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
