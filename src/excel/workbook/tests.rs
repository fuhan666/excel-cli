use super::Workbook;
use crate::excel::Sheet;

fn blank_sheet(name: &str) -> Sheet {
    Sheet::blank(name.to_string())
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
