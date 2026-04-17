use std::process::Command;
use std::path::PathBuf;

fn excel_cli_bin() -> PathBuf {
    let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    path.push("target");
    path.push("debug");
    path.push("excel-cli");
    path
}

fn create_test_workbook(path: &std::path::Path) {
    use rust_xlsxwriter::Workbook as XlsxWorkbook;

    let mut workbook = XlsxWorkbook::new();

    // Sheet 1: Summary
    let sheet1 = workbook.add_worksheet();
    sheet1.set_name("Summary").unwrap();
    sheet1.write_string(0, 0, "Total").unwrap();
    sheet1.write_number(0, 1, 1234.5).unwrap();

    // Sheet 2: Orders
    let sheet2 = workbook.add_worksheet();
    sheet2.set_name("Orders").unwrap();
    sheet2.write_string(0, 0, "order_id").unwrap();
    sheet2.write_string(0, 1, "customer").unwrap();
    sheet2.write_string(1, 0, "1001").unwrap();
    sheet2.write_string(1, 1, "Alice").unwrap();

    // Sheet 3: 客户 (non-ASCII name)
    let sheet3 = workbook.add_worksheet();
    sheet3.set_name("客户").unwrap();
    sheet3.write_string(0, 0, "姓名").unwrap();
    sheet3.write_string(1, 0, "张三").unwrap();

    // Empty sheet
    let sheet4 = workbook.add_worksheet();
    sheet4.set_name("EmptySheet").unwrap();

    workbook.save(path).unwrap();
}

#[test]
fn test_sheets_text_output() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_sheets.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg(&file_path)
        .arg("--sheets")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("0\tSummary"), "Expected Summary sheet, got: {}", stdout);
    assert!(stdout.contains("1\tOrders"), "Expected Orders sheet, got: {}", stdout);
    assert!(stdout.contains("2\t客户"), "Expected 客户 sheet, got: {}", stdout);
    assert!(stdout.contains("3\tEmptySheet"), "Expected EmptySheet, got: {}", stdout);
}

#[test]
fn test_sheets_json_output() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_sheets_json.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg(&file_path)
        .arg("--sheets")
        .arg("--format")
        .arg("json")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("\"index\": 0"));
    assert!(stdout.contains("\"name\": \"Summary\""));
    assert!(stdout.contains("\"name\": \"客户\""));
}

#[test]
fn test_sheet_info_by_name_text() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_sheet_name.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg(&file_path)
        .arg("--sheet")
        .arg("Orders")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("name\tOrders"), "Expected Orders name, got: {}", stdout);
    assert!(stdout.contains("index\t1"), "Expected index 1, got: {}", stdout);
    assert!(stdout.contains("used_range\t"), "Expected used_range, got: {}", stdout);
}

#[test]
fn test_sheet_info_by_index_text() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_sheet_index.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg(&file_path)
        .arg("--sheet")
        .arg("0")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("name\tSummary"), "Expected Summary, got: {}", stdout);
    assert!(stdout.contains("index\t0"), "Expected index 0, got: {}", stdout);
}

#[test]
fn test_sheet_info_json() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_sheet_json.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg(&file_path)
        .arg("--sheet")
        .arg("Orders")
        .arg("--format")
        .arg("json")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let stdout = String::from_utf8_lossy(&output.stdout);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("Invalid JSON output");
    assert_eq!(json["name"], "Orders");
    assert_eq!(json["index"], 1);
    assert!(json["used_range"].as_str().unwrap().len() > 0);
    assert_eq!(json["max_rows"], 2);
    assert_eq!(json["max_cols"], 2);
}

#[test]
fn test_sheet_info_non_ascii_name() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_non_ascii.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg(&file_path)
        .arg("--sheet")
        .arg("客户")
        .arg("--format")
        .arg("json")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let stdout = String::from_utf8_lossy(&output.stdout);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("Invalid JSON output");
    assert_eq!(json["name"], "客户");
    assert_eq!(json["index"], 2);
}

#[test]
fn test_sheet_info_empty_sheet() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_empty.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg(&file_path)
        .arg("--sheet")
        .arg("EmptySheet")
        .arg("--format")
        .arg("json")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let stdout = String::from_utf8_lossy(&output.stdout);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("Invalid JSON output");
    assert_eq!(json["name"], "EmptySheet");
    assert_eq!(json["used_range"], "");
}

#[test]
fn test_sheet_not_found_exits_nonzero() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_not_found.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg(&file_path)
        .arg("--sheet")
        .arg("NonExistent")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(!output.status.success(), "Expected failure for non-existent sheet");
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains("Sheet 'NonExistent' not found"), "Expected meaningful error, got: {}", stderr);
}

#[test]
fn test_sheet_index_out_of_range_exits_nonzero() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_idx_range.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg(&file_path)
        .arg("--sheet")
        .arg("99")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(!output.status.success(), "Expected failure for out-of-range index");
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains("Sheet '99' not found"), "Expected meaningful error, got: {}", stderr);
}

#[test]
fn test_sheets_and_json_export_mutually_exclusive() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_mutex.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg(&file_path)
        .arg("--sheets")
        .arg("--json-export")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(!output.status.success(), "Expected failure for mutually exclusive flags");
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains("mutually exclusive"), "Expected mutual exclusion error, got: {}", stderr);
}
