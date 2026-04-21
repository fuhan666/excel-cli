use std::path::PathBuf;
use std::process::Command;

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

fn create_read_contract_workbook(path: &std::path::Path) {
    use rust_xlsxwriter::{ExcelDateTime, Format, Workbook as XlsxWorkbook};

    let mut workbook = XlsxWorkbook::new();

    let typed_sheet = workbook.add_worksheet();
    typed_sheet.set_name("TypedCells").unwrap();
    typed_sheet.write_string(0, 0, "text_value").unwrap();
    typed_sheet.write_string(0, 1, "number_value").unwrap();
    typed_sheet.write_string(0, 2, "date_value").unwrap();
    typed_sheet.write_string(0, 3, "boolean_value").unwrap();
    typed_sheet.write_string(0, 4, "formula_value").unwrap();
    typed_sheet.write_string(0, 5, "empty_value").unwrap();
    typed_sheet.write_string(1, 0, "hello").unwrap();
    typed_sheet.write_number(1, 1, 42.5).unwrap();
    let date_format = Format::new().set_num_format("yyyy-mm-dd");
    let date = ExcelDateTime::from_ymd(2024, 2, 3).unwrap();
    typed_sheet
        .write_datetime_with_format(1, 2, &date, &date_format)
        .unwrap();
    typed_sheet.write_boolean(1, 3, true).unwrap();
    typed_sheet.write_formula(1, 4, "=B2*2").unwrap();
    typed_sheet.set_formula_result(1, 4, "85");

    let rows_sheet = workbook.add_worksheet();
    rows_sheet.set_name("RowCases").unwrap();
    rows_sheet.write_string(0, 0, "Quarterly export").unwrap();
    rows_sheet.write_string(1, 0, "order_id").unwrap();
    rows_sheet.write_string(1, 1, "customer").unwrap();
    rows_sheet.write_string(1, 2, "customer").unwrap();
    rows_sheet.write_string(2, 0, "1001").unwrap();
    rows_sheet.write_string(2, 1, "Alice").unwrap();
    rows_sheet.write_string(2, 2, "VIP").unwrap();
    rows_sheet.write_boolean(2, 3, true).unwrap();
    rows_sheet.write_string(3, 0, "1002").unwrap();
    rows_sheet.write_string(3, 1, "Bob").unwrap();
    rows_sheet.write_string(3, 2, "Standard").unwrap();
    rows_sheet.write_boolean(3, 3, false).unwrap();

    workbook.save(path).unwrap();
}

fn create_columns_contract_workbook(path: &std::path::Path) {
    use rust_xlsxwriter::Workbook as XlsxWorkbook;

    let mut workbook = XlsxWorkbook::new();

    let sheet = workbook.add_worksheet();
    sheet.set_name("ColumnCases").unwrap();
    sheet.write_string(0, 0, "customer").unwrap();
    sheet.write_string(0, 1, "customer").unwrap();
    sheet.write_string(0, 2, "").unwrap();
    sheet.write_string(0, 3, "数量").unwrap();
    sheet.write_string(0, 4, "mixed").unwrap();
    sheet.write_string(0, 5, "active").unwrap();
    sheet.write_string(0, 6, "total_formula").unwrap();
    sheet.write_string(1, 0, "Alice").unwrap();
    sheet.write_string(1, 1, "VIP").unwrap();
    sheet.write_string(1, 2, "needs safe name").unwrap();
    sheet.write_number(1, 3, 3.0).unwrap();
    sheet.write_number(1, 4, 10.0).unwrap();
    sheet.write_boolean(1, 5, true).unwrap();
    sheet.write_formula(1, 6, "=D2*2").unwrap();
    sheet.set_formula_result(1, 6, "6");
    sheet.write_string(2, 0, "Bob").unwrap();
    sheet.write_string(2, 1, "Standard").unwrap();
    sheet.write_number(2, 3, 4.0).unwrap();
    sheet.write_string(2, 4, "later").unwrap();
    sheet.write_boolean(2, 5, false).unwrap();
    sheet.write_formula(2, 6, "=D3*2").unwrap();
    sheet.set_formula_result(2, 6, "8");

    let offset = workbook.add_worksheet();
    offset.set_name("OffsetHeader").unwrap();
    offset.write_string(0, 0, "Quarterly export").unwrap();
    offset.write_string(1, 0, "item").unwrap();
    offset.write_string(1, 1, "amount").unwrap();
    offset.write_string(2, 0, "Widget").unwrap();
    offset.write_number(2, 1, 12.0).unwrap();

    workbook.save(path).unwrap();
}

fn assert_json_success(output: &std::process::Output) -> serde_json::Value {
    assert!(
        output.status.success(),
        "Expected success. stderr: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    assert!(
        output.stderr.is_empty(),
        "Expected empty stderr on success. stderr: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    serde_json::from_slice(&output.stdout).expect("Expected valid JSON output")
}

fn assert_json_error(output: &std::process::Output, expected_exit_code: i32) {
    assert!(!output.status.success(), "Expected failure but got success");
    assert!(
        output.stdout.is_empty(),
        "Expected empty stdout on error. stdout: {}",
        String::from_utf8_lossy(&output.stdout)
    );
    let actual = output.status.code().unwrap_or(-1);
    assert_eq!(
        actual,
        expected_exit_code,
        "Expected exit code {}, got {}. stderr: {}",
        expected_exit_code,
        actual,
        String::from_utf8_lossy(&output.stderr)
    );
    // Error should be valid JSON on stderr
    let err_json: serde_json::Value =
        serde_json::from_slice(&output.stderr).expect("Expected valid JSON error on stderr");
    assert!(
        err_json["error"].is_object(),
        "Error envelope missing error field"
    );
    assert!(
        err_json["error"]["code"].is_string(),
        "Error code should be a string"
    );
    assert!(
        err_json["error"]["message"].is_string(),
        "Error message should be a string"
    );
}

#[test]
fn test_inspect_workbook_json() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_inspect_workbook.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("inspect")
        .arg("workbook")
        .arg(&file_path)
        .output()
        .expect("Failed to execute excel-cli");

    let json = assert_json_success(&output);
    assert_eq!(json["schema_version"], "1.0");
    assert_eq!(json["command"], "inspect.workbook");
    assert!(json["file"]["path"]
        .as_str()
        .unwrap()
        .contains("excel_cli_test_inspect_workbook"));
    assert_eq!(json["file"]["format"], "xlsx");
    assert_eq!(json["data"]["sheet_count"], 4);

    let sheets = json["data"]["sheets"].as_array().unwrap();
    assert_eq!(sheets.len(), 4);
    assert_eq!(sheets[0]["name"], "Summary");
    assert_eq!(sheets[0]["index"], 0);
    assert_eq!(sheets[1]["name"], "Orders");
    assert_eq!(sheets[2]["name"], "客户");
    assert_eq!(sheets[3]["name"], "EmptySheet");
    assert_eq!(sheets[3]["is_empty"], true);
}

#[test]
fn test_inspect_workbook_text() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_inspect_wb_text.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("inspect")
        .arg("workbook")
        .arg(&file_path)
        .arg("--format")
        .arg("text")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(output.status.success());
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("0\tSummary"));
    assert!(stdout.contains("1\tOrders"));
    assert!(stdout.contains("2\t客户"));
    assert!(stdout.contains("3\tEmptySheet"));
}

#[test]
fn test_inspect_sheet_by_name_json() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_inspect_sheet_name.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("inspect")
        .arg("sheet")
        .arg(&file_path)
        .arg("--sheet")
        .arg("Orders")
        .output()
        .expect("Failed to execute excel-cli");

    let json = assert_json_success(&output);
    assert_eq!(json["command"], "inspect.sheet");
    assert_eq!(json["target"]["sheet"], "Orders");
    assert_eq!(json["target"]["sheet_index"], 1);
    assert_eq!(json["data"]["name"], "Orders");
    assert_eq!(json["data"]["index"], 1);
    assert_eq!(json["data"]["max_rows"], 2);
    assert_eq!(json["data"]["max_cols"], 2);
    assert!(!json["data"]["used_range"].as_str().unwrap().is_empty());
    assert!(json["data"]["non_empty_rows"].is_number());
    assert!(json["data"]["non_empty_cols"].is_number());
    assert!(json["data"]["header_candidates"].is_array());
}

#[test]
fn test_inspect_sheet_by_index_json() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_inspect_sheet_idx.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("inspect")
        .arg("sheet")
        .arg(&file_path)
        .arg("--sheet-index")
        .arg("0")
        .output()
        .expect("Failed to execute excel-cli");

    let json = assert_json_success(&output);
    assert_eq!(json["target"]["sheet"], "Summary");
    assert_eq!(json["target"]["sheet_index"], 0);
    assert_eq!(json["data"]["name"], "Summary");
}

#[test]
fn test_inspect_sheet_non_ascii_name() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_inspect_non_ascii.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("inspect")
        .arg("sheet")
        .arg(&file_path)
        .arg("--sheet")
        .arg("客户")
        .output()
        .expect("Failed to execute excel-cli");

    let json = assert_json_success(&output);
    assert_eq!(json["data"]["name"], "客户");
    assert_eq!(json["target"]["sheet_index"], 2);
}

#[test]
fn test_inspect_sheet_empty_sheet() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_inspect_empty.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("inspect")
        .arg("sheet")
        .arg(&file_path)
        .arg("--sheet")
        .arg("EmptySheet")
        .output()
        .expect("Failed to execute excel-cli");

    let json = assert_json_success(&output);
    assert_eq!(json["data"]["name"], "EmptySheet");
    assert_eq!(json["data"]["used_range"], "");
}

#[test]
fn test_inspect_sheet_not_found() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_sheet_not_found.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("inspect")
        .arg("sheet")
        .arg(&file_path)
        .arg("--sheet")
        .arg("NonExistent")
        .output()
        .expect("Failed to execute excel-cli");

    assert_json_error(&output, 5); // EXIT_TARGET_NOT_FOUND
}

#[test]
fn test_inspect_sheet_index_out_of_range() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_sheet_idx_range.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("inspect")
        .arg("sheet")
        .arg(&file_path)
        .arg("--sheet-index")
        .arg("99")
        .output()
        .expect("Failed to execute excel-cli");

    assert_json_error(&output, 5); // EXIT_TARGET_NOT_FOUND
}

#[test]
fn test_read_cell_json() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_read_cell.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("read")
        .arg("cell")
        .arg(&file_path)
        .arg("--sheet")
        .arg("Orders")
        .arg("--cell")
        .arg("B2")
        .output()
        .expect("Failed to execute excel-cli");

    let json = assert_json_success(&output);
    assert_eq!(json["command"], "read.cell");
    assert_eq!(json["target"]["sheet"], "Orders");
    assert_eq!(json["target"]["cell"], "B2");
    assert_eq!(json["data"]["value"], "Alice");
    assert_eq!(json["data"]["type"], "text");
}

#[test]
fn test_read_cell_text() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_read_cell_text.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("read")
        .arg("cell")
        .arg(&file_path)
        .arg("--sheet")
        .arg("Orders")
        .arg("--cell")
        .arg("B2")
        .arg("--format")
        .arg("text")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(output.status.success());
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert_eq!(stdout.trim(), "Alice");
}

#[test]
fn test_read_cell_non_ascii_sheet() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_cell_non_ascii.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("read")
        .arg("cell")
        .arg(&file_path)
        .arg("--sheet")
        .arg("客户")
        .arg("--cell")
        .arg("A2")
        .arg("--format")
        .arg("text")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(output.status.success());
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert_eq!(stdout.trim(), "张三");
}

#[test]
fn test_read_cell_invalid_reference() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_cell_bad.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("read")
        .arg("cell")
        .arg(&file_path)
        .arg("--sheet")
        .arg("Orders")
        .arg("--cell")
        .arg("BAD")
        .output()
        .expect("Failed to execute excel-cli");

    assert_json_error(&output, 6); // EXIT_INVALID_QUERY
}

#[test]
fn test_read_cell_formula_reports_formula_metadata() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_read_cell_formula.xlsx");
    create_read_contract_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("read")
        .arg("cell")
        .arg(&file_path)
        .arg("--sheet")
        .arg("TypedCells")
        .arg("--cell")
        .arg("E2")
        .output()
        .expect("Failed to execute excel-cli");

    let json = assert_json_success(&output);
    assert_eq!(json["data"]["cell"], "E2");
    assert_eq!(json["data"]["type"], "formula");
    assert_eq!(json["data"]["value"], 85);
    assert_eq!(json["data"]["formula"], "=B2*2");
}

#[test]
fn test_read_range_json() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_read_range.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("read")
        .arg("range")
        .arg(&file_path)
        .arg("--sheet")
        .arg("Orders")
        .arg("--range")
        .arg("A1:B2")
        .output()
        .expect("Failed to execute excel-cli");

    let json = assert_json_success(&output);
    assert_eq!(json["command"], "read.range");
    assert_eq!(json["target"]["sheet"], "Orders");
    assert_eq!(json["data"]["range"], "A1:B2");
    let rows = json["data"]["rows"].as_array().unwrap();
    assert_eq!(rows.len(), 2);
    assert_eq!(rows[0][0], "order_id");
    assert_eq!(rows[1][1], "Alice");
}

#[test]
fn test_read_range_text() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_range_text.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("read")
        .arg("range")
        .arg(&file_path)
        .arg("--sheet")
        .arg("Orders")
        .arg("--range")
        .arg("A1:B2")
        .arg("--format")
        .arg("text")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(output.status.success());
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("order_id\tcustomer"));
    assert!(stdout.contains("1001\tAlice"));
}

#[test]
fn test_read_range_invalid_format() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_range_bad.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("read")
        .arg("range")
        .arg(&file_path)
        .arg("--sheet")
        .arg("Orders")
        .arg("--range")
        .arg("BAD")
        .output()
        .expect("Failed to execute excel-cli");

    assert_json_error(&output, 6); // EXIT_INVALID_QUERY
}

#[test]
fn test_read_range_preserves_typed_values() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_read_range_typed.xlsx");
    create_read_contract_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("read")
        .arg("range")
        .arg(&file_path)
        .arg("--sheet")
        .arg("TypedCells")
        .arg("--range")
        .arg("A2:F2")
        .output()
        .expect("Failed to execute excel-cli");

    let json = assert_json_success(&output);
    let rows = json["data"]["rows"].as_array().unwrap();
    assert_eq!(rows.len(), 1);
    assert_eq!(rows[0][0], "hello");
    assert_eq!(rows[0][1], 42.5);
    assert_eq!(rows[0][2], "2024-02-03");
    assert_eq!(rows[0][3], true);
    assert_eq!(rows[0][4], 85);
    assert!(rows[0][5].is_null());
}

#[test]
fn test_read_rows_json() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_read_rows.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("read")
        .arg("rows")
        .arg(&file_path)
        .arg("--sheet")
        .arg("Orders")
        .output()
        .expect("Failed to execute excel-cli");

    let json = assert_json_success(&output);
    assert_eq!(json["command"], "read.rows");
    // Should return records mode because header row 1 is detected
    assert_eq!(json["data"]["mode"], "records");
    let records = json["data"]["records"].as_array().unwrap();
    assert_eq!(records.len(), 1);
    assert_eq!(records[0]["order_id"], "1001");
    assert_eq!(records[0]["customer"], "Alice");
}

#[test]
fn test_read_rows_with_header_row() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_read_rows_hdr.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("read")
        .arg("rows")
        .arg(&file_path)
        .arg("--sheet")
        .arg("Orders")
        .arg("--header-row")
        .arg("1")
        .output()
        .expect("Failed to execute excel-cli");

    let json = assert_json_success(&output);
    assert_eq!(json["data"]["mode"], "records");
    assert_eq!(json["data"]["resolved_header_row"], 1);
    let records = json["data"]["records"].as_array().unwrap();
    assert_eq!(records.len(), 1);
    assert_eq!(records[0]["order_id"], "1001");
}

#[test]
fn test_read_rows_no_header() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_read_rows_no_hdr.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("read")
        .arg("rows")
        .arg(&file_path)
        .arg("--sheet")
        .arg("Orders")
        .arg("--header-row")
        .arg("999")
        .output()
        .expect("Failed to execute excel-cli");

    let json = assert_json_success(&output);
    assert_eq!(json["data"]["mode"], "rows");
    let rows = json["data"]["rows"].as_array().unwrap();
    assert_eq!(rows.len(), 2);
}

#[test]
fn test_read_rows_auto_skips_preamble_and_keeps_unique_keys() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_read_rows_auto_contract.xlsx");
    create_read_contract_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("read")
        .arg("rows")
        .arg(&file_path)
        .arg("--sheet")
        .arg("RowCases")
        .arg("--range")
        .arg("A1:D4")
        .output()
        .expect("Failed to execute excel-cli");

    let json = assert_json_success(&output);
    assert_eq!(json["data"]["mode"], "records");
    assert_eq!(json["data"]["resolved_header_row"], 2);
    let records = json["data"]["records"].as_array().unwrap();
    assert_eq!(records.len(), 2);
    assert_eq!(records[0]["order_id"], "1001");
    assert_eq!(records[0]["customer"], "Alice");
    assert_eq!(records[0]["customer_2"], "VIP");
    assert_eq!(records[0]["col_D"], true);
    assert_eq!(records[1]["order_id"], "1002");
}

#[test]
fn test_read_rows_explicit_header_row_respects_selected_range() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_read_rows_explicit_contract.xlsx");
    create_read_contract_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("read")
        .arg("rows")
        .arg(&file_path)
        .arg("--sheet")
        .arg("RowCases")
        .arg("--range")
        .arg("A1:D4")
        .arg("--header-row")
        .arg("2")
        .output()
        .expect("Failed to execute excel-cli");

    let json = assert_json_success(&output);
    assert_eq!(json["data"]["mode"], "records");
    assert_eq!(json["data"]["resolved_header_row"], 2);
    let records = json["data"]["records"].as_array().unwrap();
    assert_eq!(records.len(), 2);
    assert!(records[0].get("Quarterly export").is_none());
}

#[test]
fn test_inspect_sample_json() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_inspect_sample.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("inspect")
        .arg("sample")
        .arg(&file_path)
        .arg("--sheet")
        .arg("Orders")
        .arg("--rows")
        .arg("2")
        .output()
        .expect("Failed to execute excel-cli");

    let json = assert_json_success(&output);
    assert_eq!(json["command"], "inspect.sample");
    assert_eq!(json["target"]["sheet"], "Orders");
    assert!(json["data"]["sample_mode"].is_string());
    assert!(json["data"]["rows"].is_array() || json["data"]["records"].is_array());
}

#[test]
fn test_inspect_columns_json_contract() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_inspect_columns_contract.xlsx");
    create_columns_contract_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("inspect")
        .arg("columns")
        .arg(&file_path)
        .arg("--sheet")
        .arg("ColumnCases")
        .output()
        .expect("Failed to execute excel-cli");

    let json = assert_json_success(&output);
    assert_eq!(json["schema_version"], "1.0");
    assert_eq!(json["command"], "inspect.columns");
    assert_eq!(json["file"]["format"], "xlsx");
    assert_eq!(json["target"]["sheet"], "ColumnCases");
    assert_eq!(json["target"]["sheet_index"], 0);
    assert_eq!(json["meta"]["header_row_mode"], "auto");
    assert_eq!(json["meta"]["resolved_header_row"], 1);
    assert_eq!(json["meta"]["column_count"], 7);
    assert_eq!(json["meta"]["data_row_count"], 2);
    assert!(json["warnings"].as_array().unwrap().is_empty());

    let columns = json["data"]["columns"].as_array().unwrap();
    assert_eq!(columns.len(), 7);
    assert_eq!(columns[0]["index"], 1);
    assert_eq!(columns[0]["name"], "customer");
    assert_eq!(columns[0]["safe_name"], "customer");
    assert_eq!(columns[0]["is_duplicate"], true);
    assert_eq!(columns[0]["inferred_type"], "string");
    assert_eq!(columns[0]["non_null_ratio"], 1.0);
    assert_eq!(columns[0]["formula_ratio"], 0.0);
    assert_eq!(
        columns[0]["sample_values"],
        serde_json::json!(["Alice", "Bob"])
    );

    assert_eq!(columns[1]["name"], "customer");
    assert_eq!(columns[1]["safe_name"], "customer_2");
    assert_eq!(columns[1]["is_duplicate"], true);

    assert_eq!(columns[2]["name"], "");
    assert_eq!(columns[2]["safe_name"], "col_C");
    assert_eq!(columns[2]["is_duplicate"], false);
    assert_eq!(columns[2]["non_null_ratio"], 0.5);

    assert_eq!(columns[3]["name"], "数量");
    assert_eq!(columns[3]["safe_name"], "数量");
    assert_eq!(columns[3]["inferred_type"], "number");
    assert_eq!(columns[3]["sample_values"], serde_json::json!([3, 4]));

    assert_eq!(columns[4]["inferred_type"], "mixed");
    assert_eq!(columns[5]["inferred_type"], "boolean");
    assert_eq!(columns[6]["inferred_type"], "number");
    assert_eq!(columns[6]["formula_ratio"], 1.0);
    assert_eq!(columns[6]["sample_values"], serde_json::json!([6, 8]));
}

#[test]
fn test_inspect_columns_explicit_header_row_and_text() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_inspect_columns_text.xlsx");
    create_columns_contract_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("inspect")
        .arg("columns")
        .arg(&file_path)
        .arg("--sheet")
        .arg("OffsetHeader")
        .arg("--header-row")
        .arg("2")
        .arg("--format")
        .arg("text")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(
        output.status.success(),
        "Expected success. stderr: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    assert!(output.stderr.is_empty());
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(!stdout.trim_start().starts_with('{'));
    assert!(stdout.contains("index\tname\tsafe_name"));
    assert!(stdout.contains("1\titem\titem\tfalse\tstring\t1\t0"));
    assert!(stdout.contains("2\tamount\tamount\tfalse\tnumber\t1\t0"));
}

#[test]
fn test_inspect_columns_invalid_header_row() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_inspect_columns_bad_header.xlsx");
    create_columns_contract_workbook(&file_path);

    let non_numeric = Command::new(excel_cli_bin())
        .arg("inspect")
        .arg("columns")
        .arg(&file_path)
        .arg("--sheet")
        .arg("ColumnCases")
        .arg("--header-row")
        .arg("nope")
        .output()
        .expect("Failed to execute excel-cli");
    assert_json_error(&non_numeric, 6);

    let out_of_range = Command::new(excel_cli_bin())
        .arg("inspect")
        .arg("columns")
        .arg(&file_path)
        .arg("--sheet")
        .arg("ColumnCases")
        .arg("--header-row")
        .arg("999")
        .output()
        .expect("Failed to execute excel-cli");
    assert_json_error(&out_of_range, 6);
}

#[test]
fn test_check_namespace_only() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_check.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("check")
        .arg(&file_path)
        .arg("--rule")
        .arg("missing_values")
        .output()
        .expect("Failed to execute excel-cli");

    assert_json_error(&output, 6); // EXIT_INVALID_QUERY
    let err_json: serde_json::Value =
        serde_json::from_slice(&output.stderr).expect("Valid JSON error");
    assert_eq!(err_json["error"]["code"], "check_not_implemented");
}

#[test]
fn test_check_help_without_rule() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_check_no_rule.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("check")
        .arg(&file_path)
        .output()
        .expect("Failed to execute excel-cli");

    assert_json_error(&output, 6); // EXIT_INVALID_QUERY
    let err_json: serde_json::Value =
        serde_json::from_slice(&output.stderr).expect("Valid JSON error");
    assert_eq!(err_json["error"]["code"], "check_not_implemented");
}

#[test]
fn test_bare_file_path_is_error() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_bare.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg(&file_path)
        .output()
        .expect("Failed to execute excel-cli");

    assert_json_error(&output, 2); // EXIT_INVALID_ARGS
}
