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

fn create_read_rows_extensions_workbook(path: &std::path::Path) {
    use rust_xlsxwriter::Workbook as XlsxWorkbook;

    let mut workbook = XlsxWorkbook::new();

    let sheet = workbook.add_worksheet();
    sheet.set_name("FilterRows").unwrap();
    sheet.write_string(0, 0, "name").unwrap();
    sheet.write_string(0, 1, "amount").unwrap();
    sheet.write_string(0, 2, "status").unwrap();
    sheet.write_string(0, 3, "note").unwrap();
    sheet.write_string(0, 4, "optional").unwrap();

    sheet.write_string(1, 0, "Alice").unwrap();
    sheet.write_number(1, 1, 10.0).unwrap();
    sheet.write_string(1, 2, "open").unwrap();
    sheet.write_string(1, 3, "alpha").unwrap();

    sheet.write_string(2, 0, "Bob").unwrap();
    sheet.write_number(2, 1, 25.0).unwrap();
    sheet.write_string(2, 2, "closed").unwrap();
    sheet.write_string(2, 3, "beta").unwrap();
    sheet.write_string(2, 4, "x").unwrap();

    sheet.write_string(3, 0, "Carol").unwrap();
    sheet.write_number(3, 1, 40.0).unwrap();
    sheet.write_string(3, 2, "open").unwrap();
    sheet.write_string(3, 3, "carol:tag").unwrap();

    sheet.write_string(5, 0, "Delta").unwrap();
    sheet.write_number(5, 1, 55.0).unwrap();
    sheet.write_string(5, 2, "open").unwrap();
    sheet.write_string(5, 3, "delta-special").unwrap();

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

fn create_table_detection_workbook(path: &std::path::Path) {
    use rust_xlsxwriter::Workbook as XlsxWorkbook;

    let mut workbook = XlsxWorkbook::new();

    let single = workbook.add_worksheet();
    single.set_name("SingleTable").unwrap();
    single.write_string(0, 0, "order_id").unwrap();
    single.write_string(0, 1, "customer").unwrap();
    single.write_string(0, 2, "amount").unwrap();
    single.write_string(1, 0, "1001").unwrap();
    single.write_string(1, 1, "Alice").unwrap();
    single.write_number(1, 2, 125.0).unwrap();
    single.write_string(2, 0, "1002").unwrap();
    single.write_string(2, 1, "Bob").unwrap();
    single.write_number(2, 2, 250.0).unwrap();
    single.write_string(3, 0, "1003").unwrap();
    single.write_string(3, 1, "Carol").unwrap();
    single.write_number(3, 2, 375.0).unwrap();

    let multiple = workbook.add_worksheet();
    multiple.set_name("MultipleTables").unwrap();
    multiple.write_string(0, 0, "region").unwrap();
    multiple.write_string(0, 1, "sales").unwrap();
    multiple.write_string(1, 0, "East").unwrap();
    multiple.write_number(1, 1, 10.0).unwrap();
    multiple.write_string(2, 0, "West").unwrap();
    multiple.write_number(2, 1, 12.0).unwrap();
    multiple.write_string(0, 4, "sku").unwrap();
    multiple.write_string(0, 5, "qty").unwrap();
    multiple.write_string(0, 6, "status").unwrap();
    multiple.write_string(1, 4, "A-1").unwrap();
    multiple.write_number(1, 5, 5.0).unwrap();
    multiple.write_string(1, 6, "open").unwrap();
    multiple.write_string(2, 4, "B-2").unwrap();
    multiple.write_number(2, 5, 7.0).unwrap();
    multiple.write_string(2, 6, "closed").unwrap();
    multiple.write_string(3, 4, "C-3").unwrap();
    multiple.write_number(3, 5, 9.0).unwrap();
    multiple.write_string(3, 6, "open").unwrap();

    let preamble = workbook.add_worksheet();
    preamble.set_name("Preamble").unwrap();
    preamble.write_string(0, 0, "Quarterly export").unwrap();
    preamble
        .write_string(1, 0, "Generated for finance")
        .unwrap();
    preamble.write_string(3, 0, "Prepared by ops").unwrap();
    preamble.write_string(4, 0, "invoice_id").unwrap();
    preamble.write_string(4, 1, "customer").unwrap();
    preamble.write_string(4, 2, "total").unwrap();
    preamble.write_string(5, 0, "I-001").unwrap();
    preamble.write_string(5, 1, "Delta").unwrap();
    preamble.write_number(5, 2, 99.0).unwrap();
    preamble.write_string(6, 0, "I-002").unwrap();
    preamble.write_string(6, 1, "Echo").unwrap();
    preamble.write_number(6, 2, 149.0).unwrap();

    let empty = workbook.add_worksheet();
    empty.set_name("EmptyTables").unwrap();

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

fn assert_jsonl_success(output: &std::process::Output) -> Vec<serde_json::Value> {
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

    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(
        !stdout.trim_start().starts_with('['),
        "JSONL output must not be an array: {stdout}"
    );
    assert!(
        !stdout.trim_start().starts_with("{\n  \"schema_version\""),
        "JSONL output must not be an envelope: {stdout}"
    );

    stdout
        .lines()
        .map(|line| serde_json::from_str(line).expect("Expected valid JSONL record"))
        .collect()
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
        .arg("--output-shape")
        .arg("records")
        .output()
        .expect("Failed to execute excel-cli");

    let json = assert_json_success(&output);
    assert_eq!(json["command"], "read.rows");
    // Explicit records shape should use detected header row 1.
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
        .arg("--output-shape")
        .arg("records")
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
        .arg("--output-shape")
        .arg("records")
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
        .arg("--output-shape")
        .arg("records")
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
fn test_read_rows_select_filters_pagination_and_meta() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_read_rows_extensions.xlsx");
    create_read_rows_extensions_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("read")
        .arg("rows")
        .arg(&file_path)
        .arg("--sheet")
        .arg("FilterRows")
        .arg("--range")
        .arg("A1:E6")
        .arg("--select")
        .arg("name,amount")
        .arg("--filter")
        .arg("status:eq:open")
        .arg("--filter")
        .arg("amount:gte:40")
        .arg("--limit")
        .arg("1")
        .arg("--offset")
        .arg("0")
        .arg("--output-shape")
        .arg("records")
        .output()
        .expect("Failed to execute excel-cli");

    let json = assert_json_success(&output);
    assert_eq!(json["data"]["mode"], "records");
    assert_eq!(
        json["meta"]["applied_filters"],
        serde_json::json!(["status:eq:open", "amount:gte:40"])
    );
    assert_eq!(
        json["meta"]["selected_columns"],
        serde_json::json!(["name", "amount"])
    );
    assert_eq!(json["meta"]["row_count"], 1);
    assert_eq!(json["meta"]["truncated"], true);

    let records = json["data"]["records"].as_array().unwrap();
    assert_eq!(records.len(), 1);
    assert_eq!(
        records[0],
        serde_json::json!({"name": "Carol", "amount": 40})
    );
}

#[test]
fn test_read_rows_output_shape_rows_returns_positional_rows() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_read_rows_output_shape_rows.xlsx");
    create_read_rows_extensions_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("read")
        .arg("rows")
        .arg(&file_path)
        .arg("--sheet")
        .arg("FilterRows")
        .arg("--range")
        .arg("A1:C4")
        .arg("--select")
        .arg("name,status")
        .arg("--filter")
        .arg("status:eq:open")
        .arg("--output-shape")
        .arg("rows")
        .output()
        .expect("Failed to execute excel-cli");

    let json = assert_json_success(&output);
    assert_eq!(json["command"], "read.rows");
    assert_eq!(json["meta"]["output_shape"], "rows");
    assert_eq!(json["data"]["mode"], "rows");
    assert!(json["data"].get("records").is_none());
    assert_eq!(
        json["data"]["rows"],
        serde_json::json!([["Alice", "open"], ["Carol", "open"]])
    );
}

#[test]
fn test_read_records_defaults_to_records_shape() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_read_records_default.xlsx");
    create_read_rows_extensions_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("read")
        .arg("records")
        .arg(&file_path)
        .arg("--sheet")
        .arg("FilterRows")
        .arg("--range")
        .arg("A1:C4")
        .arg("--select")
        .arg("name,amount")
        .arg("--filter")
        .arg("status:eq:open")
        .output()
        .expect("Failed to execute excel-cli");

    let json = assert_json_success(&output);
    assert_eq!(json["command"], "read.records");
    assert_eq!(json["meta"]["output_shape"], "records");
    assert_eq!(json["data"]["mode"], "records");
    assert!(json["data"].get("rows").is_none());
    assert_eq!(
        json["data"]["records"],
        serde_json::json!([
            {"name": "Alice", "amount": 10},
            {"name": "Carol", "amount": 40}
        ])
    );
}

#[test]
fn test_output_shape_jsonl_writes_newline_delimited_records() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_read_records_jsonl.xlsx");
    create_read_rows_extensions_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("read")
        .arg("records")
        .arg(&file_path)
        .arg("--sheet")
        .arg("FilterRows")
        .arg("--range")
        .arg("A1:E6")
        .arg("--select")
        .arg("name,amount")
        .arg("--filter")
        .arg("status:eq:open")
        .arg("--limit")
        .arg("2")
        .arg("--output-shape")
        .arg("jsonl")
        .output()
        .expect("Failed to execute excel-cli");

    let records = assert_jsonl_success(&output);
    assert_eq!(
        records,
        vec![
            serde_json::json!({"name": "Alice", "amount": 10}),
            serde_json::json!({"name": "Carol", "amount": 40}),
        ]
    );
}

#[test]
fn test_output_shape_does_not_change_selection_filter_or_pagination() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_read_output_shape_invariance.xlsx");
    create_read_rows_extensions_workbook(&file_path);

    let mut base_args = vec![
        "read",
        "rows",
        file_path.to_str().unwrap(),
        "--sheet",
        "FilterRows",
        "--range",
        "A1:E6",
        "--select",
        "name,amount",
        "--filter",
        "status:eq:open",
        "--offset",
        "1",
        "--limit",
        "2",
    ];

    let rows_output = Command::new(excel_cli_bin())
        .args(&base_args)
        .arg("--output-shape")
        .arg("rows")
        .output()
        .expect("Failed to execute excel-cli");
    let rows_json = assert_json_success(&rows_output);

    base_args[1] = "records";
    let records_output = Command::new(excel_cli_bin())
        .args(&base_args)
        .arg("--output-shape")
        .arg("records")
        .output()
        .expect("Failed to execute excel-cli");
    let records_json = assert_json_success(&records_output);

    let jsonl_output = Command::new(excel_cli_bin())
        .args(&base_args)
        .arg("--output-shape")
        .arg("jsonl")
        .output()
        .expect("Failed to execute excel-cli");
    let jsonl_records = assert_jsonl_success(&jsonl_output);

    assert_eq!(rows_json["meta"]["row_count"], 2);
    assert_eq!(rows_json["meta"]["truncated"], false);
    assert_eq!(
        rows_json["meta"]["selected_columns"],
        serde_json::json!(["name", "amount"])
    );
    assert_eq!(
        records_json["meta"]["row_count"],
        rows_json["meta"]["row_count"]
    );
    assert_eq!(
        records_json["meta"]["selected_columns"],
        rows_json["meta"]["selected_columns"]
    );
    assert_eq!(
        rows_json["data"]["rows"],
        serde_json::json!([["Carol", 40], ["Delta", 55]])
    );
    assert_eq!(
        records_json["data"]["records"],
        serde_json::json!([
            {"name": "Carol", "amount": 40},
            {"name": "Delta", "amount": 55}
        ])
    );
    assert_eq!(
        serde_json::Value::Array(jsonl_records),
        records_json["data"]["records"]
    );
}

#[test]
fn test_read_records_requires_resolved_header() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_read_records_header_error.xlsx");
    create_read_rows_extensions_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("read")
        .arg("records")
        .arg(&file_path)
        .arg("--sheet")
        .arg("FilterRows")
        .arg("--range")
        .arg("A2:E6")
        .arg("--header-row")
        .arg("999")
        .output()
        .expect("Failed to execute excel-cli");

    assert_json_error(&output, 6);

    let output = Command::new(excel_cli_bin())
        .arg("read")
        .arg("rows")
        .arg(&file_path)
        .arg("--sheet")
        .arg("FilterRows")
        .arg("--range")
        .arg("A2:E6")
        .arg("--header-row")
        .arg("999")
        .arg("--output-shape")
        .arg("jsonl")
        .output()
        .expect("Failed to execute excel-cli");

    assert_json_error(&output, 6);
}

#[test]
fn test_output_shape_jsonl_rejects_text_format() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_read_records_jsonl_text.xlsx");
    create_read_rows_extensions_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("read")
        .arg("records")
        .arg(&file_path)
        .arg("--sheet")
        .arg("FilterRows")
        .arg("--output-shape")
        .arg("jsonl")
        .arg("--format")
        .arg("text")
        .output()
        .expect("Failed to execute excel-cli");

    assert_json_error(&output, 2);
}

#[test]
fn test_read_rows_filter_operators_and_non_empty() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_read_rows_filter_ops.xlsx");
    create_read_rows_extensions_workbook(&file_path);

    let cases = [
        ("name:ne:Alice", vec!["Bob", "Carol", "Delta"]),
        ("amount:gt:25", vec!["Carol", "Delta"]),
        ("amount:lt:40", vec!["Alice", "Bob"]),
        ("amount:lte:25", vec!["Alice", "Bob"]),
        ("note:contains:special", vec!["Delta"]),
        ("name:regex:^(Alice|Carol)$", vec!["Alice", "Carol"]),
        ("optional:isnull:", vec!["Alice", "Carol", "Delta"]),
        ("optional:notnull:", vec!["Bob"]),
    ];

    for (filter, expected_names) in cases {
        let output = Command::new(excel_cli_bin())
            .arg("read")
            .arg("rows")
            .arg(&file_path)
            .arg("--sheet")
            .arg("FilterRows")
            .arg("--range")
            .arg("A1:E6")
            .arg("--select")
            .arg("name")
            .arg("--filter")
            .arg(filter)
            .arg("--non-empty")
            .arg("--output-shape")
            .arg("records")
            .output()
            .expect("Failed to execute excel-cli");

        let json = assert_json_success(&output);
        let actual_names: Vec<&str> = json["data"]["records"]
            .as_array()
            .unwrap()
            .iter()
            .map(|record| record["name"].as_str().unwrap())
            .collect();
        assert_eq!(actual_names, expected_names, "filter {filter}");
        assert_eq!(json["meta"]["row_count"], expected_names.len());
        assert_eq!(json["meta"]["truncated"], false);
    }
}

#[test]
fn test_read_rows_no_match_and_raw_column_selection() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_read_rows_raw_select.xlsx");
    create_read_rows_extensions_workbook(&file_path);

    let no_match = Command::new(excel_cli_bin())
        .arg("read")
        .arg("rows")
        .arg(&file_path)
        .arg("--sheet")
        .arg("FilterRows")
        .arg("--range")
        .arg("A1:E6")
        .arg("--filter")
        .arg("name:eq:Missing")
        .arg("--output-shape")
        .arg("records")
        .output()
        .expect("Failed to execute excel-cli");

    let json = assert_json_success(&no_match);
    assert!(json["data"]["records"].as_array().unwrap().is_empty());
    assert_eq!(json["meta"]["row_count"], 0);
    assert_eq!(json["meta"]["truncated"], false);

    let raw = Command::new(excel_cli_bin())
        .arg("read")
        .arg("rows")
        .arg(&file_path)
        .arg("--sheet")
        .arg("FilterRows")
        .arg("--range")
        .arg("A2:E6")
        .arg("--header-row")
        .arg("999")
        .arg("--select")
        .arg("col_A,col_C")
        .arg("--filter")
        .arg("col_B:gte:40")
        .arg("--non-empty")
        .output()
        .expect("Failed to execute excel-cli");

    let json = assert_json_success(&raw);
    assert_eq!(json["data"]["mode"], "rows");
    assert_eq!(
        json["meta"]["selected_columns"],
        serde_json::json!(["col_A", "col_C"])
    );
    assert_eq!(json["meta"]["row_count"], 2);
    assert_eq!(
        json["data"]["rows"],
        serde_json::json!([["Carol", "open"], ["Delta", "open"]])
    );
}

#[test]
fn test_read_rows_invalid_select_and_filters_are_structured_errors() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_read_rows_invalid_filters.xlsx");
    create_read_rows_extensions_workbook(&file_path);

    let cases = [
        vec!["--select", "missing"],
        vec!["--filter", "missing:eq:value"],
        vec!["--filter", "name:starts:value"],
        vec!["--filter", "name:eq"],
        vec!["--filter", "name:regex:["],
        vec!["--filter", "amount:gt:not-a-number"],
    ];

    for args in cases {
        let output = Command::new(excel_cli_bin())
            .arg("read")
            .arg("rows")
            .arg(&file_path)
            .arg("--sheet")
            .arg("FilterRows")
            .arg("--range")
            .arg("A1:E6")
            .args(args)
            .output()
            .expect("Failed to execute excel-cli");

        assert_json_error(&output, 6);
    }
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
fn test_inspect_tables_single_table() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_inspect_tables_single.xlsx");
    create_table_detection_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("inspect")
        .arg("tables")
        .arg(&file_path)
        .arg("--sheet")
        .arg("SingleTable")
        .output()
        .expect("Failed to execute excel-cli");

    let json = assert_json_success(&output);
    assert_eq!(json["command"], "inspect.tables");
    assert_eq!(json["target"]["sheet"], "SingleTable");
    assert_eq!(json["data"]["candidates"].as_array().unwrap().len(), 1);
    assert_eq!(json["meta"]["candidate_count"], 1);

    let candidate = &json["data"]["candidates"][0];
    assert_eq!(candidate["range"], "A1:C4");
    assert_eq!(candidate["header_row"], 1);
    assert_eq!(candidate["column_count"], 3);
    assert_eq!(candidate["row_count"], 4);
    let confidence = candidate["confidence"].as_f64().unwrap();
    assert!((0.7..=1.0).contains(&confidence));
}

#[test]
fn test_inspect_tables_multiple_tables() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_inspect_tables_multiple.xlsx");
    create_table_detection_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("inspect")
        .arg("tables")
        .arg(&file_path)
        .arg("--sheet")
        .arg("MultipleTables")
        .output()
        .expect("Failed to execute excel-cli");

    let json = assert_json_success(&output);
    let candidates = json["data"]["candidates"].as_array().unwrap();
    assert_eq!(candidates.len(), 2);
    assert!(json["data"].get("selected").is_none());
    assert!(json["data"].get("recommended").is_none());

    assert_eq!(candidates[0]["range"], "A1:B3");
    assert_eq!(candidates[0]["header_row"], 1);
    assert_eq!(candidates[0]["column_count"], 2);
    assert_eq!(candidates[0]["row_count"], 3);
    assert!((0.7..=1.0).contains(&candidates[0]["confidence"].as_f64().unwrap()));

    assert_eq!(candidates[1]["range"], "E1:G4");
    assert_eq!(candidates[1]["header_row"], 1);
    assert_eq!(candidates[1]["column_count"], 3);
    assert_eq!(candidates[1]["row_count"], 4);
    assert!((0.7..=1.0).contains(&candidates[1]["confidence"].as_f64().unwrap()));
}

#[test]
fn test_inspect_tables_preamble_late_header() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_inspect_tables_preamble.xlsx");
    create_table_detection_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("inspect")
        .arg("tables")
        .arg(&file_path)
        .arg("--sheet")
        .arg("Preamble")
        .output()
        .expect("Failed to execute excel-cli");

    let json = assert_json_success(&output);
    let candidates = json["data"]["candidates"].as_array().unwrap();
    assert_eq!(candidates.len(), 1);

    let candidate = &candidates[0];
    assert_eq!(candidate["range"], "A5:C7");
    assert_eq!(candidate["header_row"], 5);
    assert_eq!(candidate["column_count"], 3);
    assert_eq!(candidate["row_count"], 3);
    assert!((0.7..=1.0).contains(&candidate["confidence"].as_f64().unwrap()));
}

#[test]
fn test_inspect_tables_empty_sheet() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_inspect_tables_empty.xlsx");
    create_table_detection_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("inspect")
        .arg("tables")
        .arg(&file_path)
        .arg("--sheet")
        .arg("EmptyTables")
        .output()
        .expect("Failed to execute excel-cli");

    let json = assert_json_success(&output);
    assert_eq!(json["command"], "inspect.tables");
    assert_eq!(json["target"]["sheet"], "EmptyTables");
    assert_eq!(json["meta"]["candidate_count"], 0);
    assert!(json["data"]["candidates"].as_array().unwrap().is_empty());
}

#[test]
fn test_inspect_tables_text() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_inspect_tables_text.xlsx");
    create_table_detection_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("inspect")
        .arg("tables")
        .arg(&file_path)
        .arg("--sheet")
        .arg("SingleTable")
        .arg("--format")
        .arg("text")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(output.status.success());
    assert!(
        output.stderr.is_empty(),
        "Expected empty stderr on success. stderr: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("A1:C4\theader_row=1"));
    assert!(stdout.contains("columns=3"));
    assert!(stdout.contains("rows=4"));
    assert!(stdout.contains("confidence="));
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
