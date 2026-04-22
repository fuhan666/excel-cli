use std::path::{Path, PathBuf};
use std::process::{Command, Output};

use serde_json::{json, Value};

fn excel_cli_bin() -> PathBuf {
    PathBuf::from(env!("CARGO_BIN_EXE_excel-cli"))
}

fn create_check_workbook(path: &Path) {
    use rust_xlsxwriter::Workbook as XlsxWorkbook;

    let mut workbook = XlsxWorkbook::new();

    let summary = workbook.add_worksheet();
    summary.set_name("Summary").unwrap();
    summary.write_string(0, 0, "Total").unwrap();
    summary.write_number(0, 1, 1234.5).unwrap();

    let orders = workbook.add_worksheet();
    orders.set_name("Orders").unwrap();
    orders.write_string(0, 0, "order_id").unwrap();
    orders.write_string(0, 1, "customer").unwrap();
    orders.write_string(1, 0, "1001").unwrap();
    orders.write_string(1, 1, "Alice").unwrap();

    let customers = workbook.add_worksheet();
    customers.set_name("客户").unwrap();
    customers.write_string(0, 0, "姓名").unwrap();
    customers.write_string(1, 0, "张三").unwrap();

    let empty = workbook.add_worksheet();
    empty.set_name("EmptySheet").unwrap();

    workbook.save(path).unwrap();
}

fn temp_workbook(name: &str) -> PathBuf {
    let path = std::env::temp_dir().join(name);
    create_check_workbook(&path);
    path
}

fn run_check(args: &[&str]) -> Output {
    Command::new(excel_cli_bin())
        .args(args)
        .output()
        .unwrap_or_else(|_| panic!("Failed to execute excel-cli {}", args.join(" ")))
}

fn parse_stdout(output: &Output) -> Value {
    serde_json::from_slice(&output.stdout).expect("stdout should be valid JSON")
}

fn parse_stderr(output: &Output) -> Value {
    serde_json::from_slice(&output.stderr).expect("stderr should be valid JSON")
}

fn assert_success(output: &Output, code: i32) {
    assert_eq!(output.status.code(), Some(code));
    assert!(
        output.stderr.is_empty(),
        "stderr should be empty: {}",
        String::from_utf8_lossy(&output.stderr)
    );
}

fn assert_json_error(output: &Output, code: i32, error_code: &str) -> Value {
    assert_eq!(output.status.code(), Some(code));
    assert!(
        output.stdout.is_empty(),
        "stdout should be empty: {}",
        String::from_utf8_lossy(&output.stdout)
    );
    let err = parse_stderr(output);
    assert_eq!(err["error"]["code"], error_code);
    err
}

#[test]
fn check_workbook_returns_stable_empty_report_contract() {
    let file_path = temp_workbook("excel_cli_check_contract_workbook.xlsx");
    let output = Command::new(excel_cli_bin())
        .arg("check")
        .arg(&file_path)
        .output()
        .expect("Failed to execute excel-cli check");

    assert_success(&output, 0);
    let json = parse_stdout(&output);

    assert_eq!(json["schema_version"], "1.0");
    assert_eq!(json["command"], "check");
    assert_eq!(json["file"]["path"], file_path.to_string_lossy().as_ref());
    assert_eq!(json["file"]["format"], "xlsx");
    assert_eq!(json["target"], json!({}));
    assert_eq!(json["warnings"], json!([]));

    assert_eq!(
        json["data"]["summary"],
        json!({
            "status": "pass",
            "finding_count": 0,
            "error_count": 0,
            "warning_count": 0,
            "info_count": 0
        })
    );
    assert_eq!(json["data"]["findings"], json!([]));
    assert_eq!(json["data"]["stats"]["sheet_count"], 4);
    assert_eq!(json["data"]["stats"]["checked_sheet_count"], 4);
    assert_eq!(json["data"]["stats"]["severity_threshold"], "info");
    assert_eq!(json["data"]["stats"]["finding_count_before_threshold"], 0);
    assert_eq!(
        json["data"]["stats"]["rules_run"],
        json!([
            "blank_headers",
            "duplicate_headers",
            "blank_rows",
            "blank_columns",
            "null_ratio",
            "duplicate_values",
            "type_drift",
            "formula_presence"
        ])
    );
    assert_eq!(
        json["data"]["stats"]["checked_sheets"],
        json!([
            {
                "name": "Summary",
                "index": 0,
                "used_range": "A1:B1",
                "max_rows": 1,
                "max_cols": 2
            },
            {
                "name": "Orders",
                "index": 1,
                "used_range": "A1:B2",
                "max_rows": 2,
                "max_cols": 2
            },
            {
                "name": "客户",
                "index": 2,
                "used_range": "A1:A2",
                "max_rows": 2,
                "max_cols": 1
            },
            {
                "name": "EmptySheet",
                "index": 3,
                "used_range": "",
                "max_rows": 0,
                "max_cols": 0
            }
        ])
    );
}

#[test]
fn check_sheet_accepts_rules_and_threshold_with_registry_order() {
    let file_path = temp_workbook("excel_cli_check_contract_sheet.xlsx");
    let output = Command::new(excel_cli_bin())
        .arg("check")
        .arg(&file_path)
        .arg("--sheet")
        .arg("客户")
        .arg("--rules")
        .arg("duplicate_headers,blank_headers")
        .arg("--severity-threshold")
        .arg("warning")
        .output()
        .expect("Failed to execute excel-cli check");

    assert_success(&output, 0);
    let json = parse_stdout(&output);

    assert_eq!(json["target"], json!({"sheet": "客户", "sheet_index": 2}));
    assert_eq!(json["data"]["summary"]["status"], "pass");
    assert_eq!(json["data"]["findings"], json!([]));
    assert_eq!(json["data"]["stats"]["sheet_count"], 4);
    assert_eq!(json["data"]["stats"]["checked_sheet_count"], 1);
    assert_eq!(json["data"]["stats"]["severity_threshold"], "warning");
    assert_eq!(
        json["data"]["stats"]["rules_run"],
        json!(["blank_headers", "duplicate_headers"])
    );
    assert_eq!(
        json["data"]["stats"]["checked_sheets"],
        json!([{
            "name": "客户",
            "index": 2,
            "used_range": "A1:A2",
            "max_rows": 2,
            "max_cols": 1
        }])
    );
}

#[test]
fn check_rejects_unknown_and_empty_rule_lists_as_query_errors() {
    let file_path = temp_workbook("excel_cli_check_contract_invalid_rules.xlsx");
    let file_arg = file_path.to_string_lossy();

    let unknown = run_check(&[
        "check",
        file_arg.as_ref(),
        "--rules",
        "blank_headers,unknown_rule",
    ]);
    let err = assert_json_error(&unknown, 6, "invalid_query");
    assert!(
        err["error"]["message"]
            .as_str()
            .unwrap()
            .contains("unknown_rule"),
        "unexpected error: {err}"
    );

    let empty = run_check(&["check", file_arg.as_ref(), "--rules", " , "]);
    let err = assert_json_error(&empty, 6, "invalid_query");
    assert!(
        err["error"]["message"]
            .as_str()
            .unwrap()
            .contains("--rules"),
        "unexpected error: {err}"
    );
}

#[test]
fn check_rejects_missing_sheet_as_target_not_found() {
    let file_path = temp_workbook("excel_cli_check_contract_missing_sheet.xlsx");
    let output = Command::new(excel_cli_bin())
        .arg("check")
        .arg(&file_path)
        .arg("--sheet")
        .arg("Missing")
        .output()
        .expect("Failed to execute excel-cli check");

    let err = assert_json_error(&output, 5, "target_not_found");
    assert!(
        err["error"]["message"]
            .as_str()
            .unwrap()
            .contains("Missing"),
        "unexpected error: {err}"
    );
}

#[test]
fn check_rejects_legacy_rule_flag_at_parser_level() {
    let file_path = temp_workbook("excel_cli_check_contract_legacy_rule.xlsx");
    let output = Command::new(excel_cli_bin())
        .arg("check")
        .arg(&file_path)
        .arg("--rule")
        .arg("missing_values")
        .output()
        .expect("Failed to execute excel-cli check");

    let err = assert_json_error(&output, 2, "invalid_args");
    assert_ne!(err["error"]["code"], "check_not_implemented");
    assert!(
        err["error"]["message"].as_str().unwrap().contains("--rule"),
        "unexpected error: {err}"
    );
}
