use std::path::{Path, PathBuf};
use std::process::{Command, Output};

use serde_json::{json, Value};

fn excel_cli_bin() -> PathBuf {
    PathBuf::from(env!("CARGO_BIN_EXE_excel-cli"))
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

fn assert_success(output: &Output, code: i32) {
    assert_eq!(output.status.code(), Some(code));
    assert!(
        output.stderr.is_empty(),
        "stderr should be empty: {}",
        String::from_utf8_lossy(&output.stderr)
    );
}

fn temp_path(name: &str) -> PathBuf {
    std::env::temp_dir().join(name)
}

fn create_null_ratio_positive(path: &Path) {
    use rust_xlsxwriter::Workbook as XlsxWorkbook;

    let mut workbook = XlsxWorkbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Orders").unwrap();
    sheet.write_string(0, 0, "order_id").unwrap();
    sheet.write_string(0, 1, "customer").unwrap();
    sheet.write_string(0, 2, "email").unwrap();
    sheet.write_string(1, 0, "1001").unwrap();
    sheet.write_string(1, 1, "Alice").unwrap();
    sheet.write_string(1, 2, "alice@example.test").unwrap();
    sheet.write_string(2, 0, "1002").unwrap();
    sheet.write_string(2, 1, "Bob").unwrap();
    sheet.write_string(3, 0, "1003").unwrap();
    sheet.write_string(3, 1, "Cara").unwrap();
    workbook.save(path).unwrap();
}

fn create_clean_values(path: &Path) {
    use rust_xlsxwriter::Workbook as XlsxWorkbook;

    let mut workbook = XlsxWorkbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Orders").unwrap();
    sheet.write_string(0, 0, "order_id").unwrap();
    sheet.write_string(0, 1, "customer").unwrap();
    sheet.write_string(0, 2, "amount").unwrap();
    sheet.write_string(1, 0, "1001").unwrap();
    sheet.write_string(1, 1, "Alice").unwrap();
    sheet.write_number(1, 2, 12.5).unwrap();
    sheet.write_string(2, 0, "1002").unwrap();
    sheet.write_string(2, 1, "Bob").unwrap();
    sheet.write_number(2, 2, 18.0).unwrap();
    sheet.write_string(3, 0, "1003").unwrap();
    sheet.write_string(3, 1, "Cara").unwrap();
    sheet.write_number(3, 2, 21.0).unwrap();
    workbook.save(path).unwrap();
}

fn create_duplicate_values_positive(path: &Path) {
    use rust_xlsxwriter::Workbook as XlsxWorkbook;

    let mut workbook = XlsxWorkbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Orders").unwrap();
    sheet.write_string(0, 0, "order_id").unwrap();
    sheet.write_string(0, 1, "customer").unwrap();
    sheet.write_string(1, 0, "1001").unwrap();
    sheet.write_string(1, 1, "Alice").unwrap();
    sheet.write_string(2, 0, "1002").unwrap();
    sheet.write_string(2, 1, "Bob").unwrap();
    sheet.write_string(3, 0, "1001").unwrap();
    sheet.write_string(3, 1, "Cara").unwrap();
    workbook.save(path).unwrap();
}

fn create_type_drift_positive(path: &Path) {
    use rust_xlsxwriter::Workbook as XlsxWorkbook;

    let mut workbook = XlsxWorkbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Orders").unwrap();
    sheet.write_string(0, 0, "order_id").unwrap();
    sheet.write_string(0, 1, "amount").unwrap();
    sheet.write_string(1, 0, "1001").unwrap();
    sheet.write_number(1, 1, 12.5).unwrap();
    sheet.write_string(2, 0, "1002").unwrap();
    sheet.write_number(2, 1, 18.0).unwrap();
    sheet.write_string(3, 0, "1003").unwrap();
    sheet.write_string(3, 1, "unknown").unwrap();
    workbook.save(path).unwrap();
}

fn create_formula_presence_positive(path: &Path) {
    use rust_xlsxwriter::Workbook as XlsxWorkbook;

    let mut workbook = XlsxWorkbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Orders").unwrap();
    sheet.write_string(0, 0, "order_id").unwrap();
    sheet.write_string(0, 1, "amount").unwrap();
    sheet.write_string(0, 2, "total").unwrap();
    sheet.write_string(1, 0, "1001").unwrap();
    sheet.write_number(1, 1, 12.5).unwrap();
    sheet.write_formula(1, 2, "=B2*2").unwrap();
    sheet.write_string(2, 0, "1002").unwrap();
    sheet.write_number(2, 1, 18.0).unwrap();
    sheet.write_formula(2, 2, "=B3*2").unwrap();
    workbook.save(path).unwrap();
}

fn create_combined_positive(path: &Path) {
    use rust_xlsxwriter::Workbook as XlsxWorkbook;

    let mut workbook = XlsxWorkbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Orders").unwrap();
    sheet.write_string(0, 0, "order_id").unwrap();
    sheet.write_string(0, 1, "amount").unwrap();
    sheet.write_string(0, 2, "reviewer").unwrap();
    sheet.write_string(0, 3, "total").unwrap();
    sheet.write_string(1, 0, "1001").unwrap();
    sheet.write_number(1, 1, 12.5).unwrap();
    sheet.write_string(1, 2, "Alice").unwrap();
    sheet.write_formula(1, 3, "=B2*2").unwrap();
    sheet.write_string(2, 0, "1002").unwrap();
    sheet.write_number(2, 1, 18.0).unwrap();
    sheet.write_formula(2, 3, "=B3*2").unwrap();
    sheet.write_string(3, 0, "1001").unwrap();
    sheet.write_string(3, 1, "unknown").unwrap();
    sheet.write_formula(3, 3, "=B4*2").unwrap();
    workbook.save(path).unwrap();
}

#[test]
fn check_analytical_null_ratio_reports_blank_data_cells() {
    let file_path = temp_path("excel_cli_check_analytical_null_ratio.xlsx");
    create_null_ratio_positive(&file_path);
    let file_arg = file_path.to_string_lossy();

    let output = run_check(&["check", file_arg.as_ref(), "--rules", "null_ratio"]);

    assert_success(&output, 1);
    let json = parse_stdout(&output);
    assert_eq!(json["data"]["summary"]["warning_count"], 1);
    assert_eq!(json["data"]["stats"]["finding_count_before_threshold"], 1);
    assert_eq!(
        json["data"]["findings"][0],
        json!({
            "rule_id": "null_ratio",
            "severity": "warning",
            "sheet": "Orders",
            "row": 3,
            "column": 3,
            "range": "C2:C4",
            "message": "Column 'email' has blank values in 2 of 3 data rows.",
            "details": {
                "column_name": "email",
                "data_row_count": 3,
                "first_null_cell": "C3",
                "null_count": 2,
                "null_ratio": 0.6667,
                "severity_threshold": {
                    "info": "> 0 and < 0.5",
                    "warning": ">= 0.5 and < 1.0",
                    "error": "1.0"
                }
            }
        })
    );
}

#[test]
fn check_analytical_null_ratio_ignores_complete_columns() {
    let file_path = temp_path("excel_cli_check_analytical_null_ratio_clean.xlsx");
    create_clean_values(&file_path);
    let file_arg = file_path.to_string_lossy();

    let output = run_check(&["check", file_arg.as_ref(), "--rules", "null_ratio"]);

    assert_success(&output, 0);
    let json = parse_stdout(&output);
    assert_eq!(json["data"]["summary"]["status"], "pass");
    assert_eq!(json["data"]["findings"], json!([]));
}

#[test]
fn check_analytical_duplicate_values_checks_default_candidate_column() {
    let file_path = temp_path("excel_cli_check_analytical_duplicate_values.xlsx");
    create_duplicate_values_positive(&file_path);
    let file_arg = file_path.to_string_lossy();

    let output = run_check(&["check", file_arg.as_ref(), "--rules", "duplicate_values"]);

    assert_success(&output, 1);
    let json = parse_stdout(&output);
    assert_eq!(
        json["data"]["findings"][0],
        json!({
            "rule_id": "duplicate_values",
            "severity": "warning",
            "sheet": "Orders",
            "row": 2,
            "column": 1,
            "range": "A2:A4",
            "message": "Column 'order_id' has duplicate value '1001' in 2 rows.",
            "details": {
                "candidate_column": {
                    "column": 1,
                    "column_name": "order_id",
                    "selection": "first non-empty header data column"
                },
                "duplicate_value": "1001",
                "occurrence_count": 2,
                "rows": [2, 4],
                "cells": ["A2", "A4"]
            }
        })
    );
}

#[test]
fn check_analytical_duplicate_values_ignores_unique_candidate_values() {
    let file_path = temp_path("excel_cli_check_analytical_duplicate_values_clean.xlsx");
    create_clean_values(&file_path);
    let file_arg = file_path.to_string_lossy();

    let output = run_check(&["check", file_arg.as_ref(), "--rules", "duplicate_values"]);

    assert_success(&output, 0);
    let json = parse_stdout(&output);
    assert_eq!(json["data"]["findings"], json!([]));
}

#[test]
fn check_analytical_type_drift_reports_mixed_column_types() {
    let file_path = temp_path("excel_cli_check_analytical_type_drift.xlsx");
    create_type_drift_positive(&file_path);
    let file_arg = file_path.to_string_lossy();

    let output = run_check(&["check", file_arg.as_ref(), "--rules", "type_drift"]);

    assert_success(&output, 1);
    let json = parse_stdout(&output);
    assert_eq!(
        json["data"]["findings"][0],
        json!({
            "rule_id": "type_drift",
            "severity": "warning",
            "sheet": "Orders",
            "row": 4,
            "column": 2,
            "range": "B2:B4",
            "message": "Column 'amount' mixes string values with dominant number values.",
            "details": {
                "column_name": "amount",
                "dominant_type": "number",
                "drift_type": "string",
                "drift_count": 1,
                "type_counts": {
                    "number": 2,
                    "string": 1
                },
                "sample_drift_cells": ["B4"]
            }
        })
    );
}

#[test]
fn check_analytical_type_drift_ignores_consistent_columns() {
    let file_path = temp_path("excel_cli_check_analytical_type_drift_clean.xlsx");
    create_clean_values(&file_path);
    let file_arg = file_path.to_string_lossy();

    let output = run_check(&["check", file_arg.as_ref(), "--rules", "type_drift"]);

    assert_success(&output, 0);
    let json = parse_stdout(&output);
    assert_eq!(json["data"]["findings"], json!([]));
}

#[test]
fn check_analytical_formula_presence_reports_formula_cells() {
    let file_path = temp_path("excel_cli_check_analytical_formula_presence.xlsx");
    create_formula_presence_positive(&file_path);
    let file_arg = file_path.to_string_lossy();

    let output = run_check(&["check", file_arg.as_ref(), "--rules", "formula_presence"]);

    assert_success(&output, 1);
    let json = parse_stdout(&output);
    assert_eq!(
        json["data"]["findings"][0],
        json!({
            "rule_id": "formula_presence",
            "severity": "info",
            "sheet": "Orders",
            "row": 2,
            "column": 3,
            "range": "C2:C3",
            "message": "Sheet 'Orders' contains 2 formula cells.",
            "details": {
                "data_row_count": 2,
                "formula_count": 2,
                "formula_ratio": 1.0,
                "sample_formula_cells": [
                    {"cell": "C2", "formula": "=B2*2"},
                    {"cell": "C3", "formula": "=B3*2"}
                ]
            }
        })
    );
}

#[test]
fn check_analytical_formula_presence_ignores_non_formula_sheets() {
    let file_path = temp_path("excel_cli_check_analytical_formula_presence_clean.xlsx");
    create_clean_values(&file_path);
    let file_arg = file_path.to_string_lossy();

    let output = run_check(&["check", file_arg.as_ref(), "--rules", "formula_presence"]);

    assert_success(&output, 0);
    let json = parse_stdout(&output);
    assert_eq!(json["data"]["findings"], json!([]));
}

#[test]
fn check_analytical_rule_combinations_and_thresholds_are_stable() {
    let file_path = temp_path("excel_cli_check_analytical_combined.xlsx");
    create_combined_positive(&file_path);
    let file_arg = file_path.to_string_lossy();

    let output = run_check(&[
        "check",
        file_arg.as_ref(),
        "--rules",
        "formula_presence,null_ratio,duplicate_values,type_drift",
        "--severity-threshold",
        "warning",
    ]);

    assert_success(&output, 1);
    let json = parse_stdout(&output);
    assert_eq!(
        json["data"]["stats"]["rules_run"],
        json!([
            "null_ratio",
            "duplicate_values",
            "type_drift",
            "formula_presence"
        ])
    );
    assert_eq!(json["data"]["stats"]["severity_threshold"], "warning");
    assert_eq!(json["data"]["stats"]["finding_count_before_threshold"], 4);
    assert_eq!(
        json["data"]["summary"],
        json!({
            "status": "fail",
            "finding_count": 3,
            "error_count": 0,
            "warning_count": 3,
            "info_count": 0
        })
    );
    assert_eq!(
        json["data"]["findings"]
            .as_array()
            .unwrap()
            .iter()
            .map(|finding| finding["rule_id"].as_str().unwrap())
            .collect::<Vec<_>>(),
        vec!["null_ratio", "duplicate_values", "type_drift"]
    );
}
