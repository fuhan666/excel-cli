use std::path::{Path, PathBuf};
use std::process::{Command, Output};

use serde_json::{json, Value};

fn excel_cli_bin() -> PathBuf {
    PathBuf::from(env!("CARGO_BIN_EXE_excel-cli"))
}

fn create_structural_workbook(path: &Path) {
    use rust_xlsxwriter::Workbook as XlsxWorkbook;

    let mut workbook = XlsxWorkbook::new();

    let clean = workbook.add_worksheet();
    clean.set_name("Clean").unwrap();
    clean.write_string(0, 0, "客户").unwrap();
    clean.write_string(0, 1, "订单").unwrap();
    clean.write_string(0, 2, "金额").unwrap();
    clean.write_string(1, 0, "张三").unwrap();
    clean.write_string(1, 1, "A-100").unwrap();
    clean.write_number(1, 2, 10).unwrap();
    clean.write_string(2, 0, "李四").unwrap();
    clean.write_string(2, 1, "A-101").unwrap();
    clean.write_number(2, 2, 20).unwrap();

    let structural = workbook.add_worksheet();
    structural.set_name("结构").unwrap();
    structural.write_string(0, 0, "客户").unwrap();
    structural.write_string(0, 2, "客户").unwrap();
    structural.write_string(0, 4, "金额").unwrap();
    structural.write_string(1, 0, "张三").unwrap();
    structural.write_string(1, 1, "A-100").unwrap();
    structural.write_string(1, 2, "张三").unwrap();
    structural.write_number(1, 4, 10).unwrap();
    structural.write_string(3, 0, "李四").unwrap();
    structural.write_string(3, 1, "A-101").unwrap();
    structural.write_string(3, 2, "李四").unwrap();
    structural.write_number(3, 4, 20).unwrap();

    workbook.save(path).unwrap();
}

fn temp_workbook(name: &str) -> PathBuf {
    let path = std::env::temp_dir().join(name);
    create_structural_workbook(&path);
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

fn assert_success(output: &Output, code: i32) {
    assert_eq!(output.status.code(), Some(code));
    assert!(
        output.stderr.is_empty(),
        "stderr should be empty: {}",
        String::from_utf8_lossy(&output.stderr)
    );
}

#[test]
fn structural_rules_report_stable_locations_and_details() {
    let file_path = temp_workbook("excel_cli_check_structural_positive.xlsx");
    let file_arg = file_path.to_string_lossy();
    let output = run_check(&[
        "check",
        file_arg.as_ref(),
        "--sheet",
        "结构",
        "--rules",
        "blank_headers,duplicate_headers,blank_rows,blank_columns",
    ]);

    assert_success(&output, 1);
    let json = parse_stdout(&output);

    assert_eq!(json["target"], json!({"sheet": "结构", "sheet_index": 1}));
    assert_eq!(json["data"]["summary"]["status"], "fail");
    assert_eq!(json["data"]["summary"]["finding_count"], 5);
    assert_eq!(json["data"]["summary"]["warning_count"], 5);
    assert_eq!(json["data"]["stats"]["finding_count_before_threshold"], 5);
    assert_eq!(
        json["data"]["stats"]["rules_run"],
        json!([
            "blank_headers",
            "duplicate_headers",
            "blank_rows",
            "blank_columns"
        ])
    );

    assert_eq!(
        json["data"]["findings"],
        json!([
            {
                "rule_id": "blank_headers",
                "severity": "warning",
                "sheet": "结构",
                "row": 1,
                "column": 2,
                "range": "B1",
                "message": "Blank header at B1.",
                "details": {
                    "header_row": 1,
                    "column_label": "B",
                    "reason": "blank_header"
                }
            },
            {
                "rule_id": "blank_headers",
                "severity": "warning",
                "sheet": "结构",
                "row": 1,
                "column": 4,
                "range": "D1",
                "message": "Blank header at D1.",
                "details": {
                    "header_row": 1,
                    "column_label": "D",
                    "reason": "blank_header"
                }
            },
            {
                "rule_id": "duplicate_headers",
                "severity": "warning",
                "sheet": "结构",
                "row": 1,
                "column": 3,
                "range": "C1",
                "message": "Duplicate header '客户' at C1.",
                "details": {
                    "header": "客户",
                    "normalized_header": "客户",
                    "first_column": 1,
                    "first_range": "A1",
                    "duplicate_count": 2
                }
            },
            {
                "rule_id": "blank_rows",
                "severity": "warning",
                "sheet": "结构",
                "row": 3,
                "column": null,
                "range": "A3:E3",
                "message": "Blank row 3 in used range A1:E4.",
                "details": {
                    "used_range": "A1:E4",
                    "max_columns": 5,
                    "reason": "blank_row"
                }
            },
            {
                "rule_id": "blank_columns",
                "severity": "warning",
                "sheet": "结构",
                "row": null,
                "column": 4,
                "range": "D1:D4",
                "message": "Blank column D in used range A1:E4.",
                "details": {
                    "used_range": "A1:E4",
                    "column_label": "D",
                    "max_rows": 4,
                    "reason": "blank_column"
                }
            }
        ])
    );
}

#[test]
fn structural_rules_have_clean_negative_cases() {
    let file_path = temp_workbook("excel_cli_check_structural_negative.xlsx");
    let file_arg = file_path.to_string_lossy();

    for rule in [
        "blank_headers",
        "duplicate_headers",
        "blank_rows",
        "blank_columns",
    ] {
        let output = run_check(&[
            "check",
            file_arg.as_ref(),
            "--sheet",
            "Clean",
            "--rules",
            rule,
        ]);

        assert_success(&output, 0);
        let json = parse_stdout(&output);
        assert_eq!(json["data"]["summary"]["status"], "pass", "rule: {rule}");
        assert_eq!(json["data"]["findings"], json!([]), "rule: {rule}");
        assert_eq!(
            json["data"]["stats"]["finding_count_before_threshold"], 0,
            "rule: {rule}"
        );
    }
}

#[test]
fn structural_rules_aggregate_workbook_and_filter_by_sheet() {
    let file_path = temp_workbook("excel_cli_check_structural_targets.xlsx");
    let file_arg = file_path.to_string_lossy();

    let workbook = run_check(&[
        "check",
        file_arg.as_ref(),
        "--rules",
        "blank_headers,duplicate_headers,blank_rows,blank_columns",
    ]);
    assert_success(&workbook, 1);
    let json = parse_stdout(&workbook);
    assert_eq!(json["target"], json!({}));
    assert_eq!(json["data"]["stats"]["sheet_count"], 2);
    assert_eq!(json["data"]["stats"]["checked_sheet_count"], 2);
    assert_eq!(json["data"]["summary"]["finding_count"], 5);
    assert!(json["data"]["findings"]
        .as_array()
        .unwrap()
        .iter()
        .all(|finding| finding["sheet"] == "结构"));

    let sheet = run_check(&[
        "check",
        file_arg.as_ref(),
        "--sheet",
        "Clean",
        "--rules",
        "blank_headers,duplicate_headers,blank_rows,blank_columns",
    ]);
    assert_success(&sheet, 0);
    let json = parse_stdout(&sheet);
    assert_eq!(json["target"], json!({"sheet": "Clean", "sheet_index": 0}));
    assert_eq!(json["data"]["stats"]["checked_sheet_count"], 1);
    assert_eq!(json["data"]["summary"]["finding_count"], 0);
    assert_eq!(json["data"]["findings"], json!([]));
}

#[test]
fn severity_threshold_filters_warning_structural_findings() {
    let file_path = temp_workbook("excel_cli_check_structural_threshold.xlsx");
    let file_arg = file_path.to_string_lossy();
    let output = run_check(&[
        "check",
        file_arg.as_ref(),
        "--sheet",
        "结构",
        "--rules",
        "blank_headers,duplicate_headers,blank_rows,blank_columns",
        "--severity-threshold",
        "error",
    ]);

    assert_success(&output, 0);
    let json = parse_stdout(&output);
    assert_eq!(json["data"]["summary"]["status"], "pass");
    assert_eq!(json["data"]["summary"]["finding_count"], 0);
    assert_eq!(json["data"]["stats"]["finding_count_before_threshold"], 5);
    assert_eq!(json["data"]["findings"], json!([]));
}
