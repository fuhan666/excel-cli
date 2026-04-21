use std::path::{Path, PathBuf};
use std::process::{Command, Output};

use serde_json::{json, Value};

fn excel_cli_bin() -> PathBuf {
    PathBuf::from(env!("CARGO_BIN_EXE_excel-cli"))
}

fn temp_path(name: &str) -> PathBuf {
    std::env::temp_dir().join(name)
}

fn create_workbook(path: &Path) {
    use rust_xlsxwriter::Workbook as XlsxWorkbook;

    let mut workbook = XlsxWorkbook::new();

    let structural = workbook.add_worksheet();
    structural.set_name("Structural").unwrap();
    structural.write_string(0, 0, "customer").unwrap();
    structural.write_string(0, 2, "customer").unwrap();
    structural.write_string(1, 0, "Alice").unwrap();
    structural.write_string(1, 1, "A-100").unwrap();
    structural.write_string(1, 2, "Alice").unwrap();

    let analytical = workbook.add_worksheet();
    analytical.set_name("Analytical").unwrap();
    analytical.write_string(0, 0, "order_id").unwrap();
    analytical.write_string(0, 1, "amount").unwrap();
    analytical.write_string(0, 2, "total").unwrap();
    analytical.write_string(1, 0, "1001").unwrap();
    analytical.write_number(1, 1, 12.5).unwrap();
    analytical.write_formula(1, 2, "=B2*2").unwrap();
    analytical.write_string(2, 0, "1001").unwrap();
    analytical.write_number(2, 1, 18.0).unwrap();
    analytical.write_formula(2, 2, "=B3*2").unwrap();

    workbook.save(path).unwrap();
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
fn workbook_check_combines_structural_and_analytical_findings() {
    let file_path = temp_path("excel_cli_check_rule_integration.xlsx");
    create_workbook(&file_path);
    let file_arg = file_path.to_string_lossy();

    let output = run_check(&[
        "check",
        file_arg.as_ref(),
        "--rules",
        "blank_headers,duplicate_headers,duplicate_values,formula_presence",
    ]);

    assert_success(&output, 1);
    let json = parse_stdout(&output);

    assert_eq!(json["target"], json!({}));
    assert_eq!(
        json["data"]["summary"],
        json!({
            "status": "fail",
            "finding_count": 4,
            "error_count": 0,
            "warning_count": 3,
            "info_count": 1
        })
    );
    assert_eq!(
        json["data"]["stats"]["rules_run"],
        json!([
            "blank_headers",
            "duplicate_headers",
            "duplicate_values",
            "formula_presence"
        ])
    );
    assert_eq!(json["data"]["stats"]["checked_sheet_count"], 2);
    assert_eq!(json["data"]["stats"]["finding_count_before_threshold"], 4);
    assert_eq!(
        json["data"]["findings"]
            .as_array()
            .unwrap()
            .iter()
            .map(|finding| (
                finding["sheet"].as_str().unwrap(),
                finding["rule_id"].as_str().unwrap()
            ))
            .collect::<Vec<_>>(),
        vec![
            ("Structural", "blank_headers"),
            ("Structural", "duplicate_headers"),
            ("Analytical", "duplicate_values"),
            ("Analytical", "formula_presence"),
        ]
    );
}
