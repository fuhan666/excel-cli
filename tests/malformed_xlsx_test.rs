use std::path::PathBuf;
use std::process::Command;

fn excel_cli_bin() -> PathBuf {
    let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    path.push("target");
    path.push("debug");
    path.push("excel-cli");
    path
}

fn malformed_fixture_path() -> PathBuf {
    let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    path.push("tests");
    path.push("fixtures");
    path.push("invalid_shared_strings.xlsx");
    path
}

#[test]
fn malformed_xlsx_inspect_workbook_returns_controlled_error() {
    let output = Command::new(excel_cli_bin())
        .arg("inspect")
        .arg("workbook")
        .arg(malformed_fixture_path())
        .output()
        .expect("Failed to execute excel-cli");

    assert!(
        !output.status.success(),
        "Expected failure for malformed workbook, got success"
    );
    let actual = output.status.code().unwrap_or(-1);
    assert_eq!(actual, 4, "Expected exit code 4 (parse_error), got {}", actual);

    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(
        stderr.contains("parse_error") || stderr.contains("parser panic: malformed workbook data"),
        "Expected controlled parser error, got: {}",
        stderr
    );
    // Must not contain a Rust panic backtrace indicator
    assert!(
        !stderr.contains("stack backtrace"),
        "Should not contain a Rust backtrace, got: {}",
        stderr
    );
}

#[test]
fn malformed_xlsx_read_cell_returns_controlled_error() {
    let output = Command::new(excel_cli_bin())
        .arg("read")
        .arg("cell")
        .arg(malformed_fixture_path())
        .arg("--sheet-index")
        .arg("0")
        .arg("--cell")
        .arg("A1")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(
        !output.status.success(),
        "Expected failure for malformed workbook in read mode, got success"
    );
    let actual = output.status.code().unwrap_or(-1);
    assert_eq!(actual, 4, "Expected exit code 4 (parse_error), got {}", actual);

    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(
        stderr.contains("parse_error") || stderr.contains("parser panic: malformed workbook data"),
        "Expected controlled parser error in read mode, got: {}",
        stderr
    );
    assert!(
        !stderr.contains("stack backtrace"),
        "Should not contain a Rust backtrace in read mode, got: {}",
        stderr
    );
}
