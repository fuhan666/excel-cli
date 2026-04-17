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
fn malformed_xlsx_non_lazy_returns_controlled_error() {
    let output = Command::new(excel_cli_bin())
        .arg(malformed_fixture_path())
        .arg("--sheets")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(
        !output.status.success(),
        "Expected failure for malformed workbook, got success"
    );
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(
        stderr.contains("parser panic: malformed workbook data"),
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
fn malformed_xlsx_lazy_loading_returns_controlled_error() {
    let output = Command::new(excel_cli_bin())
        .arg(malformed_fixture_path())
        .arg("-l")
        .arg("--peek")
        .arg("0!A1:B2")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(
        !output.status.success(),
        "Expected failure for malformed workbook in lazy mode, got success"
    );
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(
        stderr.contains("parser panic: malformed workbook data"),
        "Expected controlled parser error in lazy mode, got: {}",
        stderr
    );
    assert!(
        !stderr.contains("stack backtrace"),
        "Should not contain a Rust backtrace in lazy mode, got: {}",
        stderr
    );
}
