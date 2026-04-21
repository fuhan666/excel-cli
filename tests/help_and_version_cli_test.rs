use std::path::PathBuf;
use std::process::Command;

fn excel_cli_bin() -> PathBuf {
    PathBuf::from(env!("CARGO_BIN_EXE_excel-cli"))
}

fn assert_successful_help(args: &[&str]) -> String {
    let output = Command::new(excel_cli_bin())
        .args(args)
        .output()
        .unwrap_or_else(|_| panic!("Failed to execute excel-cli {}", args.join(" ")));

    assert_eq!(output.status.code(), Some(0));
    assert!(
        output.stderr.is_empty(),
        "stderr should be empty: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    String::from_utf8(output.stdout).expect("help output should be valid UTF-8")
}

#[test]
fn top_level_help_prints_to_stdout_and_exits_zero() {
    let stdout = assert_successful_help(&["--help"]);
    assert!(
        stdout.contains("Usage: excel-cli <COMMAND>"),
        "unexpected stdout: {stdout}"
    );
    assert!(stdout.contains("inspect"), "unexpected stdout: {stdout}");
    assert!(stdout.contains("read"), "unexpected stdout: {stdout}");
    assert!(stdout.contains("ui"), "unexpected stdout: {stdout}");
}

#[test]
fn subcommand_help_prints_to_stdout_and_exits_zero() {
    let stdout = assert_successful_help(&["ui", "--help"]);
    assert!(
        stdout.contains("Open interactive TUI browser"),
        "unexpected stdout: {stdout}"
    );
    assert!(
        stdout.contains("Usage: excel-cli ui <FILE>"),
        "unexpected stdout: {stdout}"
    );
}

#[test]
fn read_help_lists_records_subcommand() {
    let stdout = assert_successful_help(&["read", "--help"]);

    assert!(
        stdout.contains("Read records from a sheet using a resolved header row"),
        "unexpected stdout: {stdout}"
    );
    assert!(stdout.contains("records"), "unexpected stdout: {stdout}");
}

#[test]
fn read_rows_help_documents_v12_query_flags() {
    let stdout = assert_successful_help(&["read", "rows", "--help"]);

    for expected in [
        "--select <SELECT>",
        "--filter <FILTERS>",
        "field:op:value",
        "eq|ne|gt|gte|lt|lte|contains|regex|isnull|notnull",
        "--limit <LIMIT>",
        "--offset <OFFSET>",
        "--non-empty",
        "--output-shape <OUTPUT_SHAPE>",
        "rows, records, jsonl",
    ] {
        assert!(
            stdout.contains(expected),
            "expected {expected:?} in stdout: {stdout}"
        );
    }
}

#[test]
fn read_records_help_documents_default_record_shape() {
    let stdout = assert_successful_help(&["read", "records", "--help"]);

    for expected in [
        "Read records from a sheet using a resolved header row",
        "--select <SELECT>",
        "--filter <FILTERS>",
        "records by default",
        "rows, records, jsonl",
    ] {
        assert!(
            stdout.contains(expected),
            "expected {expected:?} in stdout: {stdout}"
        );
    }
}

#[test]
fn version_prints_to_stdout_and_exits_zero() {
    let output = Command::new(excel_cli_bin())
        .arg("--version")
        .output()
        .expect("Failed to execute excel-cli --version");

    assert_eq!(output.status.code(), Some(0));
    assert!(
        output.stderr.is_empty(),
        "stderr should be empty: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let stdout = String::from_utf8_lossy(&output.stdout);
    assert_eq!(stdout.trim(), "excel-cli 1.2.0");
}
