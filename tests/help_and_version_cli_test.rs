use std::path::PathBuf;
use std::process::Command;

fn excel_cli_bin() -> PathBuf {
    PathBuf::from(env!("CARGO_BIN_EXE_excel-cli"))
}

#[test]
fn top_level_help_prints_to_stdout_and_exits_zero() {
    let output = Command::new(excel_cli_bin())
        .arg("--help")
        .output()
        .expect("Failed to execute excel-cli --help");

    assert_eq!(output.status.code(), Some(0));
    assert!(
        output.stderr.is_empty(),
        "stderr should be empty: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let stdout = String::from_utf8_lossy(&output.stdout);
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
    let output = Command::new(excel_cli_bin())
        .arg("ui")
        .arg("--help")
        .output()
        .expect("Failed to execute excel-cli ui --help");

    assert_eq!(output.status.code(), Some(0));
    assert!(
        output.stderr.is_empty(),
        "stderr should be empty: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let stdout = String::from_utf8_lossy(&output.stdout);
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
    assert_eq!(
        stdout.trim(),
        format!("excel-cli {}", env!("CARGO_PKG_VERSION"))
    );
}
