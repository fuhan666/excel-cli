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

    // Sheet 1: Orders
    let sheet = workbook.add_worksheet();
    sheet.set_name("Orders").unwrap();
    sheet.write_string(0, 0, "order_id").unwrap();
    sheet.write_string(0, 1, "customer").unwrap();
    sheet.write_string(0, 2, "note").unwrap();

    sheet.write_string(1, 0, "1001").unwrap();
    sheet.write_string(1, 1, "Alice").unwrap();
    sheet.write_string(1, 2, "First line\nSecond line").unwrap();

    sheet.write_string(2, 0, "1002").unwrap();
    sheet.write_string(2, 1, "Bob").unwrap();
    sheet.write_string(2, 2, "Simple note").unwrap();

    workbook.save(path).unwrap();
}

fn create_pipe_workbook(path: &std::path::Path) {
    use rust_xlsxwriter::Workbook as XlsxWorkbook;

    let mut workbook = XlsxWorkbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Data").unwrap();
    sheet.write_string(0, 0, "col|A").unwrap();
    sheet.write_string(0, 1, "colB").unwrap();
    sheet.write_string(1, 0, "a|b").unwrap();
    sheet.write_string(1, 1, "c|d|e").unwrap();
    workbook.save(path).unwrap();
}

#[test]
fn test_read_rows_markdown_format() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_markdown.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("read")
        .arg("rows")
        .arg(&file_path)
        .arg("--sheet")
        .arg("Orders")
        .arg("--format")
        .arg("markdown")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(output.status.success());
    let stdout = String::from_utf8_lossy(&output.stdout);

    // Check Markdown table output structure
    assert!(stdout.contains("| order_id | customer | note |"));
    assert!(stdout.contains("| --- | --- | --- |"));
    assert!(stdout.contains("| 1001 | Alice | First line<br>Second line |"));
    assert!(stdout.contains("| 1002 | Bob | Simple note |"));
}

#[test]
fn test_read_rows_markdown_short_flag() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_markdown_short.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("read")
        .arg("rows")
        .arg(&file_path)
        .arg("--sheet")
        .arg("Orders")
        .arg("-f")
        .arg("markdown")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(output.status.success());
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("| order_id | customer | note |"));
    assert!(stdout.contains("| --- | --- | --- |"));
}

#[test]
fn test_markdown_pipe_escaping() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_pipe.xlsx");
    create_pipe_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("read")
        .arg("records")
        .arg(&file_path)
        .arg("--sheet")
        .arg("Data")
        .arg("-f")
        .arg("markdown")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(output.status.success());
    let stdout = String::from_utf8_lossy(&output.stdout);

    // Pipes in header and cell values should be escaped
    assert!(stdout.contains("col\\|A"));
    assert!(stdout.contains("a\\|b"));
    assert!(stdout.contains("c\\|d\\|e"));
}

#[test]
fn test_grep_command_default_is_markdown() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_grep_default.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("grep")
        .arg("Alice")
        .arg(&file_path)
        .output()
        .expect("Failed to execute excel-cli");

    assert!(output.status.success());
    let stdout = String::from_utf8_lossy(&output.stdout);
    // Default grep output is Markdown table
    assert!(stdout.contains("| file | sheet | cell | content |"));
    assert!(stdout.contains("| --- | --- | --- | --- |"));
    assert!(stdout.contains("Orders"));
    assert!(stdout.contains("B2"));
    assert!(stdout.contains("Alice"));
}

#[test]
fn test_grep_json_machine_output() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_grep_json.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("grep")
        .arg("Alice")
        .arg(&file_path)
        .arg("-f")
        .arg("json")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(output.status.success());
    let stdout = String::from_utf8_lossy(&output.stdout);
    let value: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");
    assert_eq!(value["command"], "grep");

    let matches = value["data"]["matches"].as_array().expect("matches array");
    assert!(!matches.is_empty());
    assert_eq!(matches[0]["sheet"], "Orders");
    assert_eq!(matches[0]["cell"], "B2");
    assert_eq!(matches[0]["content"], "Alice");
}

#[test]
fn test_grep_case_insensitive_and_regex() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_grep_ci_regex.xlsx");
    create_test_workbook(&file_path);

    let output_ci = Command::new(excel_cli_bin())
        .arg("grep")
        .arg("alice")
        .arg(&file_path)
        .arg("-i")
        .arg("-f")
        .arg("json")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(output_ci.status.success());
    let stdout_ci = String::from_utf8_lossy(&output_ci.stdout);
    let value_ci: serde_json::Value = serde_json::from_str(&stdout_ci).expect("valid JSON");
    let matches_ci = value_ci["data"]["matches"]
        .as_array()
        .expect("matches array");
    assert_eq!(matches_ci[0]["sheet"], "Orders");
    assert_eq!(matches_ci[0]["cell"], "B2");
    assert_eq!(matches_ci[0]["content"], "Alice");

    let output_re = Command::new(excel_cli_bin())
        .arg("grep")
        .arg("^100[1-2]$")
        .arg(&file_path)
        .arg("-r")
        .arg("-f")
        .arg("json")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(output_re.status.success());
    let stdout_re = String::from_utf8_lossy(&output_re.stdout);
    let value_re: serde_json::Value = serde_json::from_str(&stdout_re).expect("valid JSON");
    let matches_re = value_re["data"]["matches"]
        .as_array()
        .expect("matches array");
    let cells: Vec<String> = matches_re
        .iter()
        .map(|m| m["cell"].as_str().unwrap_or("").to_string())
        .collect();
    assert!(cells.contains(&"A2".to_string()));
    assert!(cells.contains(&"A3".to_string()));

    let output_none = Command::new(excel_cli_bin())
        .arg("grep")
        .arg("NonExistent")
        .arg(&file_path)
        .output()
        .expect("Failed to execute excel-cli");

    assert_eq!(output_none.status.code(), Some(1));
}

#[test]
fn test_grep_no_stderr_for_warnings() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_grep_noerr.xlsx");
    create_test_workbook(&file_path);

    // Search with a valid file + a non-existent path
    let output = Command::new(excel_cli_bin())
        .arg("grep")
        .arg("Alice")
        .arg(&file_path)
        .arg("/nonexistent/path/that/does/not/exist")
        .arg("-f")
        .arg("json")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(output.status.success());
    // stderr should be empty (warnings go to JSON envelope, not stderr)
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(
        stderr.is_empty(),
        "Expected empty stderr but got: {}",
        stderr
    );

    // stdout should contain the match and the warning in JSON
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("Alice"));
    assert!(stdout.contains("Path not found"));
}

#[test]
fn test_grep_stable_sort_order() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_grep_sort.xlsx");

    // Create workbook with matches in multiple cells using different row positions
    use rust_xlsxwriter::Workbook as XlsxWorkbook;
    let mut workbook = XlsxWorkbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Sort").unwrap();
    sheet.write_string(0, 0, "targetA").unwrap(); // A1
    sheet.write_string(2, 0, "targetB").unwrap(); // A3
    sheet.write_string(5, 0, "targetC").unwrap(); // A6
    workbook.save(&file_path).unwrap();

    let output = Command::new(excel_cli_bin())
        .arg("grep")
        .arg("target")
        .arg(&file_path)
        .arg("-f")
        .arg("json")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(output.status.success());
    let stdout = String::from_utf8_lossy(&output.stdout);
    let value: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");
    let matches = value["data"]["matches"].as_array().expect("matches array");
    assert_eq!(matches.len(), 3);

    let cells: Vec<String> = matches
        .iter()
        .map(|m| m["cell"].as_str().unwrap_or("").to_string())
        .collect();
    assert_eq!(cells, vec!["A1", "A3", "A6"]);
}

#[test]
fn test_grep_markdown_format() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_grep_md.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("grep")
        .arg("Alice")
        .arg(&file_path)
        .arg("-f")
        .arg("markdown")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(output.status.success());
    let stdout = String::from_utf8_lossy(&output.stdout);

    // Should contain markdown table structure
    assert!(stdout.contains("| file | sheet | cell | content |"));
    assert!(stdout.contains("| --- | --- | --- | --- |"));
    assert!(stdout.contains("Orders"));
    assert!(stdout.contains("B2"));
    assert!(stdout.contains("Alice"));
}

#[test]
fn test_grep_multiline_markdown_escaping() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_grep_escaped.xlsx");
    create_test_workbook(&file_path);

    // Search for "First" which matches "First line\nSecond line"
    let output = Command::new(excel_cli_bin())
        .arg("grep")
        .arg("First")
        .arg(&file_path)
        .output()
        .expect("Failed to execute excel-cli");

    assert!(output.status.success());
    let stdout = String::from_utf8_lossy(&output.stdout);

    // Default grep output is Markdown table
    assert!(stdout.contains("| file | sheet | cell | content |"));
    assert!(stdout.contains("| --- | --- | --- | --- |"));
    assert!(stdout.contains("Orders"));
    assert!(stdout.contains("First line<br>Second line"));

    // Should contain exactly one match row (header, delimiter, data row)
    let data_lines: Vec<&str> = stdout
        .trim()
        .lines()
        .filter(|line| line.starts_with("| ") && !line.contains("file") && !line.contains("---"))
        .collect();
    assert_eq!(data_lines.len(), 1);
}

#[test]
fn test_grep_text_format_rejected() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_grep_text_reject.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("grep")
        .arg("Alice")
        .arg(&file_path)
        .arg("-f")
        .arg("text")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(!output.status.success());
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(
        stderr.contains("grep no longer supports --format text"),
        "Expected rejection message in stderr, got: {}",
        stderr
    );
}

#[test]
fn test_jsonl_with_markdown_rejected() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("excel_cli_test_jsonl_markdown.xlsx");
    create_test_workbook(&file_path);

    let output = Command::new(excel_cli_bin())
        .arg("read")
        .arg("records")
        .arg(&file_path)
        .arg("--sheet")
        .arg("Orders")
        .arg("--output-shape")
        .arg("jsonl")
        .arg("-f")
        .arg("markdown")
        .output()
        .expect("Failed to execute excel-cli");

    assert!(!output.status.success());
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(
        stderr.contains("jsonl cannot be combined with"),
        "Expected JSONL format rejection in stderr, got: {}",
        stderr
    );
}
