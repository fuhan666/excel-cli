use serde_json::Value;

use crate::cli::args::OutputFormat;
use crate::cli::common::tab_separated_values;
use crate::cli::error::AppError;

/// Write a success value to stdout.
pub fn write_success(value: &Value, format: &OutputFormat) -> Result<(), AppError> {
    if value["meta"]["output_shape"].as_str() == Some("jsonl") {
        return write_jsonl_records(value);
    }

    match format {
        OutputFormat::Json => {
            let s = serde_json::to_string_pretty(value).map_err(|e| AppError::InternalError {
                message: format!("JSON serialization failed: {}", e),
            })?;
            println!("{}", s);
        }
        OutputFormat::Text => {
            write_text(value)?;
        }
        OutputFormat::Markdown => {
            write_markdown(value)?;
        }
    }
    Ok(())
}

fn write_jsonl_records(value: &Value) -> Result<(), AppError> {
    let Some(records) = value["data"]["records"].as_array() else {
        return Err(AppError::InternalError {
            message: "JSONL output requires record data".to_string(),
        });
    };

    for record in records {
        let s = serde_json::to_string(record).map_err(|e| AppError::InternalError {
            message: format!("JSON serialization failed: {}", e),
        })?;
        println!("{}", s);
    }

    Ok(())
}

/// Write an error value to stderr.
pub fn write_error(value: &Value) {
    if let Ok(s) = serde_json::to_string_pretty(value) {
        eprintln!("{}", s);
    } else {
        eprintln!("{{\"error\":\"Internal JSON serialization error\"}}");
    }
}

/// Best-effort text rendering for successful envelopes.
fn write_text(value: &Value) -> Result<(), AppError> {
    let command = value["command"].as_str().unwrap_or("unknown");

    match command {
        "inspect.workbook" => write_text_workbook(value)?,
        "inspect.sheet" => write_text_sheet(value)?,
        "inspect.sample" => write_text_sample(value)?,
        "inspect.columns" => write_text_columns(value)?,
        "inspect.tables" => write_text_tables(value)?,
        "read.cell" => write_text_cell(value)?,
        "read.range" => write_text_range(value)?,
        "read.rows" => write_text_rows(value)?,
        _ => {
            // Fallback to JSON for unknown commands
            let s = serde_json::to_string_pretty(value).map_err(|e| AppError::InternalError {
                message: format!("JSON serialization failed: {}", e),
            })?;
            println!("{}", s);
        }
    }

    Ok(())
}

fn write_text_workbook(value: &Value) -> Result<(), AppError> {
    let data = &value["data"];
    if let Some(sheets) = data["sheets"].as_array() {
        for sheet in sheets {
            let name = sheet["name"].as_str().unwrap_or("");
            let index = sheet["index"].as_u64().unwrap_or(0);
            println!("{}	{}", index, name);
        }
    }
    Ok(())
}

fn write_text_sheet(value: &Value) -> Result<(), AppError> {
    let data = &value["data"];
    if let Some(name) = data["name"].as_str() {
        println!("name\t{}", name);
    }
    if let Some(index) = data["index"].as_u64() {
        println!("index\t{}", index);
    }
    if let Some(max_rows) = data["max_rows"].as_u64() {
        println!("max_rows\t{}", max_rows);
    }
    if let Some(max_cols) = data["max_cols"].as_u64() {
        println!("max_cols\t{}", max_cols);
    }
    if let Some(used_range) = data["used_range"].as_str() {
        println!("used_range\t{}", used_range);
    }
    Ok(())
}

fn write_text_sample(value: &Value) -> Result<(), AppError> {
    let data = &value["data"];
    if let Some(rows) = data["rows"].as_array() {
        write_row_arrays(rows);
    } else if let Some(records) = data["records"].as_array() {
        write_record_objects(records);
    }
    Ok(())
}

fn write_text_columns(value: &Value) -> Result<(), AppError> {
    println!("index\tname\tsafe_name\tis_duplicate\tinferred_type\tnon_null_ratio\tformula_ratio");

    if let Some(columns) = value["data"]["columns"].as_array() {
        for column in columns {
            let index = column["index"].as_u64().unwrap_or(0);
            let name = column["name"].as_str().unwrap_or("");
            let safe_name = column["safe_name"].as_str().unwrap_or("");
            let is_duplicate = column["is_duplicate"].as_bool().unwrap_or(false);
            let inferred_type = column["inferred_type"].as_str().unwrap_or("");
            let non_null_ratio = format_ratio(column["non_null_ratio"].as_f64().unwrap_or(0.0));
            let formula_ratio = format_ratio(column["formula_ratio"].as_f64().unwrap_or(0.0));

            println!(
                "{}\t{}\t{}\t{}\t{}\t{}\t{}",
                index, name, safe_name, is_duplicate, inferred_type, non_null_ratio, formula_ratio
            );
        }
    }

    Ok(())
}

fn format_ratio(value: f64) -> String {
    let rounded = (value * 1000.0).round() / 1000.0;
    if (rounded.fract()).abs() < f64::EPSILON {
        format!("{rounded:.0}")
    } else {
        let formatted = format!("{rounded:.3}");
        formatted
            .trim_end_matches('0')
            .trim_end_matches('.')
            .to_string()
    }
}

fn write_text_tables(value: &Value) -> Result<(), AppError> {
    let data = &value["data"];
    if let Some(candidates) = data["candidates"].as_array() {
        for candidate in candidates {
            let range = candidate["range"].as_str().unwrap_or("");
            let header_row = candidate["header_row"].as_u64().unwrap_or(0);
            let column_count = candidate["column_count"].as_u64().unwrap_or(0);
            let row_count = candidate["row_count"].as_u64().unwrap_or(0);
            let confidence = candidate["confidence"].as_f64().unwrap_or(0.0);
            println!(
                "{}\theader_row={}\tcolumns={}\trows={}\tconfidence={:.2}",
                range, header_row, column_count, row_count, confidence
            );
        }
    }
    Ok(())
}

fn write_text_cell(value: &Value) -> Result<(), AppError> {
    if let Some(v) = value["data"]["value"].as_str() {
        println!("{}", v);
    } else {
        println!("{}", value["data"]["value"]);
    }
    Ok(())
}

fn write_text_range(value: &Value) -> Result<(), AppError> {
    let data = &value["data"];
    if let Some(rows) = data["rows"].as_array() {
        write_row_arrays(rows);
    }
    Ok(())
}

fn write_text_rows(value: &Value) -> Result<(), AppError> {
    let data = &value["data"];
    if let Some(rows) = data["rows"].as_array() {
        write_row_arrays(rows);
    } else if let Some(records) = data["records"].as_array() {
        write_record_objects(records);
    }
    Ok(())
}

fn write_row_arrays(rows: &[Value]) {
    for row in rows {
        if let Some(cells) = row.as_array() {
            println!("{}", tab_separated_values(cells));
        }
    }
}

fn write_record_objects(records: &[Value]) {
    for record in records {
        if let Some(obj) = record.as_object() {
            let parts: Vec<String> = obj.iter().map(|(k, v)| format!("{}={}", k, v)).collect();
            println!("{}", parts.join("\t"));
        }
    }
}

fn write_markdown(value: &Value) -> Result<(), AppError> {
    let command = value["command"].as_str().unwrap_or("unknown");

    match command {
        "read.rows" | "read.records" | "inspect.sample" => {
            let data = &value["data"];
            let meta = &value["meta"];

            if let Some(records) = data["records"].as_array() {
                let selected_cols: Vec<String> =
                    if let Some(cols) = meta["selected_columns"].as_array() {
                        cols.iter()
                            .map(|v| v.as_str().unwrap_or("").to_string())
                            .collect()
                    } else if !records.is_empty() {
                        records[0]
                            .as_object()
                            .map(|obj| obj.keys().cloned().collect())
                            .unwrap_or_default()
                    } else {
                        Vec::new()
                    };

                if selected_cols.is_empty() {
                    return Ok(());
                }

                let escaped_headers: Vec<String> = selected_cols
                    .iter()
                    .map(|s| escape_markdown_header(s))
                    .collect();
                print_markdown_row(&escaped_headers);
                let delimiters: Vec<String> = vec!["---".to_string(); selected_cols.len()];
                print_markdown_row(&delimiters);

                for record in records {
                    if let Some(obj) = record.as_object() {
                        let row_vals: Vec<String> = selected_cols
                            .iter()
                            .map(|col| {
                                let val = obj.get(col).unwrap_or(&Value::Null);
                                format_markdown_cell(val)
                            })
                            .collect();
                        print_markdown_row(&row_vals);
                    }
                }
            } else if let Some(rows) = data["rows"].as_array() {
                if rows.is_empty() {
                    return Ok(());
                }

                let selected_cols: Vec<String> =
                    if let Some(cols) = meta["selected_columns"].as_array() {
                        cols.iter()
                            .map(|v| v.as_str().unwrap_or("").to_string())
                            .collect()
                    } else {
                        Vec::new()
                    };

                if !selected_cols.is_empty() {
                    let escaped_headers: Vec<String> = selected_cols
                        .iter()
                        .map(|s| escape_markdown_header(s))
                        .collect();
                    print_markdown_row(&escaped_headers);
                    let delimiters: Vec<String> = vec!["---".to_string(); selected_cols.len()];
                    print_markdown_row(&delimiters);
                }

                for row in rows {
                    if let Some(cells) = row.as_array() {
                        let row_vals: Vec<String> =
                            cells.iter().map(format_markdown_cell).collect();
                        print_markdown_row(&row_vals);
                    }
                }
            }
        }
        "read.cell" => {
            if let Some(v) = value["data"]["value"].as_str() {
                println!("{}", v);
            } else {
                println!("{}", value["data"]["value"]);
            }
        }
        "read.range" => {
            let data = &value["data"];
            if let Some(rows) = data["rows"].as_array() {
                if rows.is_empty() {
                    return Ok(());
                }

                if let Some(first_row) = rows[0].as_array() {
                    let col_count = first_row.len();
                    let headers: Vec<String> =
                        (1..=col_count).map(|i| format!("Col {i}")).collect();
                    print_markdown_row(&headers);

                    let delimiters: Vec<String> = vec!["---".to_string(); col_count];
                    print_markdown_row(&delimiters);
                }

                for row in rows {
                    if let Some(cells) = row.as_array() {
                        let row_vals: Vec<String> =
                            cells.iter().map(format_markdown_cell).collect();
                        print_markdown_row(&row_vals);
                    }
                }
            }
        }
        "grep" => {
            let headers = vec![
                "file".to_string(),
                "sheet".to_string(),
                "cell".to_string(),
                "content".to_string(),
            ];
            print_markdown_row(&headers);
            let delimiters: Vec<String> = vec!["---".to_string(); 4];
            print_markdown_row(&delimiters);

            if let Some(matches) = value["data"]["matches"].as_array() {
                for m in matches {
                    let row_vals = vec![
                        markdown_table_cell_text(m["file"].as_str().unwrap_or("")),
                        markdown_table_cell_text(m["sheet"].as_str().unwrap_or("")),
                        markdown_table_cell_text(m["cell"].as_str().unwrap_or("")),
                        markdown_table_cell_text(m["content"].as_str().unwrap_or("")),
                    ];
                    print_markdown_row(&row_vals);
                }
            }
        }
        _ => {
            write_text(value)?;
        }
    }

    Ok(())
}

fn print_markdown_row(vals: &[String]) {
    println!("| {} |", vals.join(" | "));
}

/// Escape text for use inside a Markdown table cell.
fn markdown_table_cell_text(s: &str) -> String {
    s.replace('\r', "")
        .replace('\n', "<br>")
        .replace('|', "\\|")
}

fn format_markdown_cell(val: &Value) -> String {
    match val {
        Value::Null => String::new(),
        Value::String(s) => markdown_table_cell_text(s),
        Value::Number(n) => n.to_string(),
        Value::Bool(b) => {
            if *b {
                "TRUE".to_string()
            } else {
                "FALSE".to_string()
            }
        }
        other => markdown_table_cell_text(&other.to_string()),
    }
}

fn escape_markdown_header(s: &str) -> String {
    s.replace('\r', "").replace('\n', " ").replace('|', "\\|")
}
