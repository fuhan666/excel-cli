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
