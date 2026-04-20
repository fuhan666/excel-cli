use serde_json::Value;

use crate::cli::args::OutputFormat;
use crate::cli::error::AppError;

/// Write a success value to stdout.
pub fn write_success(value: &Value, format: &OutputFormat) -> Result<(), AppError> {
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
        for row in rows {
            if let Some(cells) = row.as_array() {
                let line: Vec<String> = cells
                    .iter()
                    .map(|c| match c {
                        Value::Null => String::new(),
                        Value::String(s) => s.clone(),
                        _ => c.to_string(),
                    })
                    .collect();
                println!("{}", line.join("\t"));
            }
        }
    } else if let Some(records) = data["records"].as_array() {
        for record in records {
            if let Some(obj) = record.as_object() {
                let parts: Vec<String> = obj
                    .iter()
                    .map(|(k, v)| format!("{}={}", k, v))
                    .collect();
                println!("{}", parts.join("\t"));
            }
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
        for row in rows {
            if let Some(cells) = row.as_array() {
                let line: Vec<String> = cells
                    .iter()
                    .map(|c| match c {
                        Value::Null => String::new(),
                        Value::String(s) => s.clone(),
                        _ => c.to_string(),
                    })
                    .collect();
                println!("{}", line.join("\t"));
            }
        }
    }
    Ok(())
}

fn write_text_rows(value: &Value) -> Result<(), AppError> {
    let data = &value["data"];
    if let Some(rows) = data["rows"].as_array() {
        for row in rows {
            if let Some(cells) = row.as_array() {
                let line: Vec<String> = cells
                    .iter()
                    .map(|c| match c {
                        Value::Null => String::new(),
                        Value::String(s) => s.clone(),
                        _ => c.to_string(),
                    })
                    .collect();
                println!("{}", line.join("\t"));
            }
        }
    } else if let Some(records) = data["records"].as_array() {
        for record in records {
            if let Some(obj) = record.as_object() {
                let parts: Vec<String> = obj
                    .iter()
                    .map(|(k, v)| format!("{}={}", k, v))
                    .collect();
                println!("{}", parts.join("\t"));
            }
        }
    }
    Ok(())
}
