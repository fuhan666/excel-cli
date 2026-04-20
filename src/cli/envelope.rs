use serde_json::{json, Value};

/// Build a standard v1.0 success envelope.
pub fn success_envelope(
    command: &str,
    file_path: &str,
    file_format: &str,
    target: Value,
    meta: Value,
    data: Value,
    warnings: Vec<Value>,
) -> Value {
    json!({
        "schema_version": "1.0",
        "command": command,
        "file": {
            "path": file_path,
            "format": file_format,
        },
        "target": target,
        "meta": meta,
        "data": data,
        "warnings": warnings,
    })
}

/// Build a minimal target object for workbook-level queries.
pub fn target_workbook() -> Value {
    json!({})
}

/// Build a target object for sheet-level queries.
pub fn target_sheet(sheet_name: &str, sheet_index: usize) -> Value {
    json!({
        "sheet": sheet_name,
        "sheet_index": sheet_index,
    })
}

/// Build a target object for range-level queries.
pub fn target_range(sheet_name: &str, sheet_index: usize, range: &str) -> Value {
    json!({
        "sheet": sheet_name,
        "sheet_index": sheet_index,
        "range": range,
    })
}

/// Build a target object for cell-level queries.
pub fn target_cell(sheet_name: &str, sheet_index: usize, cell: &str) -> Value {
    json!({
        "sheet": sheet_name,
        "sheet_index": sheet_index,
        "cell": cell,
    })
}
