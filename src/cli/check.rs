use serde_json::Value;

use crate::cli::error::AppError;

pub fn handle(file: std::path::PathBuf, rule: Option<String>) -> Result<Value, AppError> {
    let _format_str = file_format(&file);
    let _path_str = file.to_string_lossy().to_string();

    // v1.0.0: check is namespace-only. Any attempt to run a rule is rejected.
    let message = if let Some(r) = rule {
        format!("Check rule '{}' is not implemented in v1.0.0. Quality checks are planned for v1.3.0.", r)
    } else {
        "Check command requires a --rule argument. Quality checks are planned for v1.3.0.".to_string()
    };

    Err(AppError::CheckNotImplemented { message })
}

fn file_format(path: &std::path::Path) -> String {
    path.extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_lowercase())
        .unwrap_or_else(|| "unknown".to_string())
}
