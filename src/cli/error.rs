use serde_json::{json, Value};
use std::fmt;

/// Exit codes for the v1.0.0 CLI contract.
pub const EXIT_SUCCESS: i32 = 0;
pub const EXIT_CHECK_FINDINGS: i32 = 1;
pub const EXIT_INVALID_ARGS: i32 = 2;
pub const EXIT_FILE_ERROR: i32 = 3;
pub const EXIT_PARSE_ERROR: i32 = 4;
pub const EXIT_TARGET_NOT_FOUND: i32 = 5;
pub const EXIT_INVALID_QUERY: i32 = 6;
pub const EXIT_INTERNAL_ERROR: i32 = 7;

/// Application errors with stable codes and exit code mapping.
#[derive(Debug, Clone)]
pub enum AppError {
    InvalidArgs { message: String },
    FileError { message: String },
    ParseError { message: String },
    TargetNotFound { message: String },
    InvalidQuery { message: String },
    InternalError { message: String },
}

impl AppError {
    pub fn code(&self) -> &'static str {
        match self {
            AppError::InvalidArgs { .. } => "invalid_args",
            AppError::FileError { .. } => "file_error",
            AppError::ParseError { .. } => "parse_error",
            AppError::TargetNotFound { .. } => "target_not_found",
            AppError::InvalidQuery { .. } => "invalid_query",
            AppError::InternalError { .. } => "internal_error",
        }
    }

    pub fn exit_code(&self) -> i32 {
        match self {
            AppError::InvalidArgs { .. } => EXIT_INVALID_ARGS,
            AppError::FileError { .. } => EXIT_FILE_ERROR,
            AppError::ParseError { .. } => EXIT_PARSE_ERROR,
            AppError::TargetNotFound { .. } => EXIT_TARGET_NOT_FOUND,
            AppError::InvalidQuery { .. } => EXIT_INVALID_QUERY,
            AppError::InternalError { .. } => EXIT_INTERNAL_ERROR,
        }
    }

    pub fn message(&self) -> String {
        match self {
            AppError::InvalidArgs { message } => message.clone(),
            AppError::FileError { message } => message.clone(),
            AppError::ParseError { message } => message.clone(),
            AppError::TargetNotFound { message } => message.clone(),
            AppError::InvalidQuery { message } => message.clone(),
            AppError::InternalError { message } => message.clone(),
        }
    }

    /// Build the standard error envelope as JSON.
    pub fn to_envelope(&self, command: &str, file_path: &str, file_format: &str) -> Value {
        json!({
            "schema_version": "1.0",
            "command": command,
            "file": {
                "path": file_path,
                "format": file_format,
            },
            "error": {
                "code": self.code(),
                "message": self.message(),
                "details": {},
            },
        })
    }
}

impl fmt::Display for AppError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.message())
    }
}

impl std::error::Error for AppError {}

/// Convert anyhow::Error to AppError by inspecting the message.
pub fn anyhow_to_app_error(err: anyhow::Error) -> AppError {
    let msg = err.to_string();
    let lower = msg.to_lowercase();

    if err
        .chain()
        .any(|cause| cause.downcast_ref::<std::io::Error>().is_some())
    {
        AppError::FileError { message: msg }
    } else if lower.contains("unable to parse excel file")
        || lower.contains("parser panic: malformed workbook data")
        || lower.contains("no worksheets found")
    {
        AppError::ParseError { message: msg }
    } else if lower.contains("unable to read worksheet")
        || lower.contains("cannot load sheet")
        || (lower.contains("sheet") && lower.contains("not found"))
    {
        AppError::TargetNotFound { message: msg }
    } else if lower.contains("invalid")
        || lower.contains("mutually exclusive")
        || lower.contains("expected")
    {
        AppError::InvalidArgs { message: msg }
    } else {
        AppError::InternalError { message: msg }
    }
}
