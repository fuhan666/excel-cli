use clap::{Parser, Subcommand};
use std::path::PathBuf;

#[derive(Parser)]
#[command(name = "excel-cli")]
#[command(author, version, about = "Excel CLI - single-file read-only inspector", long_about = None)]
pub struct Cli {
    #[command(subcommand)]
    pub command: Commands,
}

#[derive(Subcommand)]
pub enum Commands {
    /// Inspect workbook or sheet metadata
    Inspect {
        #[command(subcommand)]
        subcommand: InspectCommands,
    },
    /// Read cell, range, or row data
    Read {
        #[command(subcommand)]
        subcommand: ReadCommands,
    },
    /// Check workbook or sheet data quality
    Check {
        /// Excel file path
        file: PathBuf,

        /// Sheet name (exact match)
        #[arg(long)]
        sheet: Option<String>,

        /// Check rules to run, comma-separated
        #[arg(long)]
        rules: Option<String>,

        /// Minimum finding severity to return
        #[arg(long, value_enum, default_value = "info")]
        severity_threshold: SeverityThreshold,
    },
    /// Open interactive TUI browser
    Ui {
        /// Excel file path
        file: PathBuf,
    },
}

#[derive(Subcommand)]
pub enum InspectCommands {
    /// List all sheets in the workbook
    Workbook {
        /// Excel file path
        file: PathBuf,

        /// Output format
        #[arg(long, value_enum, default_value = "json")]
        format: OutputFormat,
    },
    /// Inspect a single sheet
    Sheet {
        /// Excel file path
        file: PathBuf,

        /// Sheet name (exact match)
        #[arg(long, group = "sheet_target")]
        sheet: Option<String>,

        /// Sheet index (0-based)
        #[arg(long, group = "sheet_target")]
        sheet_index: Option<usize>,

        /// Output format
        #[arg(long, value_enum, default_value = "json")]
        format: OutputFormat,
    },
    /// Sample data from a sheet
    Sample {
        /// Excel file path
        file: PathBuf,

        /// Sheet name (exact match)
        #[arg(long, group = "sheet_target")]
        sheet: Option<String>,

        /// Sheet index (0-based)
        #[arg(long, group = "sheet_target")]
        sheet_index: Option<usize>,

        /// Range to sample (A1 notation)
        #[arg(long)]
        range: Option<String>,

        /// Number of rows to sample
        #[arg(long)]
        rows: Option<usize>,

        /// Header row: auto or 1-based index
        #[arg(long, default_value = "auto")]
        header_row: String,

        /// Output format
        #[arg(long, value_enum, default_value = "json")]
        format: OutputFormat,
    },
    /// Inspect column headers and inferred column metadata
    Columns {
        /// Excel file path
        file: PathBuf,

        /// Sheet name (exact match)
        #[arg(long)]
        sheet: String,

        /// Header row: auto or 1-based index
        #[arg(long, default_value = "auto")]
        header_row: String,

        /// Output format
        #[arg(long, value_enum, default_value = "json")]
        format: OutputFormat,
    },
    /// Detect table-like regions in a sheet
    Tables {
        /// Excel file path
        file: PathBuf,

        /// Sheet name (exact match)
        #[arg(long)]
        sheet: String,

        /// Output format
        #[arg(long, value_enum, default_value = "json")]
        format: OutputFormat,
    },
}

#[derive(Subcommand)]
pub enum ReadCommands {
    /// Read a single cell
    Cell {
        /// Excel file path
        file: PathBuf,

        /// Sheet name (exact match)
        #[arg(long, group = "sheet_target")]
        sheet: Option<String>,

        /// Sheet index (0-based)
        #[arg(long, group = "sheet_target")]
        sheet_index: Option<usize>,

        /// Cell reference (A1 notation)
        #[arg(long)]
        cell: String,

        /// Output format
        #[arg(long, value_enum, default_value = "json")]
        format: OutputFormat,
    },
    /// Read a rectangular range
    Range {
        /// Excel file path
        file: PathBuf,

        /// Sheet name (exact match)
        #[arg(long, group = "sheet_target")]
        sheet: Option<String>,

        /// Sheet index (0-based)
        #[arg(long, group = "sheet_target")]
        sheet_index: Option<usize>,

        /// Range (A1 notation)
        #[arg(long)]
        range: String,

        /// Output format
        #[arg(long, value_enum, default_value = "json")]
        format: OutputFormat,
    },
    /// Read rows from a sheet
    Rows {
        /// Excel file path
        file: PathBuf,

        /// Sheet name (exact match)
        #[arg(long, group = "sheet_target")]
        sheet: Option<String>,

        /// Sheet index (0-based)
        #[arg(long, group = "sheet_target")]
        sheet_index: Option<usize>,

        /// Range to read (A1 notation)
        #[arg(long)]
        range: Option<String>,

        /// Header row: auto or 1-based index
        #[arg(long, default_value = "auto")]
        header_row: String,

        /// Select columns by stable column name, comma-separated
        #[arg(long)]
        select: Option<String>,

        /// Filter rows using field:op:value; operators: eq|ne|gt|gte|lt|lte|contains|regex|isnull|notnull; repeat for AND semantics
        #[arg(long = "filter")]
        filters: Vec<String>,

        /// Maximum number of rows to return
        #[arg(long)]
        limit: Option<usize>,

        /// Number of rows to skip after filtering
        #[arg(long)]
        offset: Option<usize>,

        /// Drop rows where every cell in the row is empty
        #[arg(long)]
        non_empty: bool,

        /// Output shape for row data
        #[arg(long, value_enum, default_value = "rows")]
        output_shape: OutputShape,

        /// Output format
        #[arg(long, value_enum, default_value = "json")]
        format: OutputFormat,
    },
    /// Read records from a sheet using a resolved header row
    Records {
        /// Excel file path
        file: PathBuf,

        /// Sheet name (exact match)
        #[arg(long, group = "sheet_target")]
        sheet: Option<String>,

        /// Sheet index (0-based)
        #[arg(long, group = "sheet_target")]
        sheet_index: Option<usize>,

        /// Range to read (A1 notation)
        #[arg(long)]
        range: Option<String>,

        /// Header row: auto or 1-based index
        #[arg(long, default_value = "auto")]
        header_row: String,

        /// Select columns by stable column name, comma-separated
        #[arg(long)]
        select: Option<String>,

        /// Filter rows using field:op:value; operators: eq|ne|gt|gte|lt|lte|contains|regex|isnull|notnull; repeat for AND semantics
        #[arg(long = "filter")]
        filters: Vec<String>,

        /// Maximum number of rows to return
        #[arg(long)]
        limit: Option<usize>,

        /// Number of rows to skip after filtering
        #[arg(long)]
        offset: Option<usize>,

        /// Drop rows where every cell in the row is empty
        #[arg(long)]
        non_empty: bool,

        /// Output shape for row data; records by default
        #[arg(long, value_enum, default_value = "records")]
        output_shape: OutputShape,

        /// Output format
        #[arg(long, value_enum, default_value = "json")]
        format: OutputFormat,
    },
}

#[derive(Clone, Debug, Default, clap::ValueEnum)]
pub enum OutputFormat {
    #[default]
    Json,
    Text,
}

impl OutputFormat {
    pub fn as_str(&self) -> &str {
        match self {
            OutputFormat::Json => "json",
            OutputFormat::Text => "text",
        }
    }
}

#[derive(Clone, Copy, Debug, Default, clap::ValueEnum, PartialEq, Eq)]
pub enum OutputShape {
    #[default]
    Rows,
    Records,
    Jsonl,
}

impl OutputShape {
    pub fn as_str(&self) -> &str {
        match self {
            OutputShape::Rows => "rows",
            OutputShape::Records => "records",
            OutputShape::Jsonl => "jsonl",
        }
    }
}

#[derive(Clone, Copy, Debug, Default, clap::ValueEnum, PartialEq, Eq)]
pub enum SeverityThreshold {
    #[default]
    Info,
    Warning,
    Error,
}

impl SeverityThreshold {
    pub fn as_str(&self) -> &'static str {
        match self {
            SeverityThreshold::Info => "info",
            SeverityThreshold::Warning => "warning",
            SeverityThreshold::Error => "error",
        }
    }
}

/// Resolve the sheet target (by name or index) to a sheet index.
pub fn resolve_sheet_target(
    workbook: &crate::excel::Workbook,
    sheet: &Option<String>,
    sheet_index: &Option<usize>,
) -> Result<usize, crate::cli::error::AppError> {
    use crate::cli::error::AppError;

    if let Some(name) = sheet {
        workbook
            .resolve_sheet_by_name(name)
            .map_err(|e| AppError::TargetNotFound {
                message: e.to_string(),
            })
    } else if let Some(index) = sheet_index {
        workbook
            .resolve_sheet_by_index(*index)
            .map_err(|e| AppError::TargetNotFound {
                message: e.to_string(),
            })
    } else {
        Err(AppError::InvalidArgs {
            message: "Either --sheet or --sheet-index must be provided".to_string(),
        })
    }
}
