use anyhow::{Context, Result};
use clap::Parser;
use serde_json::{json, Value};
use std::io::IsTerminal;
use std::path::PathBuf;
use std::str::FromStr;

use excel_cli::app;
use excel_cli::excel;
use excel_cli::json_export;
use excel_cli::ui;
use excel_cli::utils::{parse_cell_reference, parse_range};

#[derive(Parser)]
#[command(author, version, about, long_about = None)]
struct Cli {
    /// Excel file path
    #[arg(required = true)]
    file_path: PathBuf,

    /// Export to JSON and output to stdout (for piping)
    #[arg(long, short = 'j')]
    json_export: bool,

    /// Header direction for JSON export: 'h' for horizontal (top rows), 'v' for vertical (left columns)
    #[arg(long, short = 'd', default_value = "h")]
    direction: String,

    /// Number of header rows (for horizontal) or columns (for vertical) in JSON export
    #[arg(long, short = 'r', default_value = "1")]
    header_count: usize,

    /// Enable lazy loading for large Excel files
    #[arg(long, short = 'l')]
    lazy_loading: bool,

    /// List all sheets and exit
    #[arg(long, short = 's')]
    sheets: bool,

    /// Target sheet for inspect or export (by name or 0-based index)
    #[arg(long)]
    sheet: Option<String>,

    /// Peek a range and exit (format: <sheet>!<range>, e.g., Orders!A1:F10)
    #[arg(long, short = 'p')]
    peek: Option<String>,

    /// Read a single cell and exit (format: <sheet>!<cell>, e.g., Summary!B2)
    #[arg(long, short = 'c')]
    cell: Option<String>,

    /// Output format for inspect commands: 'text' or 'json'. Default: 'text'
    #[arg(long, short = 'f', default_value = "text")]
    format: String,
}

fn main() -> Result<()> {
    let cli = Cli::parse();

    let is_headless = cli.sheets
        || cli.sheet.is_some()
        || cli.peek.is_some()
        || cli.cell.is_some()
        || cli.json_export;

    if !std::io::stdout().is_terminal() && !is_headless {
        eprintln!("Excel-cli error: Pipe detected but no headless flag provided.");
        std::process::exit(1);
    }

    // Open Excel file
    let mut workbook = excel::open_workbook(&cli.file_path, cli.lazy_loading)?;

    // Headless inspect/query commands
    if cli.sheets {
        if cli.json_export {
            anyhow::bail!("--sheets and --json-export are mutually exclusive");
        }
        return list_sheets(&workbook, &cli.format);
    }

    if let Some(peek_spec) = cli.peek {
        if cli.cell.is_some() {
            anyhow::bail!("--peek and --cell are mutually exclusive");
        }
        return peek_range(&mut workbook, &peek_spec, &cli.format);
    }

    if let Some(cell_spec) = cli.cell {
        return read_cell(&mut workbook, &cell_spec, &cli.format);
    }

    if let Some(sheet_spec) = cli.sheet {
        if cli.json_export {
            return export_single_sheet(
                &mut workbook,
                &sheet_spec,
                &cli.direction,
                cli.header_count,
            );
        }
        return sheet_info(&mut workbook, &sheet_spec, &cli.format);
    }

    if cli.json_export {
        // Legacy: export all sheets
        let direction = json_export::HeaderDirection::from_str(&cli.direction)
            .map_err(|_| anyhow::anyhow!("Invalid header direction: {}", cli.direction))?;

        let all_sheets =
            json_export::generate_all_sheets_json(&workbook, direction, cli.header_count)?;

        let json_string = json_export::serialize_to_json(&all_sheets)?;
        println!("{json_string}");

        return Ok(());
    }

    // Otherwise, run the interactive UI
    let app_state = app::AppState::new(workbook, cli.file_path)?;
    ui::run_app(app_state)?;

    Ok(())
}

fn list_sheets(workbook: &excel::Workbook, format: &str) -> Result<()> {
    let names = workbook.get_sheet_names();
    match format.to_lowercase().as_str() {
        "json" => {
            let arr: Vec<Value> = names
                .iter()
                .enumerate()
                .map(|(index, name)| json!({"index": index, "name": name}))
                .collect();
            println!("{}", serde_json::to_string_pretty(&arr)?);
        }
        _ => {
            for (index, name) in names.iter().enumerate() {
                println!("{}\t{}", index, name);
            }
        }
    }
    Ok(())
}

fn sheet_info(workbook: &mut excel::Workbook, spec: &str, format: &str) -> Result<()> {
    let index = workbook.resolve_sheet(spec)?;
    let sheet_name = workbook.get_sheet_names()[index].clone();
    workbook.ensure_sheet_loaded(index, &sheet_name)?;
    let sheet = workbook
        .get_sheet_by_index(index)
        .with_context(|| format!("Sheet '{}' not found", spec))?;
    let used_range = workbook.get_used_range(index)?;

    match format.to_lowercase().as_str() {
        "json" => {
            let obj = json!({
                "name": sheet.name,
                "index": index,
                "max_rows": sheet.max_rows,
                "max_cols": sheet.max_cols,
                "used_range": used_range,
            });
            println!("{}", serde_json::to_string_pretty(&obj)?);
        }
        _ => {
            println!("name\t{}", sheet.name);
            println!("index\t{}", index);
            println!("max_rows\t{}", sheet.max_rows);
            println!("max_cols\t{}", sheet.max_cols);
            println!("used_range\t{}", used_range);
        }
    }
    Ok(())
}

fn peek_range(workbook: &mut excel::Workbook, spec: &str, format: &str) -> Result<()> {
    let (sheet_spec, range_str) = spec
        .split_once('!')
        .with_context(|| "Invalid peek format: expected <sheet>!<range>")?;
    let ((start_row, start_col), (end_row, end_col)) =
        parse_range(range_str).with_context(|| "Invalid range format: expected A1:F10")?;

    let index = workbook.resolve_sheet(sheet_spec)?;
    let sheet_name = workbook.get_sheet_names()[index].clone();
    workbook.ensure_sheet_loaded(index, &sheet_name)?;
    let sheet = workbook
        .get_sheet_by_index(index)
        .with_context(|| format!("Sheet '{}' not found", sheet_spec))?;

    // Clamp to actual bounds and normalize inverted ranges
    let max_row = sheet.max_rows.max(1);
    let max_col = sheet.max_cols.max(1);
    let start_row = start_row.min(max_row);
    let end_row = end_row.min(max_row);
    let start_col = start_col.min(max_col);
    let end_col = end_col.min(max_col);
    let (start_row, end_row) = if start_row > end_row {
        (end_row, start_row)
    } else {
        (start_row, end_row)
    };
    let (start_col, end_col) = if start_col > end_col {
        (end_col, start_col)
    } else {
        (start_col, end_col)
    };

    match format.to_lowercase().as_str() {
        "json" => {
            let mut rows = Vec::new();
            for row in start_row..=end_row {
                let mut cols = Vec::new();
                for col in start_col..=end_col {
                    let value = if row < sheet.data.len() && col < sheet.data[row].len() {
                        json_export::process_cell_value(&sheet.data[row][col])
                    } else {
                        Value::Null
                    };
                    cols.push(value);
                }
                rows.push(cols);
            }
            let obj = json!({
                "sheet": sheet_name,
                "range": range_str.to_ascii_uppercase(),
                "rows": rows,
            });
            println!("{}", serde_json::to_string_pretty(&obj)?);
        }
        _ => {
            for row in start_row..=end_row {
                let mut line = String::new();
                for col in start_col..=end_col {
                    if col > start_col {
                        line.push('\t');
                    }
                    let value = if row < sheet.data.len() && col < sheet.data[row].len() {
                        &sheet.data[row][col].value
                    } else {
                        ""
                    };
                    line.push_str(value);
                }
                println!("{}", line);
            }
        }
    }
    Ok(())
}

fn read_cell(workbook: &mut excel::Workbook, spec: &str, format: &str) -> Result<()> {
    let (sheet_spec, cell_str) = spec
        .split_once('!')
        .with_context(|| "Invalid cell format: expected <sheet>!<cell>")?;
    let (row, col) =
        parse_cell_reference(cell_str).with_context(|| "Invalid cell reference: expected A1")?;

    let index = workbook.resolve_sheet(sheet_spec)?;
    let sheet_name = workbook.get_sheet_names()[index].clone();
    workbook.ensure_sheet_loaded(index, &sheet_name)?;
    let sheet = workbook
        .get_sheet_by_index(index)
        .with_context(|| format!("Sheet '{}' not found", sheet_spec))?;

    let in_bounds = row < sheet.data.len() && col < sheet.data[row].len();
    let (value, cell_type) = if in_bounds {
        let cell = &sheet.data[row][col];
        (
            cell.value.clone(),
            match cell.cell_type {
                excel::CellType::Text => "text",
                excel::CellType::Number => "number",
                excel::CellType::Date => "date",
                excel::CellType::Boolean => "boolean",
                excel::CellType::Empty => "empty",
            },
        )
    } else {
        (String::new(), "empty")
    };

    match format.to_lowercase().as_str() {
        "json" => {
            let obj = json!({
                "sheet": sheet_name,
                "cell": cell_str.to_ascii_uppercase(),
                "value": value,
                "type": cell_type,
            });
            println!("{}", serde_json::to_string_pretty(&obj)?);
        }
        _ => {
            println!("{}", value);
        }
    }
    Ok(())
}

fn export_single_sheet(
    workbook: &mut excel::Workbook,
    spec: &str,
    direction_str: &str,
    header_count: usize,
) -> Result<()> {
    let direction = json_export::HeaderDirection::from_str(direction_str)
        .map_err(|_| anyhow::anyhow!("Invalid header direction: {}", direction_str))?;

    let index = workbook.resolve_sheet(spec)?;
    let sheet_name = workbook.get_sheet_names()[index].clone();
    workbook.ensure_sheet_loaded(index, &sheet_name)?;

    let sheet = workbook
        .get_sheet_by_index(index)
        .with_context(|| format!("Sheet '{}' not found", spec))?;
    let sheet_data = json_export::process_sheet_for_json(sheet, direction, header_count)?;
    let json_string = json_export::serialize_to_json(&sheet_data)?;
    println!("{json_string}");

    Ok(())
}
