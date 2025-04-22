use anyhow::{Context, Result};
use clap::Parser;
use std::path::PathBuf;

mod app;
mod excel;
mod json_export;
mod ui;

#[derive(Parser)]
#[command(author, version, about, long_about = None)]
struct Cli {
    /// Excel file path
    #[arg(required = true)]
    file_path: PathBuf,

    /// Export all sheets to JSON and output to stdout (for piping)
    #[arg(long, short = 'j')]
    json_export: bool,

    /// Header direction for JSON export: 'h' for horizontal (top rows), 'v' for vertical (left columns)
    #[arg(long, short = 'd', default_value = "h")]
    direction: String,

    /// Number of header rows (for horizontal) or columns (for vertical) in JSON export
    #[arg(long, short = 'r', default_value = "1")]
    header_count: usize,
}

fn main() -> Result<()> {
    let cli = Cli::parse();

    // Open Excel file
    let workbook = excel::open_workbook(&cli.file_path)?;

    // If JSON export flag is set, export to stdout and exit
    if cli.json_export {
        // Parse header direction
        let direction = json_export::HeaderDirection::from_str(&cli.direction)
            .with_context(|| format!("Invalid header direction: {}", cli.direction))?;

        // Generate JSON for all sheets
        let all_sheets =
            json_export::generate_all_sheets_json(&workbook, direction, cli.header_count)?;

        // Serialize to JSON and print to stdout
        let json_string = json_export::serialize_to_json(&all_sheets)?;
        println!("{}", json_string);

        return Ok(());
    }

    // Otherwise, run the interactive UI
    let app_state = app::AppState::new(workbook, cli.file_path)?;
    ui::run_app(app_state)?;

    Ok(())
}
