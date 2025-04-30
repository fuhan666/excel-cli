use anyhow::Result;
use clap::Parser;
use std::io::IsTerminal;
use std::path::PathBuf;
use std::str::FromStr;

use excel_cli::app;
use excel_cli::excel;
use excel_cli::json_export;
use excel_cli::ui;

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

    /// Enable lazy loading for large Excel files
    #[arg(long, short = 'l')]
    lazy_loading: bool,
}

fn main() -> Result<()> {
    let cli = Cli::parse();

    if !std::io::stdout().is_terminal() && !cli.json_export {
        eprintln!("Excel-cli error: Pipe detected but -j or --json-export flag not provided.");
        std::process::exit(1);
    }

    // Open Excel file
    let workbook = excel::open_workbook(&cli.file_path, cli.lazy_loading)?;

    // If JSON export flag is set, export to stdout and exit
    if cli.json_export {
        // Parse header direction
        let direction = match json_export::HeaderDirection::from_str(&cli.direction) {
            Ok(dir) => dir,
            Err(_) => anyhow::bail!("Invalid header direction: {}", cli.direction),
        };

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
