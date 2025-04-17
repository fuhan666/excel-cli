use anyhow::Result;
use clap::Parser;
use std::path::PathBuf;

mod app;
mod excel;
mod ui;

#[derive(Parser)]
#[command(author, version, about, long_about = None)]
struct Cli {
    /// Excel file path
    #[arg(required = true)]
    file_path: PathBuf,
}

fn main() -> Result<()> {
    let cli = Cli::parse();
    
    // Open Excel file
    let workbook = excel::open_workbook(&cli.file_path)?;
    
    // Create application state
    let app_state = app::AppState::new(workbook, cli.file_path)?;
    
    // Run UI
    ui::run_app(app_state)?;
    
    Ok(())
}
