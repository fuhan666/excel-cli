use serde_json::Value;

use crate::cli::args::{Cli, Commands};
use crate::cli::error::AppError;

pub fn dispatch(cli: Cli) -> Result<(Value, crate::cli::args::OutputFormat), AppError> {
    match cli.command {
        Commands::Inspect { subcommand } => {
            let format = match &subcommand {
                crate::cli::args::InspectCommands::Workbook { format, .. } => format.clone(),
                crate::cli::args::InspectCommands::Sheet { format, .. } => format.clone(),
                crate::cli::args::InspectCommands::Sample { format, .. } => format.clone(),
                crate::cli::args::InspectCommands::Columns { format, .. } => format.clone(),
                crate::cli::args::InspectCommands::Tables { format, .. } => format.clone(),
            };
            let value = crate::cli::inspect::handle(subcommand)?;
            Ok((value, format))
        }
        Commands::Read { subcommand } => {
            let format = match &subcommand {
                crate::cli::args::ReadCommands::Cell { format, .. } => format.clone(),
                crate::cli::args::ReadCommands::Range { format, .. } => format.clone(),
                crate::cli::args::ReadCommands::Rows { format, .. } => format.clone(),
            };
            let value = crate::cli::read::handle(subcommand)?;
            Ok((value, format))
        }
        Commands::Check { file, rule } => {
            let value = crate::cli::check::handle(file, rule)?;
            // check returns error always in v1.0.0, but we still need a format
            Ok((value, crate::cli::args::OutputFormat::Json))
        }
        Commands::Ui { file } => {
            let workbook = crate::excel::open_workbook(&file, false)
                .map_err(crate::cli::error::anyhow_to_app_error)?;
            let app_state = crate::app::AppState::new(workbook, file)
                .map_err(crate::cli::error::anyhow_to_app_error)?;
            crate::ui::run_app(app_state).map_err(crate::cli::error::anyhow_to_app_error)?;
            Ok((
                crate::cli::envelope::success_envelope(
                    "ui",
                    "",
                    "",
                    crate::cli::envelope::target_workbook(),
                    serde_json::json!({}),
                    serde_json::json!({"status": "interactive"}),
                    vec![],
                ),
                crate::cli::args::OutputFormat::Json,
            ))
        }
    }
}
