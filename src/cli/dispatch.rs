use serde_json::Value;

use crate::cli::args::{Cli, Commands, OutputFormat};
use crate::cli::error::{AppError, EXIT_SUCCESS};

pub fn dispatch(cli: Cli) -> Result<(Value, OutputFormat, i32), AppError> {
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
            Ok((value, format, EXIT_SUCCESS))
        }
        Commands::Read { subcommand } => {
            let format = match &subcommand {
                crate::cli::args::ReadCommands::Cell { format, .. } => format.clone(),
                crate::cli::args::ReadCommands::Range { format, .. } => format.clone(),
                crate::cli::args::ReadCommands::Rows { format, .. } => format.clone(),
                crate::cli::args::ReadCommands::Records { format, .. } => format.clone(),
            };
            let value = crate::cli::read::handle(subcommand)?;
            Ok((value, format, EXIT_SUCCESS))
        }
        Commands::Check {
            file,
            sheet,
            rules,
            severity_threshold,
        } => {
            let (value, exit_code) =
                crate::cli::check::handle(file, sheet, rules, severity_threshold)?;
            Ok((value, OutputFormat::Json, exit_code))
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
                OutputFormat::Json,
                EXIT_SUCCESS,
            ))
        }
        Commands::Grep {
            query,
            paths,
            case_insensitive,
            regex,
            sheet,
            format,
            skip_errors,
        } => {
            if matches!(format, OutputFormat::Text) {
                return Err(AppError::InvalidArgs {
                    message: "grep no longer supports --format text; use --format markdown for human-readable output or --format json for machine parsing".to_string(),
                });
            }

            let (value, exit_code) = crate::cli::grep::handle(
                query,
                paths,
                case_insensitive,
                regex,
                sheet,
                skip_errors,
            )?;
            Ok((value, format, exit_code))
        }
    }
}
