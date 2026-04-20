use clap::{error::ErrorKind, Parser};

use excel_cli::cli::args::Cli;
use excel_cli::cli::dispatch;
use excel_cli::cli::error::AppError;
use excel_cli::cli::output;

fn main() {
    let cli = match Cli::try_parse() {
        Ok(c) => c,
        Err(e) => match e.kind() {
            ErrorKind::DisplayHelp | ErrorKind::DisplayVersion => {
                print!("{e}");
                std::process::exit(0);
            }
            _ => {
                let err = AppError::InvalidArgs {
                    message: e.to_string(),
                };
                let envelope = err.to_envelope("", "", "unknown");
                output::write_error(&envelope);
                std::process::exit(err.exit_code());
            }
        },
    };

    let result = dispatch::dispatch(cli);

    match result {
        Ok((value, format)) => {
            if let Err(e) = output::write_success(&value, &format) {
                let envelope = e.to_envelope("", "", "unknown");
                output::write_error(&envelope);
                std::process::exit(e.exit_code());
            }
            std::process::exit(0);
        }
        Err(e) => {
            // Try to extract file/command info for the error envelope
            let envelope = e.to_envelope("", "", "unknown");
            output::write_error(&envelope);
            std::process::exit(e.exit_code());
        }
    }
}
