use crate::app::{AppState, InputMode};
use crate::cli::args::SeverityThreshold;
use crate::cli::check::{run_check_report, CheckFinding};
use crate::utils::{cell_reference, parse_range};

#[derive(Clone, Debug, Default)]
pub(crate) struct FindingsState {
    pub(crate) items: Vec<CheckFinding>,
    pub(crate) selected: usize,
    pub(crate) last_refresh_error: Option<String>,
}

impl FindingsState {
    fn clamp_selected(&mut self) {
        if self.items.is_empty() {
            self.selected = 0;
        } else {
            self.selected = self.selected.min(self.items.len() - 1);
        }
    }
}

impl AppState<'_> {
    pub fn show_findings(&mut self) {
        self.input_mode = InputMode::Findings;
        self.refresh_findings();
    }

    pub fn close_findings(&mut self) {
        self.input_mode = InputMode::Normal;
    }

    pub fn refresh_findings(&mut self) {
        let was_modified = self.workbook.is_modified();
        let result = self.ensure_findings_workbook_ready().and_then(|_| {
            run_check_report(&mut self.workbook, None, None, SeverityThreshold::Info)
        });
        self.workbook.set_modified(was_modified);

        match result {
            Ok(report) => {
                let finding_count = report.findings.len();
                self.findings.items = report.findings;
                self.findings.last_refresh_error = None;
                self.findings.clamp_selected();

                if finding_count == 0 {
                    self.add_notification("No findings in current workbook".to_string());
                } else {
                    self.add_notification(format!("Loaded {finding_count} findings"));
                }
            }
            Err(err) => {
                self.findings.items.clear();
                self.findings.selected = 0;
                self.findings.last_refresh_error = Some(err.to_string());
                self.add_notification(format!("Findings refresh failed: {err}"));
            }
        }
    }

    pub fn select_next_finding(&mut self) {
        if self.findings.selected + 1 < self.findings.items.len() {
            self.findings.selected += 1;
        }
    }

    pub fn select_prev_finding(&mut self) {
        self.findings.selected = self.findings.selected.saturating_sub(1);
    }

    pub fn activate_selected_finding(&mut self) {
        let Some(finding) = self.findings.items.get(self.findings.selected).cloned() else {
            self.add_notification("No finding selected".to_string());
            return;
        };

        let target_index = match self.workbook.resolve_sheet_by_name(&finding.sheet) {
            Ok(index) => index,
            Err(err) => {
                self.add_notification(format!(
                    "Finding sheet '{}' not found: {err}",
                    finding.sheet
                ));
                return;
            }
        };

        if self.workbook.get_current_sheet_index() != target_index {
            if let Err(err) = self.switch_sheet_by_index(target_index) {
                self.add_notification(format!("Failed to switch to finding sheet: {err}"));
                return;
            }
        }

        let Some(target_cell) = finding_target_cell(&finding) else {
            self.add_notification(format!("Jumped to finding on sheet '{}'", finding.sheet));
            return;
        };

        let sheet = self.workbook.get_current_sheet();
        let max_row = sheet.max_rows.max(1);
        let max_col = sheet.max_cols.max(1);
        let clamped = (target_cell.0.min(max_row), target_cell.1.min(max_col));

        self.selected_cell = clamped;
        self.handle_scrolling();

        if clamped == target_cell {
            self.add_notification(format!("Jumped to finding at {}", cell_reference(clamped)));
        } else {
            self.add_notification(format!(
                "Finding target {} was out of range; jumped to {}",
                cell_reference(target_cell),
                cell_reference(clamped)
            ));
        }
    }

    fn ensure_findings_workbook_ready(&mut self) -> Result<(), crate::cli::error::AppError> {
        let sheet_names = self.workbook.get_sheet_names();
        for (index, sheet_name) in sheet_names.iter().enumerate() {
            self.workbook
                .ensure_sheet_loaded(index, sheet_name)
                .map_err(crate::cli::error::anyhow_to_app_error)?;
        }
        Ok(())
    }
}

fn finding_target_cell(finding: &CheckFinding) -> Option<(usize, usize)> {
    match (finding.row, finding.column) {
        (Some(row), Some(col)) => Some((row, col)),
        _ => finding
            .range
            .as_deref()
            .and_then(parse_range)
            .map(|(start, _)| start)
            .or_else(|| finding.row.map(|row| (row, 1)))
            .or_else(|| finding.column.map(|col| (1, col))),
    }
}
