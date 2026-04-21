use serde::Serialize;
use serde_json::{json, Value};
use std::cmp::Ordering;
use std::collections::HashMap;
use std::path::{Path, PathBuf};

use crate::cli::args::SeverityThreshold;
use crate::cli::envelope;
use crate::cli::error::{AppError, EXIT_CHECK_FINDINGS, EXIT_SUCCESS};
use crate::excel::{open_workbook, Cell, Sheet, Workbook};
use crate::utils::{cell_reference, index_to_col_name};

const RULES: [CheckRuleId; 8] = [
    CheckRuleId::BlankHeaders,
    CheckRuleId::DuplicateHeaders,
    CheckRuleId::BlankRows,
    CheckRuleId::BlankColumns,
    CheckRuleId::NullRatio,
    CheckRuleId::DuplicateValues,
    CheckRuleId::TypeDrift,
    CheckRuleId::FormulaPresence,
];

pub fn handle(
    file: PathBuf,
    sheet: Option<String>,
    rules: Option<String>,
    severity_threshold: SeverityThreshold,
) -> Result<(Value, i32), AppError> {
    let format_str = file_format(&file);
    let path_str = file.to_string_lossy().to_string();

    let mut workbook =
        open_workbook(&file, false).map_err(crate::cli::error::anyhow_to_app_error)?;
    let selected_rules = parse_rules(rules.as_deref())?;
    let threshold = Severity::from_threshold(severity_threshold);
    let checked_sheet_indices = resolve_checked_sheets(&workbook, sheet.as_deref())?;

    for index in &checked_sheet_indices {
        let sheet_name = workbook.get_sheet_names()[*index].clone();
        workbook
            .ensure_sheet_loaded(*index, &sheet_name)
            .map_err(crate::cli::error::anyhow_to_app_error)?;
    }

    let sheet_names = workbook.get_sheet_names();
    let mut findings = run_rules(&workbook, &selected_rules, &checked_sheet_indices)?;
    let finding_count_before_threshold = findings.len();
    findings.retain(|finding| finding.severity >= threshold);
    sort_findings(&mut findings, &sheet_names);

    let data = json!({
        "summary": summarize_findings(&findings),
        "stats": build_stats(
            &workbook,
            &checked_sheet_indices,
            &selected_rules,
            severity_threshold,
            finding_count_before_threshold,
        )?,
        "findings": findings,
    });

    let target = if let Some(sheet_name) = sheet {
        let sheet_index = checked_sheet_indices[0];
        envelope::target_sheet(&sheet_name, sheet_index)
    } else {
        envelope::target_workbook()
    };

    let exit_code = exit_code_for_findings(
        data["summary"]["finding_count"]
            .as_u64()
            .unwrap_or_default() as usize,
    );

    Ok((
        envelope::success_envelope(
            "check",
            &path_str,
            &format_str,
            target,
            json!({}),
            data,
            vec![],
        ),
        exit_code,
    ))
}

fn file_format(path: &Path) -> String {
    path.extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_lowercase())
        .unwrap_or_else(|| "unknown".to_string())
}

fn parse_rules(value: Option<&str>) -> Result<Vec<CheckRuleId>, AppError> {
    let Some(value) = value else {
        return Ok(RULES.to_vec());
    };

    let mut requested = Vec::new();
    for raw in value.split(',') {
        let id = raw.trim();
        if id.is_empty() {
            continue;
        }
        let rule = CheckRuleId::parse(id).ok_or_else(|| AppError::InvalidQuery {
            message: format!(
                "Unknown check rule '{}'. Supported rules: {}",
                id,
                RULES
                    .iter()
                    .map(CheckRuleId::as_str)
                    .collect::<Vec<_>>()
                    .join(", ")
            ),
        })?;
        if !requested.contains(&rule) {
            requested.push(rule);
        }
    }

    if requested.is_empty() {
        return Err(AppError::InvalidQuery {
            message: "--rules must include at least one rule id".to_string(),
        });
    }

    Ok(RULES
        .iter()
        .copied()
        .filter(|rule| requested.contains(rule))
        .collect())
}

fn resolve_checked_sheets(
    workbook: &Workbook,
    sheet: Option<&str>,
) -> Result<Vec<usize>, AppError> {
    if let Some(name) = sheet {
        workbook
            .resolve_sheet_by_name(name)
            .map(|index| vec![index])
            .map_err(|e| AppError::TargetNotFound {
                message: e.to_string(),
            })
    } else {
        Ok((0..workbook.get_sheet_names().len()).collect())
    }
}

fn run_rules(
    workbook: &Workbook,
    rules: &[CheckRuleId],
    sheet_indices: &[usize],
) -> Result<Vec<CheckFinding>, AppError> {
    let mut findings = Vec::new();

    for sheet_index in sheet_indices {
        let sheet =
            workbook
                .get_sheet_by_index(*sheet_index)
                .ok_or_else(|| AppError::TargetNotFound {
                    message: format!("Sheet index {} not found", sheet_index),
                })?;
        let used_range = workbook
            .get_used_range(*sheet_index)
            .map_err(crate::cli::error::anyhow_to_app_error)?;
        let (_, header_row) = workbook
            .find_header_candidates(*sheet_index)
            .map_err(crate::cli::error::anyhow_to_app_error)?;

        for rule in rules {
            match rule {
                CheckRuleId::BlankHeaders => {
                    findings.extend(find_blank_headers(sheet, header_row));
                }
                CheckRuleId::DuplicateHeaders => {
                    findings.extend(find_duplicate_headers(sheet, header_row));
                }
                CheckRuleId::BlankRows => {
                    findings.extend(find_blank_rows(sheet, &used_range));
                }
                CheckRuleId::BlankColumns => {
                    findings.extend(find_blank_columns(sheet, &used_range));
                }
                CheckRuleId::NullRatio
                | CheckRuleId::DuplicateValues
                | CheckRuleId::TypeDrift
                | CheckRuleId::FormulaPresence => {}
            }
        }
    }

    Ok(findings)
}

fn find_blank_headers(sheet: &Sheet, header_row: Option<usize>) -> Vec<CheckFinding> {
    let Some(header_row) = header_row else {
        return Vec::new();
    };

    (1..=sheet.max_cols)
        .filter(|col| is_blank_cell(cell_at(sheet, header_row, *col)))
        .map(|col| {
            let column_label = index_to_col_name(col);
            let range = cell_reference((header_row, col));
            CheckFinding {
                rule_id: CheckRuleId::BlankHeaders,
                severity: Severity::Warning,
                sheet: sheet.name.clone(),
                row: Some(header_row),
                column: Some(col),
                range: Some(range.clone()),
                message: format!("Blank header at {range}."),
                details: json!({
                    "header_row": header_row,
                    "column_label": column_label,
                    "reason": "blank_header",
                }),
            }
        })
        .collect()
}

fn find_duplicate_headers(sheet: &Sheet, header_row: Option<usize>) -> Vec<CheckFinding> {
    let Some(header_row) = header_row else {
        return Vec::new();
    };

    let mut counts: HashMap<String, usize> = HashMap::new();
    let mut first_locations: HashMap<String, (usize, String)> = HashMap::new();
    let headers: Vec<_> = (1..=sheet.max_cols)
        .map(|col| {
            let header = header_value(sheet, header_row, col);
            if !header.is_empty() {
                *counts.entry(header.clone()).or_insert(0) += 1;
                first_locations
                    .entry(header.clone())
                    .or_insert_with(|| (col, cell_reference((header_row, col))));
            }
            header
        })
        .collect();

    let mut seen: HashMap<String, usize> = HashMap::new();
    let mut findings = Vec::new();
    for (offset, header) in headers.into_iter().enumerate() {
        if header.is_empty() {
            continue;
        }

        let occurrence = seen.entry(header.clone()).or_insert(0);
        *occurrence += 1;
        if *occurrence == 1 {
            continue;
        }

        let col = offset + 1;
        let range = cell_reference((header_row, col));
        let (first_column, first_range) = first_locations
            .get(&header)
            .cloned()
            .unwrap_or_else(|| (col, range.clone()));
        findings.push(CheckFinding {
            rule_id: CheckRuleId::DuplicateHeaders,
            severity: Severity::Warning,
            sheet: sheet.name.clone(),
            row: Some(header_row),
            column: Some(col),
            range: Some(range.clone()),
            message: format!("Duplicate header '{header}' at {range}."),
            details: json!({
                "header": header,
                "normalized_header": header,
                "first_column": first_column,
                "first_range": first_range,
                "duplicate_count": counts.get(&header).copied().unwrap_or(0),
            }),
        });
    }

    findings
}

fn find_blank_rows(sheet: &Sheet, used_range: &str) -> Vec<CheckFinding> {
    if used_range.is_empty() || sheet.max_rows == 0 || sheet.max_cols == 0 {
        return Vec::new();
    }

    (1..=sheet.max_rows)
        .filter(|row| (1..=sheet.max_cols).all(|col| is_blank_cell(cell_at(sheet, *row, col))))
        .map(|row| {
            let end_col = index_to_col_name(sheet.max_cols);
            let range = format!("A{row}:{end_col}{row}");
            CheckFinding {
                rule_id: CheckRuleId::BlankRows,
                severity: Severity::Warning,
                sheet: sheet.name.clone(),
                row: Some(row),
                column: None,
                range: Some(range),
                message: format!("Blank row {row} in used range {used_range}."),
                details: json!({
                    "used_range": used_range,
                    "max_columns": sheet.max_cols,
                    "reason": "blank_row",
                }),
            }
        })
        .collect()
}

fn find_blank_columns(sheet: &Sheet, used_range: &str) -> Vec<CheckFinding> {
    if used_range.is_empty() || sheet.max_rows == 0 || sheet.max_cols == 0 {
        return Vec::new();
    }

    (1..=sheet.max_cols)
        .filter(|col| (1..=sheet.max_rows).all(|row| is_blank_cell(cell_at(sheet, row, *col))))
        .map(|col| {
            let column_label = index_to_col_name(col);
            let range = format!("{column_label}1:{column_label}{}", sheet.max_rows);
            CheckFinding {
                rule_id: CheckRuleId::BlankColumns,
                severity: Severity::Warning,
                sheet: sheet.name.clone(),
                row: None,
                column: Some(col),
                range: Some(range),
                message: format!("Blank column {column_label} in used range {used_range}."),
                details: json!({
                    "used_range": used_range,
                    "column_label": column_label,
                    "max_rows": sheet.max_rows,
                    "reason": "blank_column",
                }),
            }
        })
        .collect()
}

fn header_value(sheet: &Sheet, row: usize, col: usize) -> String {
    cell_at(sheet, row, col)
        .filter(|cell| !cell_has_formula(cell))
        .map(|cell| cell.value.trim().to_string())
        .unwrap_or_default()
}

fn cell_at(sheet: &Sheet, row: usize, col: usize) -> Option<&Cell> {
    sheet.data.get(row).and_then(|row_data| row_data.get(col))
}

fn is_blank_cell(cell: Option<&Cell>) -> bool {
    cell.map(|cell| !cell_has_formula(cell) && cell.value.trim().is_empty())
        .unwrap_or(true)
}

fn cell_has_formula(cell: &Cell) -> bool {
    cell.is_formula || cell.formula.is_some()
}

fn summarize_findings(findings: &[CheckFinding]) -> Value {
    let error_count = findings
        .iter()
        .filter(|finding| finding.severity == Severity::Error)
        .count();
    let warning_count = findings
        .iter()
        .filter(|finding| finding.severity == Severity::Warning)
        .count();
    let info_count = findings
        .iter()
        .filter(|finding| finding.severity == Severity::Info)
        .count();
    let finding_count = findings.len();

    json!({
        "status": if finding_count == 0 { "pass" } else { "fail" },
        "finding_count": finding_count,
        "error_count": error_count,
        "warning_count": warning_count,
        "info_count": info_count,
    })
}

fn build_stats(
    workbook: &Workbook,
    checked_sheet_indices: &[usize],
    rules: &[CheckRuleId],
    severity_threshold: SeverityThreshold,
    finding_count_before_threshold: usize,
) -> Result<Value, AppError> {
    let checked_sheets: Result<Vec<_>, AppError> = checked_sheet_indices
        .iter()
        .map(|index| {
            let sheet =
                workbook
                    .get_sheet_by_index(*index)
                    .ok_or_else(|| AppError::TargetNotFound {
                        message: format!("Sheet index {} not found", index),
                    })?;
            let used_range = workbook
                .get_used_range(*index)
                .map_err(crate::cli::error::anyhow_to_app_error)?;

            Ok(json!({
                "name": sheet.name,
                "index": index,
                "used_range": used_range,
                "max_rows": sheet.max_rows,
                "max_cols": sheet.max_cols,
            }))
        })
        .collect();

    Ok(json!({
        "sheet_count": workbook.get_sheet_names().len(),
        "checked_sheet_count": checked_sheet_indices.len(),
        "checked_sheets": checked_sheets?,
        "rules_run": rules.iter().map(CheckRuleId::as_str).collect::<Vec<_>>(),
        "severity_threshold": severity_threshold.as_str(),
        "finding_count_before_threshold": finding_count_before_threshold,
    }))
}

fn exit_code_for_findings(finding_count: usize) -> i32 {
    if finding_count == 0 {
        EXIT_SUCCESS
    } else {
        EXIT_CHECK_FINDINGS
    }
}

fn sort_findings(findings: &mut [CheckFinding], sheet_names: &[String]) {
    let sheet_order: HashMap<&str, usize> = sheet_names
        .iter()
        .enumerate()
        .map(|(index, name)| (name.as_str(), index))
        .collect();

    findings.sort_by(|left, right| {
        compare_usize(
            sheet_order.get(left.sheet.as_str()).copied(),
            sheet_order.get(right.sheet.as_str()).copied(),
        )
        .then_with(|| left.rule_id.order().cmp(&right.rule_id.order()))
        .then_with(|| compare_location(left.row, right.row))
        .then_with(|| compare_location(left.column, right.column))
        .then_with(|| left.range.cmp(&right.range))
        .then_with(|| left.message.cmp(&right.message))
        .then_with(|| left.details.to_string().cmp(&right.details.to_string()))
    });
}

fn compare_location(left: Option<usize>, right: Option<usize>) -> Ordering {
    match (left, right) {
        (Some(left), Some(right)) => left.cmp(&right),
        (Some(_), None) => Ordering::Less,
        (None, Some(_)) => Ordering::Greater,
        (None, None) => Ordering::Equal,
    }
}

fn compare_usize(left: Option<usize>, right: Option<usize>) -> Ordering {
    match (left, right) {
        (Some(left), Some(right)) => left.cmp(&right),
        (Some(_), None) => Ordering::Less,
        (None, Some(_)) => Ordering::Greater,
        (None, None) => Ordering::Equal,
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
#[serde(rename_all = "snake_case")]
enum CheckRuleId {
    BlankHeaders,
    DuplicateHeaders,
    BlankRows,
    BlankColumns,
    NullRatio,
    DuplicateValues,
    TypeDrift,
    FormulaPresence,
}

impl CheckRuleId {
    fn parse(value: &str) -> Option<Self> {
        RULES.iter().copied().find(|rule| rule.as_str() == value)
    }

    fn as_str(&self) -> &'static str {
        match self {
            CheckRuleId::BlankHeaders => "blank_headers",
            CheckRuleId::DuplicateHeaders => "duplicate_headers",
            CheckRuleId::BlankRows => "blank_rows",
            CheckRuleId::BlankColumns => "blank_columns",
            CheckRuleId::NullRatio => "null_ratio",
            CheckRuleId::DuplicateValues => "duplicate_values",
            CheckRuleId::TypeDrift => "type_drift",
            CheckRuleId::FormulaPresence => "formula_presence",
        }
    }

    fn order(&self) -> usize {
        RULES
            .iter()
            .position(|rule| rule == self)
            .unwrap_or(usize::MAX)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Serialize)]
#[serde(rename_all = "lowercase")]
enum Severity {
    Info,
    Warning,
    Error,
}

impl Severity {
    fn from_threshold(threshold: SeverityThreshold) -> Self {
        match threshold {
            SeverityThreshold::Info => Severity::Info,
            SeverityThreshold::Warning => Severity::Warning,
            SeverityThreshold::Error => Severity::Error,
        }
    }
}

#[derive(Clone, Debug, Serialize)]
struct CheckFinding {
    rule_id: CheckRuleId,
    severity: Severity,
    sheet: String,
    row: Option<usize>,
    column: Option<usize>,
    range: Option<String>,
    message: String,
    details: Value,
}

#[cfg(test)]
mod tests {
    use serde_json::json;

    use super::*;
    use crate::cli::error::{EXIT_CHECK_FINDINGS, EXIT_SUCCESS};

    #[test]
    fn exit_code_uses_one_for_successful_reports_with_findings() {
        assert_eq!(exit_code_for_findings(0), EXIT_SUCCESS);
        assert_eq!(exit_code_for_findings(2), EXIT_CHECK_FINDINGS);
    }

    #[test]
    fn findings_sort_by_sheet_rule_position_then_location() {
        let mut findings = vec![
            CheckFinding {
                rule_id: CheckRuleId::DuplicateHeaders,
                severity: Severity::Warning,
                sheet: "Orders".to_string(),
                row: Some(3),
                column: Some(2),
                range: Some("B3".to_string()),
                message: "later".to_string(),
                details: json!({"field": "customer"}),
            },
            CheckFinding {
                rule_id: CheckRuleId::BlankHeaders,
                severity: Severity::Warning,
                sheet: "Summary".to_string(),
                row: None,
                column: None,
                range: None,
                message: "workbook-level".to_string(),
                details: json!({}),
            },
            CheckFinding {
                rule_id: CheckRuleId::BlankHeaders,
                severity: Severity::Warning,
                sheet: "Orders".to_string(),
                row: Some(2),
                column: Some(1),
                range: Some("A2".to_string()),
                message: "earlier".to_string(),
                details: json!({}),
            },
        ];

        sort_findings(
            &mut findings,
            &["Summary".to_string(), "Orders".to_string()],
        );

        assert_eq!(findings[0].sheet, "Summary");
        assert_eq!(findings[1].rule_id, CheckRuleId::BlankHeaders);
        assert_eq!(findings[1].row, Some(2));
        assert_eq!(findings[2].rule_id, CheckRuleId::DuplicateHeaders);
    }
}
