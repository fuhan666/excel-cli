use serde::Serialize;
use serde_json::{json, Value};
use std::cmp::Ordering;
use std::collections::{BTreeMap, HashMap};
use std::path::{Path, PathBuf};

use crate::cli::args::SeverityThreshold;
use crate::cli::envelope;
use crate::cli::error::{AppError, EXIT_CHECK_FINDINGS, EXIT_SUCCESS};
use crate::excel::{open_workbook, Cell, CellType, Sheet, Workbook};
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
    let report = run_check_report(
        &mut workbook,
        sheet.as_deref(),
        rules.as_deref(),
        severity_threshold,
    )?;

    let data = json!({
        "summary": report.summary,
        "stats": report.stats,
        "findings": report.findings,
    });

    let target = if let Some(sheet_name) = sheet {
        let sheet_index =
            workbook
                .resolve_sheet_by_name(&sheet_name)
                .map_err(|e| AppError::TargetNotFound {
                    message: e.to_string(),
                })?;
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

pub(crate) fn run_check_report(
    workbook: &mut Workbook,
    sheet: Option<&str>,
    rules: Option<&str>,
    severity_threshold: SeverityThreshold,
) -> Result<CheckReport, AppError> {
    let selected_rules = parse_rules(rules)?;
    let threshold = Severity::from_threshold(severity_threshold);
    let checked_sheet_indices = resolve_checked_sheets(workbook, sheet)?;

    for index in &checked_sheet_indices {
        let sheet_name = workbook.get_sheet_names()[*index].clone();
        workbook
            .ensure_sheet_loaded(*index, &sheet_name)
            .map_err(crate::cli::error::anyhow_to_app_error)?;
    }

    let sheet_names = workbook.get_sheet_names();
    let mut findings = run_rules(workbook, &selected_rules, &checked_sheet_indices)?;
    let finding_count_before_threshold = findings.len();
    findings.retain(|finding| finding.severity >= threshold);
    sort_findings(&mut findings, &sheet_names);

    Ok(CheckReport {
        summary: summarize_findings(&findings),
        stats: build_stats(
            workbook,
            &checked_sheet_indices,
            &selected_rules,
            severity_threshold,
            finding_count_before_threshold,
        )?,
        findings,
    })
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
        let context = SheetCheckContext::new(workbook, *sheet_index)?;
        for rule in rules {
            match rule {
                CheckRuleId::BlankHeaders => findings.extend(find_blank_headers(&context)),
                CheckRuleId::DuplicateHeaders => findings.extend(find_duplicate_headers(&context)),
                CheckRuleId::BlankRows => findings.extend(find_blank_rows(&context)),
                CheckRuleId::BlankColumns => findings.extend(find_blank_columns(&context)),
                CheckRuleId::NullRatio => findings.extend(check_null_ratio(&context)),
                CheckRuleId::DuplicateValues => findings.extend(check_duplicate_values(&context)),
                CheckRuleId::TypeDrift => findings.extend(check_type_drift(&context)),
                CheckRuleId::FormulaPresence => findings.extend(check_formula_presence(&context)),
            }
        }
    }

    Ok(findings)
}

struct SheetCheckContext<'a> {
    sheet: &'a Sheet,
    header_row: Option<usize>,
    used_range: String,
    data_start_row: usize,
    data_row_count: usize,
}

impl<'a> SheetCheckContext<'a> {
    fn new(workbook: &'a Workbook, sheet_index: usize) -> Result<Self, AppError> {
        let sheet =
            workbook
                .get_sheet_by_index(sheet_index)
                .ok_or_else(|| AppError::TargetNotFound {
                    message: format!("Sheet index {} not found", sheet_index),
                })?;
        let used_range = workbook
            .get_used_range(sheet_index)
            .map_err(crate::cli::error::anyhow_to_app_error)?;
        let (_, header_row) = workbook
            .find_header_candidates(sheet_index)
            .map_err(crate::cli::error::anyhow_to_app_error)?;
        let data_start_row = header_row.map_or(1, |row| row.saturating_add(1));
        let data_row_count = if sheet.max_rows >= data_start_row {
            sheet.max_rows - data_start_row + 1
        } else {
            0
        };

        Ok(Self {
            sheet,
            header_row,
            used_range,
            data_start_row,
            data_row_count,
        })
    }

    fn column_name(&self, col: usize) -> String {
        self.header_row
            .and_then(|row| cell_at(self.sheet, row, col))
            .map(|cell| cell.value.trim())
            .filter(|value| !value.is_empty())
            .map(ToOwned::to_owned)
            .unwrap_or_else(|| format!("col_{}", index_to_col_name(col)))
    }

    fn data_column_range(&self, col: usize) -> Option<String> {
        if self.data_row_count == 0 {
            None
        } else {
            Some(format!(
                "{}{}:{}{}",
                index_to_col_name(col),
                self.data_start_row,
                index_to_col_name(col),
                self.sheet.max_rows
            ))
        }
    }
}

fn find_blank_headers(context: &SheetCheckContext<'_>) -> Vec<CheckFinding> {
    let Some(header_row) = context.header_row else {
        return Vec::new();
    };

    (1..=context.sheet.max_cols)
        .filter(|col| is_blank_cell(cell_at(context.sheet, header_row, *col)))
        .map(|col| {
            let column_label = index_to_col_name(col);
            let range = cell_reference((header_row, col));
            CheckFinding {
                rule_id: CheckRuleId::BlankHeaders,
                severity: Severity::Warning,
                sheet: context.sheet.name.clone(),
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

fn find_duplicate_headers(context: &SheetCheckContext<'_>) -> Vec<CheckFinding> {
    let Some(header_row) = context.header_row else {
        return Vec::new();
    };

    let mut counts: HashMap<String, usize> = HashMap::new();
    let mut first_locations: HashMap<String, (usize, String)> = HashMap::new();
    let headers: Vec<_> = (1..=context.sheet.max_cols)
        .map(|col| {
            let header = header_value(context.sheet, header_row, col);
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
            sheet: context.sheet.name.clone(),
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

fn find_blank_rows(context: &SheetCheckContext<'_>) -> Vec<CheckFinding> {
    if context.used_range.is_empty() || context.sheet.max_rows == 0 || context.sheet.max_cols == 0 {
        return Vec::new();
    }

    (1..=context.sheet.max_rows)
        .filter(|row| {
            (1..=context.sheet.max_cols).all(|col| is_blank_cell(cell_at(context.sheet, *row, col)))
        })
        .map(|row| {
            let end_col = index_to_col_name(context.sheet.max_cols);
            let range = format!("A{row}:{end_col}{row}");
            CheckFinding {
                rule_id: CheckRuleId::BlankRows,
                severity: Severity::Warning,
                sheet: context.sheet.name.clone(),
                row: Some(row),
                column: None,
                range: Some(range),
                message: format!("Blank row {row} in used range {}.", context.used_range),
                details: json!({
                    "used_range": context.used_range,
                    "max_columns": context.sheet.max_cols,
                    "reason": "blank_row",
                }),
            }
        })
        .collect()
}

fn find_blank_columns(context: &SheetCheckContext<'_>) -> Vec<CheckFinding> {
    if context.used_range.is_empty() || context.sheet.max_rows == 0 || context.sheet.max_cols == 0 {
        return Vec::new();
    }

    (1..=context.sheet.max_cols)
        .filter(|col| {
            (1..=context.sheet.max_rows).all(|row| is_blank_cell(cell_at(context.sheet, row, *col)))
        })
        .map(|col| {
            let column_label = index_to_col_name(col);
            let range = format!("{column_label}1:{column_label}{}", context.sheet.max_rows);
            CheckFinding {
                rule_id: CheckRuleId::BlankColumns,
                severity: Severity::Warning,
                sheet: context.sheet.name.clone(),
                row: None,
                column: Some(col),
                range: Some(range),
                message: format!(
                    "Blank column {column_label} in used range {}.",
                    context.used_range
                ),
                details: json!({
                    "used_range": context.used_range,
                    "column_label": column_label,
                    "max_rows": context.sheet.max_rows,
                    "reason": "blank_column",
                }),
            }
        })
        .collect()
}

fn check_null_ratio(context: &SheetCheckContext<'_>) -> Vec<CheckFinding> {
    if context.data_row_count == 0 {
        return Vec::new();
    }

    let mut findings = Vec::new();
    for col in 1..=context.sheet.max_cols {
        let null_rows: Vec<usize> = (context.data_start_row..=context.sheet.max_rows)
            .filter(|row| !cell_is_present(cell_at(context.sheet, *row, col)))
            .collect();

        if null_rows.is_empty() {
            continue;
        }

        let null_count = null_rows.len();
        let null_ratio = rounded_ratio(null_count, context.data_row_count);
        let severity = if null_count == context.data_row_count {
            Severity::Error
        } else if null_ratio >= 0.5 {
            Severity::Warning
        } else {
            Severity::Info
        };
        let column_name = context.column_name(col);
        let first_null_row = null_rows[0];
        let first_null_cell = cell_reference((first_null_row, col));

        findings.push(CheckFinding {
            rule_id: CheckRuleId::NullRatio,
            severity,
            sheet: context.sheet.name.clone(),
            row: Some(first_null_row),
            column: Some(col),
            range: context.data_column_range(col),
            message: format!(
                "Column '{}' has blank values in {} of {} data rows.",
                column_name, null_count, context.data_row_count
            ),
            details: json!({
                "column_name": column_name,
                "data_row_count": context.data_row_count,
                "first_null_cell": first_null_cell,
                "null_count": null_count,
                "null_ratio": null_ratio,
                "severity_threshold": {
                    "info": "> 0 and < 0.5",
                    "warning": ">= 0.5 and < 1.0",
                    "error": "1.0"
                }
            }),
        });
    }

    findings
}

fn check_duplicate_values(context: &SheetCheckContext<'_>) -> Vec<CheckFinding> {
    let Some((candidate_col, selection)) = default_duplicate_candidate(context) else {
        return Vec::new();
    };

    let mut values: BTreeMap<String, Vec<usize>> = BTreeMap::new();
    for row in context.data_start_row..=context.sheet.max_rows {
        if let Some(cell) = cell_at(context.sheet, row, candidate_col) {
            let value = cell.value.trim();
            if !value.is_empty() {
                values.entry(value.to_string()).or_default().push(row);
            }
        }
    }

    let column_name = context.column_name(candidate_col);
    values
        .into_iter()
        .filter(|(_, rows)| rows.len() > 1)
        .map(|(duplicate_value, rows)| {
            let cells: Vec<String> = rows
                .iter()
                .map(|row| cell_reference((*row, candidate_col)))
                .collect();

            CheckFinding {
                rule_id: CheckRuleId::DuplicateValues,
                severity: Severity::Warning,
                sheet: context.sheet.name.clone(),
                row: rows.first().copied(),
                column: Some(candidate_col),
                range: context.data_column_range(candidate_col),
                message: format!(
                    "Column '{}' has duplicate value '{}' in {} rows.",
                    column_name,
                    duplicate_value,
                    rows.len()
                ),
                details: json!({
                    "candidate_column": {
                        "column": candidate_col,
                        "column_name": column_name,
                        "selection": selection
                    },
                    "duplicate_value": duplicate_value,
                    "occurrence_count": rows.len(),
                    "rows": rows,
                    "cells": cells
                }),
            }
        })
        .collect()
}

fn default_duplicate_candidate(context: &SheetCheckContext<'_>) -> Option<(usize, &'static str)> {
    if context.data_row_count == 0 {
        return None;
    }

    if let Some(header_row) = context.header_row {
        for col in 1..=context.sheet.max_cols {
            let has_header = cell_at(context.sheet, header_row, col)
                .map(|cell| !cell.value.trim().is_empty())
                .unwrap_or(false);
            if has_header && data_column_has_value(context, col) {
                return Some((col, "first non-empty header data column"));
            }
        }
    }

    (1..=context.sheet.max_cols)
        .find(|col| data_column_has_value(context, *col))
        .map(|col| (col, "first data column with values"))
}

fn check_type_drift(context: &SheetCheckContext<'_>) -> Vec<CheckFinding> {
    if context.data_row_count == 0 {
        return Vec::new();
    }

    let mut findings = Vec::new();
    for col in 1..=context.sheet.max_cols {
        let mut type_counts: BTreeMap<&'static str, usize> = BTreeMap::new();
        let mut cells_by_type: BTreeMap<&'static str, Vec<String>> = BTreeMap::new();

        for row in context.data_start_row..=context.sheet.max_rows {
            let Some(cell) = cell_at(context.sheet, row, col) else {
                continue;
            };
            let Some(kind) = cell_kind(cell) else {
                continue;
            };

            *type_counts.entry(kind).or_default() += 1;
            cells_by_type
                .entry(kind)
                .or_default()
                .push(cell_reference((row, col)));
        }

        if type_counts.len() < 2 {
            continue;
        }

        let dominant_type = dominant_type(&type_counts);
        let Some((drift_type, drift_count)) = first_drift_type(&type_counts, dominant_type) else {
            continue;
        };
        let Some(first_drift_cell) = cells_by_type
            .get(drift_type)
            .and_then(|cells| cells.first())
            .cloned()
        else {
            continue;
        };
        let Some((first_drift_row, _)) = parse_cell_for_row(&first_drift_cell) else {
            continue;
        };
        let column_name = context.column_name(col);
        let sample_drift_cells: Vec<String> = cells_by_type
            .get(drift_type)
            .into_iter()
            .flat_map(|cells| cells.iter().take(5).cloned())
            .collect();

        findings.push(CheckFinding {
            rule_id: CheckRuleId::TypeDrift,
            severity: Severity::Warning,
            sheet: context.sheet.name.clone(),
            row: Some(first_drift_row),
            column: Some(col),
            range: context.data_column_range(col),
            message: format!(
                "Column '{}' mixes {} values with dominant {} values.",
                column_name, drift_type, dominant_type
            ),
            details: json!({
                "column_name": column_name,
                "dominant_type": dominant_type,
                "drift_type": drift_type,
                "drift_count": drift_count,
                "type_counts": type_counts,
                "sample_drift_cells": sample_drift_cells
            }),
        });
    }

    findings
}

fn check_formula_presence(context: &SheetCheckContext<'_>) -> Vec<CheckFinding> {
    if context.data_row_count == 0 {
        return Vec::new();
    }

    let mut formulas = Vec::new();
    let mut min_row = usize::MAX;
    let mut min_col = usize::MAX;
    let mut max_row = 0;
    let mut max_col = 0;

    for row in context.data_start_row..=context.sheet.max_rows {
        for col in 1..=context.sheet.max_cols {
            let Some(cell) = cell_at(context.sheet, row, col) else {
                continue;
            };
            if !cell_has_formula(cell) {
                continue;
            }

            min_row = min_row.min(row);
            min_col = min_col.min(col);
            max_row = max_row.max(row);
            max_col = max_col.max(col);
            formulas.push(json!({
                "cell": cell_reference((row, col)),
                "formula": cell.formula.clone().unwrap_or_else(|| cell.value.clone())
            }));
        }
    }

    if formulas.is_empty() {
        return Vec::new();
    }

    let formula_count = formulas.len();
    let formula_ratio = rounded_ratio(formula_count, context.data_row_count);
    formulas.truncate(5);

    vec![CheckFinding {
        rule_id: CheckRuleId::FormulaPresence,
        severity: Severity::Info,
        sheet: context.sheet.name.clone(),
        row: Some(min_row),
        column: Some(min_col),
        range: Some(format!(
            "{}{}:{}{}",
            index_to_col_name(min_col),
            min_row,
            index_to_col_name(max_col),
            max_row
        )),
        message: format!(
            "Sheet '{}' contains {} formula cells.",
            context.sheet.name, formula_count
        ),
        details: json!({
            "data_row_count": context.data_row_count,
            "formula_count": formula_count,
            "formula_ratio": formula_ratio,
            "sample_formula_cells": formulas
        }),
    }]
}

fn cell_at(sheet: &Sheet, row: usize, col: usize) -> Option<&Cell> {
    sheet.data.get(row).and_then(|row_data| row_data.get(col))
}

fn header_value(sheet: &Sheet, row: usize, col: usize) -> String {
    cell_at(sheet, row, col)
        .filter(|cell| !cell_has_formula(cell))
        .map(|cell| cell.value.trim().to_string())
        .unwrap_or_default()
}

fn is_blank_cell(cell: Option<&Cell>) -> bool {
    cell.map(|cell| !cell_has_formula(cell) && cell.value.trim().is_empty())
        .unwrap_or(true)
}

fn cell_has_formula(cell: &Cell) -> bool {
    cell.is_formula || cell.formula.is_some()
}

fn cell_is_present(cell: Option<&Cell>) -> bool {
    cell.map(|cell| !cell.value.trim().is_empty() || cell_has_formula(cell))
        .unwrap_or(false)
}

fn data_column_has_value(context: &SheetCheckContext<'_>, col: usize) -> bool {
    (context.data_start_row..=context.sheet.max_rows)
        .any(|row| cell_is_present(cell_at(context.sheet, row, col)))
}

fn cell_kind(cell: &Cell) -> Option<&'static str> {
    if !cell_is_present(Some(cell)) {
        return None;
    }

    match cell.cell_type {
        CellType::Text => Some("string"),
        CellType::Number => Some("number"),
        CellType::Date => Some("date"),
        CellType::Boolean => Some("boolean"),
        CellType::Empty => None,
    }
}

fn dominant_type(type_counts: &BTreeMap<&'static str, usize>) -> &'static str {
    type_counts
        .iter()
        .max_by(|(left_type, left_count), (right_type, right_count)| {
            left_count
                .cmp(right_count)
                .then_with(|| right_type.cmp(left_type))
        })
        .map(|(kind, _)| *kind)
        .unwrap_or("string")
}

fn first_drift_type(
    type_counts: &BTreeMap<&'static str, usize>,
    dominant_type: &'static str,
) -> Option<(&'static str, usize)> {
    type_counts
        .iter()
        .filter(|(kind, _)| **kind != dominant_type)
        .min_by(|(left_type, left_count), (right_type, right_count)| {
            left_count
                .cmp(right_count)
                .then_with(|| left_type.cmp(right_type))
        })
        .map(|(kind, count)| (*kind, *count))
}

fn parse_cell_for_row(cell: &str) -> Option<(usize, usize)> {
    crate::utils::parse_cell_reference(cell)
}

fn rounded_ratio(numerator: usize, denominator: usize) -> f64 {
    if denominator == 0 {
        0.0
    } else {
        ((numerator as f64 / denominator as f64) * 10_000.0).round() / 10_000.0
    }
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

#[derive(Clone, Debug)]
pub(crate) struct CheckReport {
    pub(crate) summary: Value,
    pub(crate) stats: Value,
    pub(crate) findings: Vec<CheckFinding>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
#[serde(rename_all = "snake_case")]
pub(crate) enum CheckRuleId {
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

    pub(crate) fn as_str(&self) -> &'static str {
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
pub(crate) enum Severity {
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
pub(crate) struct CheckFinding {
    pub(crate) rule_id: CheckRuleId,
    pub(crate) severity: Severity,
    pub(crate) sheet: String,
    pub(crate) row: Option<usize>,
    pub(crate) column: Option<usize>,
    pub(crate) range: Option<String>,
    pub(crate) message: String,
    pub(crate) details: Value,
}

#[cfg(test)]
mod tests {
    use serde_json::json;

    use super::*;
    use crate::cli::error::{EXIT_CHECK_FINDINGS, EXIT_SUCCESS};
    use crate::excel::{Cell, Sheet, Workbook};

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

    fn sheet_with_values(name: &str, values: &[&[&str]]) -> Sheet {
        let max_rows = values.len();
        let max_cols = values.iter().map(|row| row.len()).max().unwrap_or(0);
        let mut data = vec![vec![Cell::empty(); max_cols + 1]; max_rows + 1];

        for (row_idx, row) in values.iter().enumerate() {
            for (col_idx, value) in row.iter().enumerate() {
                data[row_idx + 1][col_idx + 1] = Cell::new((*value).to_string(), false);
            }
        }

        Sheet {
            name: name.to_string(),
            data,
            max_rows,
            max_cols,
            is_loaded: true,
        }
    }

    #[test]
    fn run_check_report_reuses_rule_pipeline_for_structured_findings() {
        let mut workbook = Workbook::from_sheets_for_test(vec![sheet_with_values(
            "Data",
            &[&["Name", "Name"], &["Ada", ""], &["", ""]],
        )]);

        let report = run_check_report(&mut workbook, None, None, SeverityThreshold::Info).unwrap();

        assert_eq!(report.summary["status"], "fail");
        assert_eq!(report.stats["checked_sheet_count"], 1);
        assert!(!report.findings.is_empty());
        assert!(report
            .findings
            .iter()
            .any(|finding| finding.rule_id == CheckRuleId::DuplicateHeaders));
    }
}
