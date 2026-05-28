use regex::RegexBuilder;
use serde_json::{json, Value};
use std::collections::{HashSet, VecDeque};
use std::path::{Path, PathBuf};
use std::sync::{Arc, Mutex};
use std::thread;

use crate::cli::envelope;
use crate::cli::error::AppError;
use crate::excel::open_workbook;
use crate::utils::{cell_reference, parse_cell_reference};

#[derive(serde::Serialize, serde::Deserialize, Clone, Debug)]
struct GrepMatch {
    file: String,
    sheet: String,
    cell: String,
    content: String,
}

pub fn handle(
    query: String,
    mut paths: Vec<PathBuf>,
    case_insensitive: bool,
    is_regex: bool,
    sheet_filter: Option<String>,
    skip_errors: bool,
) -> Result<(Value, i32), AppError> {
    if paths.is_empty() {
        paths.push(PathBuf::from("."));
    }

    // 1. Compile regex pattern if requested
    let regex_pattern = if is_regex {
        let pattern = RegexBuilder::new(&query)
            .case_insensitive(case_insensitive)
            .build()
            .map_err(|e| AppError::InvalidQuery {
                message: format!("Invalid regex: {}", e),
            })?;
        Some(pattern)
    } else {
        None
    };

    let lower_query = if case_insensitive && !is_regex {
        Some(query.to_lowercase())
    } else {
        None
    };

    // 2. Find all candidate excel files recursively
    let mut files = Vec::new();
    let mut discovery_warnings = Vec::new();
    let mut visited_dirs = HashSet::new();
    for p in &paths {
        find_excel_files(
            p,
            &mut files,
            &mut discovery_warnings,
            &mut visited_dirs,
            true,
        );
    }
    files.sort();
    files.dedup();

    if files.is_empty() {
        let mut warning_values: Vec<Value> =
            discovery_warnings.into_iter().map(Value::String).collect();
        if warning_values.is_empty() {
            warning_values.push(Value::String("No Excel files found to search".to_string()));
        }
        let data = json!({ "matches": [] });
        return Ok((
            envelope::success_envelope(
                "grep",
                ".",
                "xlsx",
                envelope::target_workbook(),
                json!({
                    "query": query,
                    "case_insensitive": case_insensitive,
                    "regex": is_regex,
                }),
                data,
                warning_values,
            ),
            1,
        ));
    }

    // 3. Process files in parallel
    let queue = Arc::new(Mutex::new(VecDeque::from(files)));
    let matches = Arc::new(Mutex::new(Vec::<GrepMatch>::new()));
    let warnings = Arc::new(Mutex::new(Vec::<String>::new()));
    let error = Arc::new(Mutex::new(None::<AppError>));

    let num_workers = thread::available_parallelism()
        .map(|n| n.get())
        .unwrap_or(4);

    let query_arc = Arc::new(query.clone());
    let lower_query_arc = Arc::new(lower_query);
    let regex_arc = Arc::new(regex_pattern);
    let sheet_filter_arc = Arc::new(sheet_filter);

    let mut handles = Vec::new();

    for _ in 0..num_workers {
        let queue = Arc::clone(&queue);
        let matches = Arc::clone(&matches);
        let warnings = Arc::clone(&warnings);
        let error = Arc::clone(&error);
        let query = Arc::clone(&query_arc);
        let lower_query = Arc::clone(&lower_query_arc);
        let regex_pattern = Arc::clone(&regex_arc);
        let sheet_filter = Arc::clone(&sheet_filter_arc);

        let handle = thread::spawn(move || {
            loop {
                // Get next file from queue
                let file_path = {
                    let mut q = queue.lock().unwrap();
                    q.pop_front()
                };

                let Some(path) = file_path else {
                    break; // Queue is empty
                };

                let path_str = path.to_string_lossy().to_string();

                // Open workbook
                let mut workbook = match open_workbook(&path, false) {
                    Ok(wb) => wb,
                    Err(e) => {
                        let warn_msg = format!("Skipped file {}: {}", path_str, e);
                        let mut w = warnings.lock().unwrap();
                        w.push(warn_msg);
                        continue;
                    }
                };

                let sheet_names = workbook.get_sheet_names();

                for name in sheet_names {
                    if let Some(ref filter) = *sheet_filter {
                        if name != *filter {
                            continue; // Skip if sheet filter does not match
                        }
                    }

                    let sheet_idx = match workbook.resolve_sheet_by_name(&name) {
                        Ok(idx) => idx,
                        Err(_) => continue,
                    };

                    match workbook.ensure_sheet_loaded_or_skip(sheet_idx, &name, skip_errors) {
                        Ok(true) => {}
                        Ok(false) => {
                            let warn_msg = format!(
                                "Skipped sheet '{}' in {}: unable to read worksheet",
                                name, path_str
                            );
                            let mut w = warnings.lock().unwrap();
                            w.push(warn_msg);
                            continue;
                        }
                        Err(e) => {
                            let mut err = error.lock().unwrap();
                            if err.is_none() {
                                *err = Some(crate::cli::error::anyhow_to_app_error(e));
                            }
                            return;
                        }
                    }

                    let sheet_obj = match workbook.get_sheet_by_index(sheet_idx) {
                        Some(sheet) => sheet,
                        None => continue,
                    };

                    for row in 1..=sheet_obj.max_rows {
                        if row >= sheet_obj.data.len() {
                            break;
                        }
                        for col in 1..=sheet_obj.max_cols {
                            if col >= sheet_obj.data[row].len() {
                                break;
                            }
                            let cell = &sheet_obj.data[row][col];
                            let cell_ref = cell_reference((row, col));

                            let is_match = if let Some(ref re) = *regex_pattern {
                                re.is_match(&cell.value)
                                    || cell.formula.as_ref().is_some_and(|f| re.is_match(f))
                            } else if let Some(ref lq) = *lower_query {
                                cell.value.to_lowercase().contains(lq)
                                    || cell
                                        .formula
                                        .as_ref()
                                        .is_some_and(|f| f.to_lowercase().contains(lq))
                            } else {
                                cell.value.contains(&*query)
                                    || cell.formula.as_ref().is_some_and(|f| f.contains(&*query))
                            };

                            if is_match {
                                let match_content = if cell.is_formula {
                                    format!(
                                        "{} (Formula: {})",
                                        cell.value,
                                        cell.formula.as_deref().unwrap_or("")
                                    )
                                } else {
                                    cell.value.clone()
                                };

                                let mut m = matches.lock().unwrap();
                                m.push(GrepMatch {
                                    file: path_str.clone(),
                                    sheet: name.clone(),
                                    cell: cell_ref,
                                    content: match_content,
                                });
                            }
                        }
                    }
                }
            }
        });

        handles.push(handle);
    }

    // Wait for all workers to finish
    for handle in handles {
        handle.join().unwrap();
    }

    // Check if any worker encountered an error
    let final_error = Arc::try_unwrap(error).unwrap().into_inner().unwrap();
    if let Some(err) = final_error {
        return Err(err);
    }

    let mut final_matches = Arc::try_unwrap(matches).unwrap().into_inner().unwrap();
    let final_warnings = Arc::try_unwrap(warnings).unwrap().into_inner().unwrap();

    // Sort matches by file, sheet, then cell position (row, col numerically)
    final_matches.sort_by(|a, b| {
        a.file
            .cmp(&b.file)
            .then_with(|| a.sheet.cmp(&b.sheet))
            .then_with(|| {
                let a_pos = parse_cell_reference(&a.cell).unwrap_or((0, 0));
                let b_pos = parse_cell_reference(&b.cell).unwrap_or((0, 0));
                a_pos.cmp(&b_pos)
            })
            .then_with(|| a.content.cmp(&b.content))
    });

    let exit_code = if final_matches.is_empty() { 1 } else { 0 };

    let data = json!({ "matches": final_matches });
    let mut all_warnings: Vec<Value> = discovery_warnings.into_iter().map(Value::String).collect();
    all_warnings.extend(final_warnings.into_iter().map(Value::String));
    let envelope = envelope::success_envelope(
        "grep",
        paths[0].to_string_lossy().as_ref(),
        "xlsx",
        envelope::target_workbook(),
        json!({
            "query": query,
            "case_insensitive": case_insensitive,
            "regex": is_regex,
        }),
        data,
        all_warnings,
    );

    Ok((envelope, exit_code))
}

fn find_excel_files(
    path: &Path,
    files: &mut Vec<PathBuf>,
    warnings: &mut Vec<String>,
    visited_dirs: &mut HashSet<PathBuf>,
    explicit: bool,
) {
    let metadata = match std::fs::symlink_metadata(path) {
        Ok(m) => m,
        Err(_) => {
            if explicit {
                warnings.push(format!("Path not found: {}", path.display()));
            }
            return;
        }
    };

    // Skip directory symlinks to avoid infinite recursion
    if metadata.file_type().is_symlink() && path.is_dir() {
        if explicit {
            warnings.push(format!("Skipping symlink directory: {}", path.display()));
        }
        return;
    }

    if metadata.file_type().is_symlink() {
        // For file symlinks, resolve and check the target
        let resolved = match std::fs::canonicalize(path) {
            Ok(p) => p,
            Err(_) => {
                if explicit {
                    warnings.push(format!("Unable to resolve symlink: {}", path.display()));
                }
                return;
            }
        };
        if resolved.is_file() && is_excel_file(&resolved) {
            files.push(resolved);
        } else if explicit && resolved.is_file() {
            warnings.push(format!("Not an Excel file: {}", path.display()));
        }
        return;
    }

    if metadata.is_file() {
        if is_excel_file(path) {
            files.push(path.to_path_buf());
        } else if explicit {
            warnings.push(format!("Not an Excel file: {}", path.display()));
        }
        return;
    }

    if metadata.is_dir() {
        // Use canonical path to detect directory loops (hardlinks, nested symlinks)
        let canonical = match std::fs::canonicalize(path) {
            Ok(p) => p,
            Err(_) => {
                if explicit {
                    warnings.push(format!("Unable to read directory: {}", path.display()));
                }
                return;
            }
        };

        if !visited_dirs.insert(canonical) {
            return; // Already visited
        }

        let entries = match std::fs::read_dir(path) {
            Ok(e) => e,
            Err(_) => {
                if explicit {
                    warnings.push(format!("Unable to read directory: {}", path.display()));
                }
                return;
            }
        };

        let mut sub_entries: Vec<_> = entries.flatten().collect();
        sub_entries.sort_by_key(|e| e.path());

        for entry in sub_entries {
            find_excel_files(&entry.path(), files, warnings, visited_dirs, false);
        }
    }
}

fn is_excel_file(path: &Path) -> bool {
    let extension = path
        .extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| ext.to_ascii_lowercase());
    // standard temporary files in excel start with ~$ - skip them
    let is_temp = path
        .file_name()
        .and_then(|name| name.to_str())
        .is_some_and(|name| name.starts_with("~$"));

    !is_temp && matches!(extension.as_deref(), Some("xlsx" | "xlsm" | "xls"))
}
