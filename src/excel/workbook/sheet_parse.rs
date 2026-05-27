use calamine::{Data, Range};

use crate::excel::{Cell, CellType, DataTypeInfo, FreezePanes, Sheet};

pub(super) fn create_sheet_from_range(
    name: &str,
    range: Range<Data>,
    formula_range: Option<Range<String>>,
) -> Sheet {
    let (height, width) = range.get_size();
    let mut data = vec![vec![Cell::empty(); width + 1]; height + 1];

    for (row_idx, col_idx, cell) in range.used_cells() {
        let (value, cell_type, original_type) = cell_value_parts(cell);
        let is_formula = !value.is_empty() && value.starts_with('=');

        data[row_idx + 1][col_idx + 1] =
            Cell::new_with_type(value, is_formula, cell_type, original_type);
    }

    apply_formula_metadata(&mut data, formula_range);

    Sheet {
        name: name.to_string(),
        data,
        max_rows: height,
        max_cols: width,
        is_loaded: true,
        freeze_panes: FreezePanes::none(),
    }
}

fn cell_value_parts(cell: &Data) -> (String, CellType, Option<DataTypeInfo>) {
    match cell {
        Data::Empty => (String::new(), CellType::Empty, Some(DataTypeInfo::Empty)),
        Data::String(s) => (s.clone(), CellType::Text, Some(DataTypeInfo::String)),
        Data::Float(f) => {
            let value = if *f == (*f as i64) as f64 && f.abs() < 1e10 {
                (*f as i64).to_string()
            } else {
                f.to_string()
            };
            (value, CellType::Number, Some(DataTypeInfo::Float(*f)))
        }
        Data::Int(i) => (i.to_string(), CellType::Number, Some(DataTypeInfo::Int(*i))),
        Data::Bool(b) => (
            if *b {
                "TRUE".to_string()
            } else {
                "FALSE".to_string()
            },
            CellType::Boolean,
            Some(DataTypeInfo::Bool(*b)),
        ),
        Data::Error(e) => {
            let mut value = String::with_capacity(15);
            value.push_str("Error: ");
            value.push_str(&format!("{:?}", e));
            (value, CellType::Text, Some(DataTypeInfo::Error))
        }
        Data::DateTime(dt) => (
            dt.to_string(),
            CellType::Date,
            Some(DataTypeInfo::DateTime(dt.as_f64())),
        ),
        Data::DateTimeIso(s) => {
            let value = s.clone();
            (
                value.clone(),
                CellType::Date,
                Some(DataTypeInfo::DateTimeIso(value)),
            )
        }
        Data::DurationIso(s) => {
            let value = s.clone();
            (
                value.clone(),
                CellType::Text,
                Some(DataTypeInfo::DurationIso(value)),
            )
        }
    }
}

fn apply_formula_metadata(data: &mut [Vec<Cell>], formula_range: Option<Range<String>>) {
    let Some(formulas) = formula_range else {
        return;
    };

    let (start_row, start_col) = formulas.start().unwrap_or((0, 0));
    for (row_idx, col_idx, formula) in formulas.used_cells() {
        if formula.is_empty() {
            continue;
        }

        let normalized = if formula.starts_with('=') {
            formula.to_string()
        } else {
            format!("={formula}")
        };

        let row = start_row as usize + row_idx + 1;
        let col = start_col as usize + col_idx + 1;
        if row < data.len() && col < data[row].len() {
            let cell = &mut data[row][col];
            cell.is_formula = true;
            cell.formula = Some(normalized);
        }
    }
}
