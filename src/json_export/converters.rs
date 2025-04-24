use chrono::{Duration, NaiveDate, NaiveDateTime};
use serde_json::{Value, json};

use crate::excel::{Cell, CellType, DataTypeInfo};

// Convert Excel date number to ISO date string
pub fn excel_date_to_iso_string(excel_date: f64) -> String {
    let days = if excel_date > 59.0 {
        excel_date - 1.0
    } else {
        excel_date
    };

    let base_date = NaiveDate::from_ymd_opt(1900, 1, 1).unwrap();
    let whole_days = days.trunc() as i64;
    let fractional_day = days.fract();

    let date = base_date + Duration::days(whole_days - 1); // Subtract 1 because Excel day 1 is 1900-01-01

    if fractional_day > 0.0 {
        let seconds_in_day = 24.0 * 60.0 * 60.0;
        let seconds = (fractional_day * seconds_in_day).round() as u32;

        let hours = seconds / 3600;
        let minutes = (seconds % 3600) / 60;
        let secs = seconds % 60;

        let datetime = NaiveDateTime::new(
            date,
            chrono::NaiveTime::from_hms_opt(hours, minutes, secs).unwrap(),
        );

        datetime.format("%Y-%m-%dT%H:%M:%S").to_string()
    } else {
        date.format("%Y-%m-%d").to_string()
    }
}

// Process cell value based on its type
pub fn process_cell_value(cell: &Cell) -> Value {
    if cell.value.is_empty() {
        return Value::Null;
    }

    if let Some(original_type) = &cell.original_type {
        match original_type {
            DataTypeInfo::Float(f) => {
                if f.fract() == 0.0 {
                    json!(f.trunc() as i64)
                } else {
                    json!(f)
                }
            }
            DataTypeInfo::Int(i) => json!(i),
            DataTypeInfo::DateTime(dt) => {
                if *dt >= 0.0 {
                    json!(excel_date_to_iso_string(*dt))
                } else {
                    json!(cell.value)
                }
            }
            DataTypeInfo::DateTimeIso(s) => json!(s),
            DataTypeInfo::Bool(b) => json!(b),
            DataTypeInfo::Empty => Value::Null,
            _ => json!(cell.value),
        }
    } else {
        match cell.cell_type {
            CellType::Number => {
                if let Ok(num) = cell.value.parse::<f64>() {
                    if num.fract() == 0.0 {
                        json!(num.trunc() as i64)
                    } else {
                        json!(num)
                    }
                } else {
                    json!(cell.value)
                }
            }
            CellType::Boolean => {
                if cell.value.to_lowercase() == "true" {
                    json!(true)
                } else if cell.value.to_lowercase() == "false" {
                    json!(false)
                } else {
                    json!(cell.value)
                }
            }
            CellType::Date => {
                if let Ok(excel_date) = cell.value.parse::<f64>() {
                    if excel_date >= 0.0 {
                        json!(excel_date_to_iso_string(excel_date))
                    } else {
                        json!(cell.value)
                    }
                } else {
                    json!(cell.value)
                }
            }
            CellType::Empty => Value::Null,
            _ => json!(cell.value), // Text, etc.
        }
    }
}
