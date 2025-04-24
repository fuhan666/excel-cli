#[derive(Clone)]
pub struct Cell {
    pub value: String,
    pub is_formula: bool,
    pub cell_type: CellType,
    pub original_type: Option<DataTypeInfo>,
}

#[derive(Clone, PartialEq)]
pub enum CellType {
    Text,
    Number,
    Date,
    Boolean,
    Empty,
}

#[derive(Clone, PartialEq)]
pub enum DataTypeInfo {
    Empty,
    String,
    Float(f64),
    Int(i64),
    Bool(bool),
    DateTime(f64),
    Duration(f64),
    DateTimeIso(String),
    DurationIso(String),
    Error,
}

impl Cell {
    pub fn new(value: String, is_formula: bool) -> Self {
        let cell_type = if value.is_empty() {
            CellType::Empty
        } else if is_formula {
            CellType::Text
        } else if value.parse::<f64>().is_ok() {
            CellType::Number
        } else if (value.contains('/') && value.split('/').count() == 3)
            || (value.contains('-') && value.split('-').count() == 3)
        {
            CellType::Date
        } else if value == "true" || value == "false" {
            CellType::Boolean
        } else {
            CellType::Text
        };

        Self::new_with_type(value, is_formula, cell_type, None)
    }

    pub fn new_with_type(
        value: String,
        is_formula: bool,
        cell_type: CellType,
        original_type: Option<DataTypeInfo>,
    ) -> Self {
        Self {
            value,
            is_formula,
            cell_type,
            original_type,
        }
    }

    pub fn empty() -> Self {
        Self {
            value: String::new(),
            is_formula: false,
            cell_type: CellType::Empty,
            original_type: Some(DataTypeInfo::Empty),
        }
    }
}
