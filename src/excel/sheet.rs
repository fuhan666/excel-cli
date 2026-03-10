use crate::excel::Cell;

#[derive(Clone)]
pub struct Sheet {
    pub name: String,
    pub data: Vec<Vec<Cell>>,
    pub max_rows: usize,
    pub max_cols: usize,
    pub is_loaded: bool,
}

impl Sheet {
    #[must_use]
    pub fn blank(name: String) -> Self {
        Self {
            name,
            data: vec![vec![Cell::empty(); 2]; 2],
            max_rows: 1,
            max_cols: 1,
            is_loaded: true,
        }
    }
}
