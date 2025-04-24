use crate::excel::Cell;

#[derive(Clone)]
pub struct Sheet {
    pub name: String,
    pub data: Vec<Vec<Cell>>,
    pub max_rows: usize,
    pub max_cols: usize,
}
