use crate::excel::Cell;
use crate::utils::cell_reference;

#[derive(Clone)]
pub struct FreezePanes {
    pub rows: usize,
    pub cols: usize,
}

impl FreezePanes {
    #[must_use]
    pub fn none() -> Self {
        Self { rows: 0, cols: 0 }
    }

    #[must_use]
    pub fn from_split_cell(row: usize, col: usize) -> Self {
        Self {
            rows: row.saturating_sub(1),
            cols: col.saturating_sub(1),
        }
    }

    #[must_use]
    pub fn is_frozen(&self) -> bool {
        self.rows > 0 || self.cols > 0
    }

    #[must_use]
    pub fn split_cell(&self) -> (usize, usize) {
        (self.rows + 1, self.cols + 1)
    }

    #[must_use]
    pub fn split_cell_ref(&self) -> String {
        cell_reference(self.split_cell())
    }
}

impl Default for FreezePanes {
    fn default() -> Self {
        Self::none()
    }
}

#[derive(Clone)]
pub struct Sheet {
    pub name: String,
    pub data: Vec<Vec<Cell>>,
    pub max_rows: usize,
    pub max_cols: usize,
    pub is_loaded: bool,
    pub freeze_panes: FreezePanes,
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
            freeze_panes: FreezePanes::none(),
        }
    }
}
