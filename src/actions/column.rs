use super::{ActionType, Command};
use crate::excel::Cell;
use anyhow::Result;

#[derive(Clone)]
pub struct ColumnAction {
    pub sheet_index: usize,
    pub sheet_name: String,
    pub col: usize,
    pub column_data: Vec<Cell>,
    pub column_width: usize,
}

impl Command for ColumnAction {
    fn execute(&self) -> Result<()> {
        unimplemented!("Requires an ActionExecutor implementation")
    }

    fn undo(&self) -> Result<()> {
        unimplemented!("Requires an ActionExecutor implementation")
    }

    fn action_type(&self) -> ActionType {
        ActionType::DeleteColumn
    }
}

#[derive(Clone)]
pub struct MultiColumnAction {
    pub sheet_index: usize,
    pub sheet_name: String,
    pub start_col: usize,
    pub end_col: usize,
    pub columns_data: Vec<Vec<Cell>>,
    pub column_widths: Vec<usize>,
}

impl Command for MultiColumnAction {
    fn execute(&self) -> Result<()> {
        unimplemented!("Requires an ActionExecutor implementation")
    }

    fn undo(&self) -> Result<()> {
        unimplemented!("Requires an ActionExecutor implementation")
    }

    fn action_type(&self) -> ActionType {
        ActionType::DeleteMultiColumns
    }
}
