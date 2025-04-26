use super::{ActionType, Command};
use crate::excel::Cell;
use anyhow::Result;

#[derive(Clone)]
pub struct RowAction {
    pub sheet_index: usize,
    pub sheet_name: String,
    pub row: usize,
    pub row_data: Vec<Cell>,
}

impl Command for RowAction {
    fn execute(&self) -> Result<()> {
        unimplemented!("Requires an ActionExecutor implementation")
    }

    fn undo(&self) -> Result<()> {
        unimplemented!("Requires an ActionExecutor implementation")
    }

    fn action_type(&self) -> ActionType {
        ActionType::DeleteRow
    }
}

#[derive(Clone)]
pub struct MultiRowAction {
    pub sheet_index: usize,
    pub sheet_name: String,
    pub start_row: usize,
    pub end_row: usize,
    pub rows_data: Vec<Vec<Cell>>,
}

impl Command for MultiRowAction {
    fn execute(&self) -> Result<()> {
        unimplemented!("Requires an ActionExecutor implementation")
    }

    fn undo(&self) -> Result<()> {
        unimplemented!("Requires an ActionExecutor implementation")
    }

    fn action_type(&self) -> ActionType {
        ActionType::DeleteMultiRows
    }
}
