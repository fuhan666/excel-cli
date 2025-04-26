use super::{ActionType, Command};
use crate::excel::Cell;
use anyhow::Result;

#[derive(Clone)]
pub struct CellAction {
    pub sheet_index: usize,
    pub sheet_name: String,
    pub row: usize,
    pub col: usize,
    pub old_value: Cell,
    pub new_value: Cell,
    pub action_type: ActionType,
}

impl CellAction {
    pub fn new(
        sheet_index: usize,
        sheet_name: String,
        row: usize,
        col: usize,
        old_value: Cell,
        new_value: Cell,
        action_type: ActionType,
    ) -> Self {
        Self {
            sheet_index,
            sheet_name,
            row,
            col,
            old_value,
            new_value,
            action_type,
        }
    }
}

impl Command for CellAction {
    fn execute(&self) -> Result<()> {
        unimplemented!("Requires an ActionExecutor implementation")
    }

    fn undo(&self) -> Result<()> {
        unimplemented!("Requires an ActionExecutor implementation")
    }

    fn action_type(&self) -> ActionType {
        self.action_type.clone()
    }
}
