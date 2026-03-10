use super::{ActionType, Command};
use crate::excel::Sheet;
use anyhow::Result;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SheetOperation {
    Create,
    Delete,
}

#[derive(Clone)]
pub struct SheetAction {
    pub sheet_index: usize,
    pub sheet_name: String,
    pub sheet_data: Sheet,
    pub column_widths: Vec<usize>,
    pub operation: SheetOperation,
}

impl Command for SheetAction {
    fn execute(&self) -> Result<()> {
        unimplemented!("Requires an ActionExecutor implementation")
    }

    fn undo(&self) -> Result<()> {
        unimplemented!("Requires an ActionExecutor implementation")
    }

    fn action_type(&self) -> ActionType {
        match self.operation {
            SheetOperation::Create => ActionType::CreateSheet,
            SheetOperation::Delete => ActionType::DeleteSheet,
        }
    }
}
