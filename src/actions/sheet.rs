use super::{ActionType, Command};
use crate::excel::Sheet;
use anyhow::Result;

#[derive(Clone)]
pub struct SheetAction {
    pub sheet_index: usize,
    pub sheet_name: String,
    pub sheet_data: Sheet,
    pub column_widths: Vec<usize>,
}

impl Command for SheetAction {
    fn execute(&self) -> Result<()> {
        unimplemented!("Requires an ActionExecutor implementation")
    }

    fn undo(&self) -> Result<()> {
        unimplemented!("Requires an ActionExecutor implementation")
    }

    fn action_type(&self) -> ActionType {
        ActionType::DeleteSheet
    }
}
