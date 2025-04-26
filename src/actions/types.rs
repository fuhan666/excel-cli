#[derive(Clone, Debug)]
pub enum ActionType {
    Edit,
    Cut,
    Paste,
    DeleteRow,
    DeleteColumn,
    DeleteSheet,
    DeleteMultiRows,
    DeleteMultiColumns,
}

// Executor for actions in the application
pub trait ActionExecutor {
    fn execute_action(&mut self, action: &ActionCommand) -> Result<(), anyhow::Error>;
    fn execute_cell_action(
        &mut self,
        action: &crate::actions::CellAction,
    ) -> Result<(), anyhow::Error>;
    fn execute_row_action(
        &mut self,
        action: &crate::actions::RowAction,
    ) -> Result<(), anyhow::Error>;
    fn execute_column_action(
        &mut self,
        action: &crate::actions::ColumnAction,
    ) -> Result<(), anyhow::Error>;
    fn execute_sheet_action(
        &mut self,
        action: &crate::actions::SheetAction,
    ) -> Result<(), anyhow::Error>;
    fn execute_multi_row_action(
        &mut self,
        action: &crate::actions::MultiRowAction,
    ) -> Result<(), anyhow::Error>;
    fn execute_multi_column_action(
        &mut self,
        action: &crate::actions::MultiColumnAction,
    ) -> Result<(), anyhow::Error>;
}

// Command interface for actions that can be executed and undone
pub trait Command {
    fn execute(&self) -> anyhow::Result<()>;
    fn undo(&self) -> anyhow::Result<()>;
    fn action_type(&self) -> ActionType;
}

// Unified action command enum for all action types
#[derive(Clone)]
pub enum ActionCommand {
    Cell(crate::actions::CellAction),
    Row(crate::actions::RowAction),
    Column(crate::actions::ColumnAction),
    Sheet(crate::actions::SheetAction),
    MultiRow(crate::actions::MultiRowAction),
    MultiColumn(crate::actions::MultiColumnAction),
}
