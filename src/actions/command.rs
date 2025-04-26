use super::types::{ActionCommand, ActionType};

impl ActionCommand {
    // Returns the action type of this command
    pub fn get_action_type(&self) -> ActionType {
        match self {
            ActionCommand::Cell(action) => match action.action_type {
                ActionType::Edit => ActionType::Edit,
                ActionType::Cut => ActionType::Cut,
                ActionType::Paste => ActionType::Paste,
                _ => ActionType::Edit, // Default case
            },
            ActionCommand::Row(_) => ActionType::DeleteRow,
            ActionCommand::MultiRow(_) => ActionType::DeleteMultiRows,
            ActionCommand::Column(_) => ActionType::DeleteColumn,
            ActionCommand::MultiColumn(_) => ActionType::DeleteMultiColumns,
            ActionCommand::Sheet(_) => ActionType::DeleteSheet,
        }
    }
}
