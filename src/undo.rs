use crate::excel::{Cell, Sheet};

#[derive(Clone)]
pub struct CellAction {
    pub sheet_index: usize,
    pub sheet_name: String,
    pub row: usize,
    pub col: usize,
    pub old_value: Cell,
    pub new_value: Cell,
}

#[derive(Clone)]
pub struct RowAction {
    pub sheet_index: usize,
    pub sheet_name: String,
    pub row: usize,
    pub row_data: Vec<Cell>,
}

#[derive(Clone)]
pub struct MultiRowAction {
    pub sheet_index: usize,
    pub sheet_name: String,
    pub start_row: usize,
    pub end_row: usize,
    pub rows_data: Vec<Vec<Cell>>,
}

#[derive(Clone)]
pub struct ColumnAction {
    pub sheet_index: usize,
    pub sheet_name: String,
    pub col: usize,
    pub column_data: Vec<Cell>,
    pub column_width: usize,
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

#[derive(Clone)]
pub struct SheetAction {
    pub sheet_index: usize,
    pub sheet_name: String,
    pub sheet_data: Sheet,
    pub column_widths: Vec<usize>,
}

#[derive(Clone)]
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

#[derive(Clone)]
pub enum ActionData {
    Cell(CellAction),
    Row(RowAction),
    Column(ColumnAction),
    Sheet(SheetAction),
    MultiRow(MultiRowAction),
    MultiColumn(MultiColumnAction),
}

#[derive(Clone)]
pub struct UndoAction {
    pub action_type: ActionType,
    pub action_data: ActionData,
}

#[derive(Default)]
pub struct UndoHistory {
    undo_stack: Vec<UndoAction>,
    redo_stack: Vec<UndoAction>,
}

impl UndoHistory {
    pub fn new() -> Self {
        Self {
            undo_stack: Vec::new(),
            redo_stack: Vec::new(),
        }
    }

    pub fn push(&mut self, action: UndoAction) {
        self.undo_stack.push(action);
        self.redo_stack.clear();
    }

    pub fn undo(&mut self) -> Option<UndoAction> {
        if let Some(action) = self.undo_stack.pop() {
            self.redo_stack.push(action.clone());
            Some(action)
        } else {
            None
        }
    }

    pub fn redo(&mut self) -> Option<UndoAction> {
        if let Some(action) = self.redo_stack.pop() {
            self.undo_stack.push(action.clone());
            Some(action)
        } else {
            None
        }
    }

    pub fn all_undone(&self) -> bool {
        self.undo_stack.is_empty()
    }

    pub fn clear(&mut self) {
        self.undo_stack.clear();
        self.redo_stack.clear();
    }
}
