use super::ActionCommand;
use std::rc::Rc;

pub struct UndoHistory {
    undo_stack: Vec<Rc<ActionCommand>>,
    redo_stack: Vec<Rc<ActionCommand>>,
}

impl Default for UndoHistory {
    fn default() -> Self {
        Self::new()
    }
}

impl UndoHistory {
    pub fn new() -> Self {
        Self {
            undo_stack: Vec::with_capacity(100), // Pre-allocate capacity
            redo_stack: Vec::with_capacity(20),
        }
    }

    pub fn push(&mut self, action: ActionCommand) {
        // Use Rc to avoid deep cloning the entire action
        self.undo_stack.push(Rc::new(action));
        self.redo_stack.clear();
    }

    pub fn undo(&mut self) -> Option<Rc<ActionCommand>> {
        if let Some(action) = self.undo_stack.pop() {
            self.redo_stack.push(Rc::clone(&action));
            Some(action)
        } else {
            None
        }
    }

    pub fn redo(&mut self) -> Option<Rc<ActionCommand>> {
        if let Some(action) = self.redo_stack.pop() {
            self.undo_stack.push(Rc::clone(&action));
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
