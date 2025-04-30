use ratatui::style::{Color, Modifier, Style};
use ratatui::widgets::Block;
use std::fmt;
use tui_textarea::{CursorMove, Input, Key, TextArea};

use crate::app::word::move_cursor_to_word_end;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VimMode {
    Normal,
    Insert,
    Visual,
    Operator(char),
}

impl VimMode {
    pub fn block<'a>(&self) -> Block<'a> {
        let help = match self {
            Self::Normal => "Esc=exit, i=insert, v=visual, y/d/c=operator",
            Self::Insert => "Esc=normal mode",
            Self::Visual => "Esc=normal, y=yank, d=delete, c=change",
            Self::Operator(_) => "Move cursor to apply operator",
        };
        let title = format!(" {} MODE ({}) ", self, help);
        Block::default().title(title)
    }

    pub fn cursor_style(&self) -> Style {
        let color = match self {
            Self::Normal => Color::Reset,
            Self::Insert => Color::LightBlue,
            Self::Visual => Color::LightYellow,
            Self::Operator(_) => Color::LightGreen,
        };
        Style::default().fg(color).add_modifier(Modifier::REVERSED)
    }
}

impl fmt::Display for VimMode {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> Result<(), fmt::Error> {
        match self {
            Self::Normal => write!(f, "NORMAL"),
            Self::Insert => write!(f, "INSERT"),
            Self::Visual => write!(f, "VISUAL"),
            Self::Operator(c) => write!(f, "OPERATOR({})", c),
        }
    }
}

// How the Vim emulation state transitions
pub enum Transition {
    Nop,
    Mode(VimMode),
    Pending(Input),
    Exit,
}

// State of Vim emulation
#[derive(Clone)]
pub struct VimState {
    pub mode: VimMode,
    pub pending: Input, // Pending input to handle a sequence with two keys like gg
}

impl VimState {
    pub fn new(mode: VimMode) -> Self {
        Self {
            mode,
            pending: Input::default(),
        }
    }

    pub fn with_pending(self, pending: Input) -> Self {
        Self {
            mode: self.mode,
            pending,
        }
    }

    pub fn transition(&self, input: Input, textarea: &mut TextArea<'_>) -> Transition {
        if input.key == Key::Null {
            return Transition::Nop;
        }

        match self.mode {
            VimMode::Normal | VimMode::Visual | VimMode::Operator(_) => {
                match input {
                    // Navigation
                    Input {
                        key: Key::Char('h'),
                        ..
                    } => textarea.move_cursor(CursorMove::Back),
                    Input {
                        key: Key::Char('j'),
                        ..
                    } => textarea.move_cursor(CursorMove::Down),
                    Input {
                        key: Key::Char('k'),
                        ..
                    } => textarea.move_cursor(CursorMove::Up),
                    Input {
                        key: Key::Char('l'),
                        ..
                    } => textarea.move_cursor(CursorMove::Forward),
                    Input {
                        key: Key::Char('w'),
                        ..
                    } => textarea.move_cursor(CursorMove::WordForward),
                    Input {
                        key: Key::Char('e'),
                        ctrl: false,
                        ..
                    } => {
                        // Use custom WordEnd implementation
                        let lines = textarea.lines();
                        let (row, col) = textarea.cursor();
                        let (new_row, new_col) = move_cursor_to_word_end(lines, row, col);

                        // Set the cursor to the new position
                        if row != new_row {
                            // If need to move to a different row
                            while textarea.cursor().0 < new_row {
                                textarea.move_cursor(CursorMove::Down);
                            }
                            textarea.move_cursor(CursorMove::Head);
                            while textarea.cursor().1 < new_col {
                                textarea.move_cursor(CursorMove::Forward);
                            }
                        } else {
                            // If staying on the same row
                            if col < new_col {
                                // Move forward
                                while textarea.cursor().1 < new_col {
                                    textarea.move_cursor(CursorMove::Forward);
                                }
                            } else {
                                // Move backward
                                while textarea.cursor().1 > new_col {
                                    textarea.move_cursor(CursorMove::Back);
                                }
                            }
                        }

                        // For operator mode, include the character under the cursor
                        if matches!(self.mode, VimMode::Operator(_)) {
                            textarea.move_cursor(CursorMove::Forward);
                        }
                    }
                    Input {
                        key: Key::Char('b'),
                        ctrl: false,
                        ..
                    } => textarea.move_cursor(CursorMove::WordBack),
                    Input {
                        key: Key::Char('^'),
                        ..
                    } => textarea.move_cursor(CursorMove::Head),
                    Input {
                        key: Key::Char('$'),
                        ..
                    } => textarea.move_cursor(CursorMove::End),

                    // Editing operations
                    Input {
                        key: Key::Char('D'),
                        ..
                    } => {
                        textarea.delete_line_by_end();
                        return Transition::Mode(VimMode::Normal);
                    }
                    Input {
                        key: Key::Char('C'),
                        ..
                    } => {
                        textarea.delete_line_by_end();
                        textarea.cancel_selection();
                        return Transition::Mode(VimMode::Insert);
                    }
                    Input {
                        key: Key::Char('p'),
                        ..
                    } => {
                        textarea.paste();
                        return Transition::Mode(VimMode::Normal);
                    }
                    Input {
                        key: Key::Char('u'),
                        ctrl: false,
                        ..
                    } => {
                        textarea.undo();
                        return Transition::Mode(VimMode::Normal);
                    }
                    Input {
                        key: Key::Char('r'),
                        ctrl: true,
                        ..
                    } => {
                        textarea.redo();
                        return Transition::Mode(VimMode::Normal);
                    }
                    Input {
                        key: Key::Char('x'),
                        ..
                    } => {
                        textarea.delete_next_char();
                        return Transition::Mode(VimMode::Normal);
                    }

                    // Mode changes
                    Input {
                        key: Key::Char('i'),
                        ..
                    } => {
                        textarea.cancel_selection();
                        return Transition::Mode(VimMode::Insert);
                    }
                    Input {
                        key: Key::Char('a'),
                        ..
                    } => {
                        textarea.cancel_selection();
                        textarea.move_cursor(CursorMove::Forward);
                        return Transition::Mode(VimMode::Insert);
                    }
                    Input {
                        key: Key::Char('A'),
                        ..
                    } => {
                        textarea.cancel_selection();
                        textarea.move_cursor(CursorMove::End);
                        return Transition::Mode(VimMode::Insert);
                    }
                    Input {
                        key: Key::Char('I'),
                        ..
                    } => {
                        textarea.cancel_selection();
                        textarea.move_cursor(CursorMove::Head);
                        return Transition::Mode(VimMode::Insert);
                    }
                    Input {
                        key: Key::Char('o'),
                        ..
                    } => {
                        textarea.move_cursor(CursorMove::End);
                        textarea.insert_newline();
                        return Transition::Mode(VimMode::Insert);
                    }
                    Input {
                        key: Key::Char('O'),
                        ..
                    } => {
                        textarea.move_cursor(CursorMove::Head);
                        textarea.insert_newline();
                        textarea.move_cursor(CursorMove::Up);
                        return Transition::Mode(VimMode::Insert);
                    }

                    // Exit
                    Input { key: Key::Esc, .. } => {
                        if self.mode == VimMode::Visual {
                            textarea.cancel_selection();
                            return Transition::Mode(VimMode::Normal);
                        } else {
                            return Transition::Exit;
                        }
                    }

                    // Scrolling
                    Input {
                        key: Key::Char('e'),
                        ctrl: true,
                        ..
                    } => textarea.scroll((1, 0)),
                    Input {
                        key: Key::Char('y'),
                        ctrl: true,
                        ..
                    } => textarea.scroll((-1, 0)),

                    // Visual mode
                    Input {
                        key: Key::Char('v'),
                        ctrl: false,
                        ..
                    } if self.mode == VimMode::Normal => {
                        textarea.start_selection();
                        return Transition::Mode(VimMode::Visual);
                    }
                    Input {
                        key: Key::Char('V'),
                        ctrl: false,
                        ..
                    } if self.mode == VimMode::Normal => {
                        textarea.move_cursor(CursorMove::Head);
                        textarea.start_selection();
                        textarea.move_cursor(CursorMove::End);
                        return Transition::Mode(VimMode::Visual);
                    }

                    // Operators
                    Input {
                        key: Key::Char('g'),
                        ctrl: false,
                        ..
                    } if matches!(
                        self.pending,
                        Input {
                            key: Key::Char('g'),
                            ctrl: false,
                            ..
                        }
                    ) =>
                    {
                        textarea.move_cursor(CursorMove::Top)
                    }
                    Input {
                        key: Key::Char('G'),
                        ctrl: false,
                        ..
                    } => textarea.move_cursor(CursorMove::Bottom),
                    Input {
                        key: Key::Char('y'),
                        ctrl: false,
                        ..
                    } if self.mode == VimMode::Visual => {
                        textarea.move_cursor(CursorMove::Forward); // Vim's text selection is inclusive
                        textarea.copy();
                        return Transition::Mode(VimMode::Normal);
                    }
                    Input {
                        key: Key::Char('d'),
                        ctrl: false,
                        ..
                    } if self.mode == VimMode::Visual => {
                        textarea.move_cursor(CursorMove::Forward); // Vim's text selection is inclusive
                        textarea.cut();
                        return Transition::Mode(VimMode::Normal);
                    }
                    Input {
                        key: Key::Char('c'),
                        ctrl: false,
                        ..
                    } if self.mode == VimMode::Visual => {
                        textarea.move_cursor(CursorMove::Forward); // Vim's text selection is inclusive
                        textarea.cut();
                        return Transition::Mode(VimMode::Insert);
                    }
                    Input {
                        key: Key::Char(op @ ('y' | 'd' | 'c')),
                        ctrl: false,
                        ..
                    } if self.mode == VimMode::Normal => {
                        textarea.start_selection();
                        return Transition::Mode(VimMode::Operator(op));
                    }

                    input => return Transition::Pending(input),
                }

                // Handle the pending operator
                match self.mode {
                    VimMode::Operator('y') => {
                        textarea.copy();
                        Transition::Mode(VimMode::Normal)
                    }
                    VimMode::Operator('d') => {
                        textarea.cut();
                        Transition::Mode(VimMode::Normal)
                    }
                    VimMode::Operator('c') => {
                        textarea.cut();
                        Transition::Mode(VimMode::Insert)
                    }
                    _ => Transition::Nop,
                }
            }
            VimMode::Insert => match input {
                Input { key: Key::Esc, .. }
                | Input {
                    key: Key::Char('c'),
                    ctrl: true,
                    ..
                } => Transition::Mode(VimMode::Normal),
                input => {
                    textarea.input(input); // Use default key mappings in insert mode
                    Transition::Mode(VimMode::Insert)
                }
            },
        }
    }
}
