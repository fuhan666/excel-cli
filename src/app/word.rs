// Custom implementation of word navigation functions from tui-textarea v0.5.2+

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum CharKind {
    Space,
    Punctuation,
    Other,
}

impl CharKind {
    fn new(c: char) -> Self {
        if c.is_whitespace() {
            Self::Space
        } else if c.is_ascii_punctuation() {
            Self::Punctuation
        } else {
            Self::Other
        }
    }
}

/// Find the end of the next word
/// This is a custom implementation of the `find_word_end_next` function from tui-textarea v0.5.2+
pub fn find_word_end_next(line: &str, start_col: usize) -> Option<usize> {
    let mut it = line.chars().enumerate().skip(start_col);
    let (mut cur_col, cur_char) = it.next()?;
    let mut cur = CharKind::new(cur_char);

    for (next_col, c) in it {
        let next = CharKind::new(c);
        // if cursor started at the end of a word, don't stop
        if next_col.saturating_sub(start_col) > 1 && cur != CharKind::Space && next != cur {
            return Some(next_col.saturating_sub(1));
        }
        cur = next;
        cur_col = next_col;
    }

    // if end of line is whitespace, don't stop the cursor
    if cur != CharKind::Space && cur_col.saturating_sub(start_col) >= 1 {
        return Some(cur_col);
    }

    None
}

/// Move cursor to the end of the next word
pub fn move_cursor_to_word_end(text: &[String], row: usize, col: usize) -> (usize, usize) {
    if row >= text.len() {
        return (row, col);
    }

    let line = &text[row];

    if let Some(new_col) = find_word_end_next(line, col) {
        return (row, new_col);
    } else if row + 1 < text.len() {
        // Try to find word end in the next line
        if let Some(new_col) = find_word_end_next(&text[row + 1], 0) {
            return (row + 1, new_col);
        } else if !text[row + 1].is_empty() {
            // If no word end found but line is not empty, go to the end of the line
            return (row + 1, text[row + 1].chars().count().saturating_sub(1));
        }
    }

    // Can't find a word end, stay at the current position
    (row, col)
}
