use crate::excel::Sheet;

pub enum Direction {
    Left,
    Right,
    Up,
    Down,
}

// Generic function to find non-empty cells in any direction
pub fn find_non_empty_cell(
    sheet: &Sheet,
    current_pos: (usize, usize),
    direction: Direction,
    max_bounds: (usize, usize),
) -> Option<(usize, usize)> {
    let (row, col) = current_pos;
    let (max_row, max_col) = max_bounds;

    // Check if already at the boundary
    match direction {
        Direction::Left if col <= 1 => return None,
        Direction::Right if col >= max_col => return None,
        Direction::Up if row <= 1 => return None,
        Direction::Down if row >= max_row => return None,
        _ => {}
    }

    // Check if current cell is empty
    let current_cell_is_empty = row >= sheet.data.len()
        || col >= sheet.data[row].len()
        || sheet.data[row][col].value.is_empty();

    if current_cell_is_empty {
        // Find first non-empty cell in the specified direction
        match direction {
            Direction::Left => {
                for c in (1..col).rev() {
                    if row < sheet.data.len()
                        && c < sheet.data[row].len()
                        && !sheet.data[row][c].value.is_empty()
                    {
                        return Some((row, c));
                    }
                }
                // If no non-empty cell found, return the boundary
                Some((row, 1))
            }
            Direction::Right => {
                for c in (col + 1)..=max_col {
                    if row < sheet.data.len()
                        && c < sheet.data[row].len()
                        && !sheet.data[row][c].value.is_empty()
                    {
                        return Some((row, c));
                    }
                }
                // If no non-empty cell found, return the boundary
                Some((row, max_col))
            }
            Direction::Up => {
                for r in (1..row).rev() {
                    if r < sheet.data.len()
                        && col < sheet.data[r].len()
                        && !sheet.data[r][col].value.is_empty()
                    {
                        return Some((r, col));
                    }
                }
                // If no non-empty cell found, return the boundary
                Some((1, col))
            }
            Direction::Down => {
                for r in (row + 1)..=max_row {
                    if r < sheet.data.len()
                        && col < sheet.data[r].len()
                        && !sheet.data[r][col].value.is_empty()
                    {
                        return Some((r, col));
                    }
                }
                // If no non-empty cell found, return the boundary
                Some((max_row, col))
            }
        }
    } else {
        // Find the boundary of non-empty cells in the specified direction
        match direction {
            Direction::Left => {
                let mut last_non_empty = col;

                for c in (1..col).rev() {
                    if row < sheet.data.len() && c < sheet.data[row].len() {
                        if sheet.data[row][c].value.is_empty() {
                            return Some((row, c + 1));
                        } else {
                            last_non_empty = c;
                        }
                    } else {
                        return Some((row, c + 1));
                    }
                }

                Some((row, last_non_empty))
            }
            Direction::Right => {
                let mut last_non_empty = col;

                for c in (col + 1)..=max_col {
                    if row < sheet.data.len() && c < sheet.data[row].len() {
                        if sheet.data[row][c].value.is_empty() {
                            return Some((row, c - 1));
                        } else {
                            last_non_empty = c;
                        }
                    } else {
                        return Some((row, c - 1));
                    }
                }

                Some((row, last_non_empty))
            }
            Direction::Up => {
                let mut last_non_empty = row;

                for r in (1..row).rev() {
                    if r < sheet.data.len() && col < sheet.data[r].len() {
                        if sheet.data[r][col].value.is_empty() {
                            return Some((r + 1, col));
                        } else {
                            last_non_empty = r;
                        }
                    } else {
                        return Some((r + 1, col));
                    }
                }

                Some((last_non_empty, col))
            }
            Direction::Down => {
                let mut last_non_empty = row;

                for r in (row + 1)..=max_row {
                    if r < sheet.data.len() && col < sheet.data[r].len() {
                        if sheet.data[r][col].value.is_empty() {
                            return Some((r - 1, col));
                        } else {
                            last_non_empty = r;
                        }
                    } else {
                        return Some((r - 1, col));
                    }
                }

                Some((last_non_empty, col))
            }
        }
    }
}
