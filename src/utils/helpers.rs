#[must_use]
pub fn index_to_col_name(index: usize) -> String {
    let mut col_name = String::new();
    let mut n = index;

    while n > 0 {
        let remainder = (n - 1) % 26;
        col_name.insert(0, (b'A' + remainder as u8) as char);
        n = (n - 1) / 26;
    }

    if col_name.is_empty() {
        col_name.push('A');
    }

    col_name
}

#[must_use]
pub fn col_name_to_index(name: &str) -> Option<usize> {
    let mut result = 0;

    for c in name.chars() {
        if !c.is_ascii_alphabetic() {
            return None;
        }

        let val = (c.to_ascii_uppercase() as u8 - b'A' + 1) as usize;
        result = result * 26 + val;
    }

    Some(result)
}

// Format cell reference (e.g., A1, B2)
#[must_use]
pub fn cell_reference(cell: (usize, usize)) -> String {
    format!("{}{}", index_to_col_name(cell.1), cell.0)
}

/// Parse a cell reference like "A1" or "AB123" into (row, col) using 1-based indexing.
#[must_use]
pub fn parse_cell_reference(reference: &str) -> Option<(usize, usize)> {
    let reference = reference.trim();
    if reference.is_empty() {
        return None;
    }

    let first_digit_pos = reference
        .chars()
        .position(|c| c.is_ascii_digit())
        .unwrap_or(reference.len());

    if first_digit_pos == 0 || first_digit_pos == reference.len() {
        return None;
    }

    let col_part = &reference[..first_digit_pos];
    let row_part = &reference[first_digit_pos..];

    let col = col_name_to_index(col_part)?;
    let row = row_part.parse::<usize>().ok()?;

    if row == 0 {
        return None;
    }

    Some((row, col))
}

/// Parse a range like "A1:F10" into ((start_row, start_col), (end_row, end_col)) using 1-based indexing.
#[must_use]
pub fn parse_range(range: &str) -> Option<((usize, usize), (usize, usize))> {
    let range = range.trim();
    if range.is_empty() {
        return None;
    }

    let parts: Vec<&str> = range.split(':').collect();
    if parts.len() != 2 {
        return None;
    }

    let start = parse_cell_reference(parts[0])?;
    let end = parse_cell_reference(parts[1])?;

    Some((start, end))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_cell_reference_basic() {
        assert_eq!(parse_cell_reference("A1"), Some((1, 1)));
        assert_eq!(parse_cell_reference("B2"), Some((2, 2)));
        assert_eq!(parse_cell_reference("Z26"), Some((26, 26)));
        assert_eq!(parse_cell_reference("AA27"), Some((27, 27)));
        assert_eq!(parse_cell_reference("AB28"), Some((28, 28)));
    }

    #[test]
    fn test_parse_cell_reference_case_insensitive() {
        assert_eq!(parse_cell_reference("a1"), Some((1, 1)));
        assert_eq!(parse_cell_reference("aA100"), Some((100, 27)));
    }

    #[test]
    fn test_parse_cell_reference_invalid() {
        assert_eq!(parse_cell_reference(""), None);
        assert_eq!(parse_cell_reference("1A"), None);
        assert_eq!(parse_cell_reference("A"), None);
        assert_eq!(parse_cell_reference("0"), None);
        assert_eq!(parse_cell_reference("A0"), None);
        assert_eq!(parse_cell_reference("AAA"), None);
    }

    #[test]
    fn test_parse_range_basic() {
        assert_eq!(parse_range("A1:B2"), Some(((1, 1), (2, 2))));
        assert_eq!(parse_range("A1:F10"), Some(((1, 1), (10, 6))));
    }

    #[test]
    fn test_parse_range_invalid() {
        assert_eq!(parse_range(""), None);
        assert_eq!(parse_range("A1"), None);
        assert_eq!(parse_range("A1:B"), None);
        assert_eq!(parse_range("A:B2"), None);
        assert_eq!(parse_range("A1:B2:C3"), None);
    }
}
