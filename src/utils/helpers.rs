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
