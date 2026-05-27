pub struct HelpEntry {
    pub keys: &'static str,
    pub description: &'static str,
}

pub struct HelpSection {
    pub title: &'static str,
    pub entries: &'static [HelpEntry],
}

pub const LEFT_HELP_SECTIONS: &[HelpSection] = &[
    HelpSection {
        title: "NAVIGATION",
        entries: &[
            HelpEntry {
                keys: "h j k l / arrows",
                description: "Move cell",
            },
            HelpEntry {
                keys: "[ / ]",
                description: "Switch sheet",
            },
            HelpEntry {
                keys: "gg / G",
                description: "Start/end of data",
            },
            HelpEntry {
                keys: "0 / ^ / $",
                description: "Row start / first non-empty / end",
            },
            HelpEntry {
                keys: "Ctrl+arrows",
                description: "Jump to next non-empty cell",
            },
        ],
    },
    HelpSection {
        title: "SEARCH",
        entries: &[
            HelpEntry {
                keys: "/",
                description: "Search forward",
            },
            HelpEntry {
                keys: "?",
                description: "Search backward",
            },
            HelpEntry {
                keys: "n / N",
                description: "Next/previous search result",
            },
            HelpEntry {
                keys: ":noh / :nohlsearch",
                description: "Disable search highlighting",
            },
        ],
    },
    HelpSection {
        title: "JUMP & SHEETS",
        entries: &[
            HelpEntry {
                keys: ":<cell>",
                description: "Jump to cell, e.g. :B10",
            },
            HelpEntry {
                keys: ":sheet <name|index>",
                description: "Switch sheet",
            },
            HelpEntry {
                keys: ":addsheet <name>",
                description: "Add sheet after current",
            },
            HelpEntry {
                keys: ":delsheet",
                description: "Delete current sheet",
            },
        ],
    },
    HelpSection {
        title: "ROWS & COLUMNS",
        entries: &[
            HelpEntry {
                keys: ":cw fit",
                description: "Fit current column",
            },
            HelpEntry {
                keys: ":cw fit all",
                description: "Fit all columns",
            },
            HelpEntry {
                keys: ":cw min",
                description: "Minimize current column",
            },
            HelpEntry {
                keys: ":cw min all",
                description: "Minimize all columns",
            },
            HelpEntry {
                keys: ":cw <number>",
                description: "Set current column width",
            },
            HelpEntry {
                keys: ":dr / :dr <row>",
                description: "Delete current/specific row",
            },
            HelpEntry {
                keys: ":dr <start> <end>",
                description: "Delete row range",
            },
            HelpEntry {
                keys: ":dc / :dc <col>",
                description: "Delete current/specific column",
            },
            HelpEntry {
                keys: ":dc <start> <end>",
                description: "Delete column range",
            },
            HelpEntry {
                keys: ":freeze [cell]",
                description: "Freeze panes at cell",
            },
            HelpEntry {
                keys: ":unfreeze",
                description: "Clear frozen panes",
            },
        ],
    },
];

pub const RIGHT_HELP_SECTIONS: &[HelpSection] = &[
    HelpSection {
        title: "ACTIONS",
        entries: &[
            HelpEntry {
                keys: "Enter",
                description: "Edit cell",
            },
            HelpEntry {
                keys: "y / :y",
                description: "Copy current cell",
            },
            HelpEntry {
                keys: "d / :d",
                description: "Cut current cell",
            },
            HelpEntry {
                keys: "p / :put / :pu",
                description: "Paste to current cell",
            },
            HelpEntry {
                keys: "u",
                description: "Undo",
            },
            HelpEntry {
                keys: "Ctrl+r",
                description: "Redo",
            },
            HelpEntry {
                keys: "+ / = / -",
                description: "Resize info panel",
            },
        ],
    },
    HelpSection {
        title: "FILE & APP",
        entries: &[
            HelpEntry {
                keys: ":w",
                description: "Save file",
            },
            HelpEntry {
                keys: ":wq / :x",
                description: "Save and quit",
            },
            HelpEntry {
                keys: ":q",
                description: "Quit, warn if unsaved",
            },
            HelpEntry {
                keys: ":q!",
                description: "Force quit without saving",
            },
            HelpEntry {
                keys: ":help",
                description: "Show this overlay",
            },
        ],
    },
    HelpSection {
        title: "EXPORT",
        entries: &[
            HelpEntry {
                keys: ":ej",
                description: "Export current sheet JSON",
            },
            HelpEntry {
                keys: ":ej <h|v> <rows>",
                description: "Export with header direction/count",
            },
            HelpEntry {
                keys: ":eja",
                description: "Export all sheets JSON",
            },
            HelpEntry {
                keys: ":eja <h|v> <rows>",
                description: "Export all with header settings",
            },
        ],
    },
    HelpSection {
        title: "EDIT MODE",
        entries: &[
            HelpEntry {
                keys: "Esc",
                description: "Save edits and return",
            },
            HelpEntry {
                keys: "i / v",
                description: "Insert / visual mode",
            },
            HelpEntry {
                keys: "h j k l",
                description: "Move cursor",
            },
            HelpEntry {
                keys: "w / b / e",
                description: "Word navigation",
            },
            HelpEntry {
                keys: "^ / $",
                description: "Line start / end",
            },
            HelpEntry {
                keys: "gg / G",
                description: "First / last line",
            },
            HelpEntry {
                keys: "x / D / C",
                description: "Delete/change text",
            },
            HelpEntry {
                keys: "y / d / c",
                description: "Operator commands",
            },
            HelpEntry {
                keys: "p / u / Ctrl+r",
                description: "Paste / undo / redo",
            },
            HelpEntry {
                keys: "o / O / A / I",
                description: "Open or insert at line edges",
            },
        ],
    },
    HelpSection {
        title: "HELP CONTROLS",
        entries: &[
            HelpEntry {
                keys: "Esc / q / Enter",
                description: "Close overlay",
            },
            HelpEntry {
                keys: "j / k / arrows",
                description: "Scroll one line",
            },
            HelpEntry {
                keys: "PgUp / PgDn",
                description: "Scroll one page",
            },
            HelpEntry {
                keys: "Home / End",
                description: "Jump to top/bottom",
            },
        ],
    },
];

pub fn help_reference_line_count() -> usize {
    column_line_count(LEFT_HELP_SECTIONS) + column_line_count(RIGHT_HELP_SECTIONS) + 1
}

pub fn help_reference_text() -> String {
    let mut lines = Vec::new();

    for section in LEFT_HELP_SECTIONS.iter().chain(RIGHT_HELP_SECTIONS.iter()) {
        lines.push(section.title.to_string());
        for entry in section.entries {
            lines.push(format!("{} - {}", entry.keys, entry.description));
        }
        lines.push(String::new());
    }

    lines.join("\n")
}

fn column_line_count(sections: &[HelpSection]) -> usize {
    sections
        .iter()
        .map(|section| section.entries.len() + 2)
        .sum::<usize>()
        .saturating_sub(1)
}
