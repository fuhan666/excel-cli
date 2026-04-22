use ratatui::style::{Color, Style};

pub const BACKGROUND: Color = Color::Rgb(11, 16, 32);
pub const SURFACE: Color = Color::Rgb(17, 24, 39);
pub const SURFACE_MUTED: Color = Color::Rgb(31, 41, 55);
pub const GRID: Color = Color::Rgb(55, 65, 81);
pub const TEXT: Color = Color::Rgb(229, 231, 235);
pub const TEXT_SECONDARY: Color = Color::Rgb(156, 163, 175);
pub const TEXT_DISABLED: Color = Color::Rgb(107, 114, 128);
pub const ACCENT: Color = Color::Rgb(56, 189, 248);
pub const SEARCH: Color = Color::Rgb(250, 204, 21);
pub const WARNING: Color = Color::Rgb(245, 158, 11);
pub const SUCCESS: Color = Color::Rgb(34, 197, 94);

pub fn base() -> Style {
    Style::default().bg(BACKGROUND).fg(TEXT)
}

pub fn surface() -> Style {
    Style::default().bg(SURFACE).fg(TEXT)
}

pub fn muted() -> Style {
    Style::default().bg(SURFACE_MUTED).fg(TEXT_SECONDARY)
}
