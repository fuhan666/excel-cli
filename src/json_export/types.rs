use indexmap::IndexMap;
use serde_json::Value;

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum HeaderDirection {
    Horizontal,
    Vertical,
}

impl HeaderDirection {
    pub fn from_str(s: &str) -> Option<Self> {
        match s.to_lowercase().as_str() {
            "h" | "horizontal" => Some(HeaderDirection::Horizontal),
            "v" | "vertical" => Some(HeaderDirection::Vertical),
            _ => None,
        }
    }
}

pub type OrderedSheetData = Vec<IndexMap<String, Value>>;
