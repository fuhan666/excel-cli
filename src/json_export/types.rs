use indexmap::IndexMap;
use serde_json::Value;
use std::str::FromStr;

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum HeaderDirection {
    Horizontal,
    Vertical,
}

impl FromStr for HeaderDirection {
    type Err = ();

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s.to_lowercase().as_str() {
            "h" | "horizontal" => Ok(HeaderDirection::Horizontal),
            "v" | "vertical" => Ok(HeaderDirection::Vertical),
            _ => Err(()),
        }
    }
}

pub type OrderedSheetData = Vec<IndexMap<String, Value>>;
