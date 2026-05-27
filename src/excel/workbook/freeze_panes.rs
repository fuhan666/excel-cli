use quick_xml::events::Event;
use std::fs::File;
use std::path::Path;
use zip::ZipArchive;

use crate::excel::FreezePanes;

use super::formula_lookup::{attr_value, read_zip_entry, resolve_xlsx_sheet_path};

pub(super) fn lookup_freeze_panes_in_xlsx(file: &Path, sheet_name: &str) -> Option<FreezePanes> {
    let extension = file
        .extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| ext.to_ascii_lowercase())?;
    if extension != "xlsx" && extension != "xlsm" {
        return None;
    }

    let archive_file = File::open(file).ok()?;
    let mut archive = ZipArchive::new(archive_file).ok()?;
    let sheet_path = resolve_xlsx_sheet_path(&mut archive, sheet_name)?;
    let sheet_xml = read_zip_entry(&mut archive, &sheet_path)?;

    let mut reader = quick_xml::Reader::from_str(&sheet_xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf).ok()? {
            Event::Start(event) | Event::Empty(event) if event.name().as_ref() == b"pane" => {
                let state = attr_value(&reader, &event, b"state");
                if state.as_deref() != Some("frozen") {
                    return None;
                }

                let rows = attr_value(&reader, &event, b"ySplit")
                    .and_then(|value| parse_split_count(&value))
                    .unwrap_or(0);
                let cols = attr_value(&reader, &event, b"xSplit")
                    .and_then(|value| parse_split_count(&value))
                    .unwrap_or(0);

                let panes = FreezePanes { rows, cols };
                return panes.is_frozen().then_some(panes);
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    None
}

fn parse_split_count(value: &str) -> Option<usize> {
    if let Ok(count) = value.parse::<usize>() {
        return Some(count);
    }

    let numeric = value.parse::<f64>().ok()?;
    if numeric.fract() == 0.0 && numeric >= 0.0 {
        Some(numeric as usize)
    } else {
        None
    }
}
