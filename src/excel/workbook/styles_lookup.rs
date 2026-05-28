use quick_xml::events::Event;
use std::collections::HashMap;
use std::fs::File;
use std::path::Path;
use zip::ZipArchive;

use super::formula_lookup::{attr_value, read_zip_entry, resolve_xlsx_sheet_path};

/// Look up cell background styles (Hex RRGGBB) in an XLSX file for a given sheet.
pub(crate) fn lookup_styles_in_xlsx(
    file: &Path,
    sheet_name: &str,
) -> Option<HashMap<String, String>> {
    let extension = file
        .extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| ext.to_ascii_lowercase())?;
    if extension != "xlsx" && extension != "xlsm" {
        return None;
    }

    let archive_file = File::open(file).ok()?;
    let mut archive = ZipArchive::new(archive_file).ok()?;

    // 1. Read xl/styles.xml
    let styles_xml = read_zip_entry(&mut archive, "xl/styles.xml")?;
    let (fills, cell_xfs) = parse_styles_xml(&styles_xml)?;

    // 2. Resolve sheet path and read sheet XML
    let sheet_path = resolve_xlsx_sheet_path(&mut archive, sheet_name)?;
    let sheet_xml = read_zip_entry(&mut archive, &sheet_path)?;

    // 3. Parse sheet XML to map cells to color hex
    let cell_colors = parse_sheet_styles(&sheet_xml, &fills, &cell_xfs)?;
    Some(cell_colors)
}

fn parse_styles_xml(xml: &str) -> Option<(Vec<Option<String>>, Vec<usize>)> {
    let mut reader = quick_xml::Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut fills = Vec::new();
    let mut cell_xfs = Vec::new();

    let mut in_fills = false;
    let mut in_cell_xfs = false;
    let mut current_fill_fg_color = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(event)) => {
                let tag_name = event.name();
                if tag_name.as_ref() == b"fills" {
                    in_fills = true;
                } else if tag_name.as_ref() == b"cellXfs" {
                    in_cell_xfs = true;
                } else if in_fills && tag_name.as_ref() == b"fill" {
                    current_fill_fg_color = None;
                } else if in_fills && tag_name.as_ref() == b"fgColor" {
                    if let Some(rgb) = attr_value(&reader, &event, b"rgb") {
                        current_fill_fg_color = Some(rgb);
                    }
                } else if in_cell_xfs && tag_name.as_ref() == b"xf" {
                    if let Some(fill_id_str) = attr_value(&reader, &event, b"fillId") {
                        if let Ok(fill_id) = fill_id_str.parse::<usize>() {
                            cell_xfs.push(fill_id);
                        } else {
                            cell_xfs.push(0);
                        }
                    } else {
                        cell_xfs.push(0);
                    }
                }
            }
            Ok(Event::Empty(event)) => {
                let tag_name = event.name();
                if in_fills && tag_name.as_ref() == b"fgColor" {
                    if let Some(rgb) = attr_value(&reader, &event, b"rgb") {
                        current_fill_fg_color = Some(rgb);
                    }
                } else if in_cell_xfs && tag_name.as_ref() == b"xf" {
                    if let Some(fill_id_str) = attr_value(&reader, &event, b"fillId") {
                        if let Ok(fill_id) = fill_id_str.parse::<usize>() {
                            cell_xfs.push(fill_id);
                        } else {
                            cell_xfs.push(0);
                        }
                    } else {
                        cell_xfs.push(0);
                    }
                } else if in_fills && tag_name.as_ref() == b"fill" {
                    fills.push(None);
                }
            }
            Ok(Event::End(event)) => {
                let tag_name = event.name();
                if tag_name.as_ref() == b"fills" {
                    in_fills = false;
                } else if tag_name.as_ref() == b"cellXfs" {
                    in_cell_xfs = false;
                } else if in_fills && tag_name.as_ref() == b"fill" {
                    fills.push(current_fill_fg_color.clone());
                    current_fill_fg_color = None;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => return None,
            _ => {}
        }
        buf.clear();
    }

    Some((fills, cell_xfs))
}

fn parse_sheet_styles(
    xml: &str,
    fills: &[Option<String>],
    cell_xfs: &[usize],
) -> Option<HashMap<String, String>> {
    let mut reader = quick_xml::Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut cell_colors = HashMap::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(event)) | Ok(Event::Empty(event)) if event.name().as_ref() == b"c" => {
                if let Some(r) = attr_value(&reader, &event, b"r") {
                    if let Some(s_str) = attr_value(&reader, &event, b"s") {
                        if let Ok(s_idx) = s_str.parse::<usize>() {
                            if s_idx < cell_xfs.len() {
                                let fill_id = cell_xfs[s_idx];
                                if fill_id < fills.len() {
                                    if let Some(ref color) = fills[fill_id] {
                                        // color is ARGB format (like FFFF0000)
                                        // We only want the last 6 chars (RRGGBB)
                                        let clean_color = if color.len() == 8 {
                                            color[2..].to_string()
                                        } else {
                                            color.clone()
                                        };
                                        cell_colors.insert(r, clean_color);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => return None,
            _ => {}
        }
        buf.clear();
    }

    Some(cell_colors)
}
