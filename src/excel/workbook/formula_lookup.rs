use quick_xml::events::Event;
use std::fs::File;
use std::io::{Read, Seek};
use std::path::Path;
use zip::ZipArchive;

pub(super) fn lookup_formula_in_xlsx(
    file: &Path,
    sheet_name: &str,
    cell_ref: &str,
) -> Option<String> {
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
    let target_ref = cell_ref.to_ascii_uppercase();

    let mut reader = quick_xml::Reader::from_str(&sheet_xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut current_cell = None;

    loop {
        match reader.read_event_into(&mut buf).ok()? {
            Event::Start(event) if event.name().as_ref() == b"c" => {
                current_cell = attr_value(&reader, &event, b"r")
                    .map(|reference| reference.to_ascii_uppercase());
            }
            Event::End(event) if event.name().as_ref() == b"c" => {
                current_cell = None;
            }
            Event::Start(event) if event.name().as_ref() == b"f" => {
                let mut formula = String::new();
                let end_tag = event.name().as_ref().to_vec();
                let mut inner_buf = Vec::new();
                loop {
                    match reader.read_event_into(&mut inner_buf).ok()? {
                        Event::Text(text) => formula.push_str(&text.unescape().ok()?),
                        Event::End(end_event)
                            if end_event.name().as_ref() == end_tag.as_slice() =>
                        {
                            break;
                        }
                        Event::Eof => return None,
                        _ => {}
                    }
                    inner_buf.clear();
                }

                if current_cell.as_deref() == Some(target_ref.as_str()) && !formula.is_empty() {
                    return Some(if formula.starts_with('=') {
                        formula
                    } else {
                        format!("={formula}")
                    });
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    None
}

fn read_zip_entry<R: Read + Seek>(archive: &mut ZipArchive<R>, entry_name: &str) -> Option<String> {
    let mut entry = archive.by_name(entry_name).ok()?;
    let mut contents = String::new();
    entry.read_to_string(&mut contents).ok()?;
    Some(contents)
}

fn attr_value(
    reader: &quick_xml::Reader<&[u8]>,
    event: &quick_xml::events::BytesStart<'_>,
    key: &[u8],
) -> Option<String> {
    for attr in event.attributes().flatten() {
        if attr.key.as_ref() == key {
            return attr
                .decode_and_unescape_value(reader.decoder())
                .ok()
                .map(|value| value.into_owned());
        }
    }
    None
}

fn resolve_xlsx_sheet_path<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    sheet_name: &str,
) -> Option<String> {
    let workbook_xml = read_zip_entry(archive, "xl/workbook.xml")?;
    let mut workbook_reader = quick_xml::Reader::from_str(&workbook_xml);
    workbook_reader.config_mut().trim_text(true);
    let mut workbook_buf = Vec::new();
    let mut relationship_id = None;

    loop {
        match workbook_reader.read_event_into(&mut workbook_buf).ok()? {
            Event::Start(event) | Event::Empty(event) if event.name().as_ref() == b"sheet" => {
                let name = attr_value(&workbook_reader, &event, b"name");
                if name.as_deref() == Some(sheet_name) {
                    relationship_id = attr_value(&workbook_reader, &event, b"r:id");
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        workbook_buf.clear();
    }

    let relationship_id = relationship_id?;
    let rels_xml = read_zip_entry(archive, "xl/_rels/workbook.xml.rels")?;
    let mut rels_reader = quick_xml::Reader::from_str(&rels_xml);
    rels_reader.config_mut().trim_text(true);
    let mut rels_buf = Vec::new();

    loop {
        match rels_reader.read_event_into(&mut rels_buf).ok()? {
            Event::Start(event) | Event::Empty(event)
                if event.name().as_ref() == b"Relationship" =>
            {
                let id = attr_value(&rels_reader, &event, b"Id");
                if id.as_deref() == Some(relationship_id.as_str()) {
                    let target = attr_value(&rels_reader, &event, b"Target")?;
                    return Some(if target.starts_with('/') {
                        target.trim_start_matches('/').to_string()
                    } else {
                        format!("xl/{target}")
                    });
                }
            }
            Event::Eof => break,
            _ => {}
        }
        rels_buf.clear();
    }

    None
}
