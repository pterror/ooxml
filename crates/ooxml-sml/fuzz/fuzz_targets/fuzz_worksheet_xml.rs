//! Fuzz the FromXml parser for Worksheet.
//!
//! This target tests that malformed XML input is handled gracefully
//! without panics or crashes.

#![no_main]

use libfuzzer_sys::fuzz_target;
use ooxml_sml::parsers::{FromXml, ParseError};
use ooxml_sml::types::Worksheet;
use quick_xml::Reader;
use quick_xml::events::Event;
use std::io::Cursor;

/// Parse worksheet XML from bytes using the FromXml trait.
fn parse_worksheet(xml: &[u8]) -> Result<Worksheet, ParseError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"worksheet" => {
                return Worksheet::from_xml(&mut reader, &e, false);
            }
            Ok(Event::Empty(e)) if e.name().as_ref() == b"worksheet" => {
                return Worksheet::from_xml(&mut reader, &e, true);
            }
            Ok(Event::Eof) => {
                return Err(ParseError::UnexpectedElement(
                    "EOF before worksheet".to_string(),
                ));
            }
            Err(e) => return Err(ParseError::Xml(e)),
            _ => {}
        }
    }
}

fuzz_target!(|data: &[u8]| {
    // The parser should never panic, only return errors
    let _ = parse_worksheet(data);
});
