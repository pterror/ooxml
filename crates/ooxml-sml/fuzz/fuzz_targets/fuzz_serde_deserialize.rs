//! Fuzz the serde deserialization path for Worksheet.
//!
//! This target tests that malformed XML input to quick-xml's serde
//! deserialization is handled gracefully without panics.

#![no_main]

use libfuzzer_sys::fuzz_target;
use ooxml_sml::types::{Worksheet, SheetData, Row, Cell};
use quick_xml::de::from_str;

fuzz_target!(|data: &[u8]| {
    // Try to interpret as UTF-8, skip if invalid
    let Ok(s) = std::str::from_utf8(data) else {
        return;
    };

    // Test various types at different granularities
    // Each should gracefully return an error, never panic

    // Full worksheet
    let _: Result<Worksheet, _> = from_str(s);

    // Sheet data (smaller target)
    let _: Result<SheetData, _> = from_str(s);

    // Single row
    let _: Result<Row, _> = from_str(s);

    // Single cell
    let _: Result<Cell, _> = from_str(s);
});
