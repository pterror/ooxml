//! Benchmark comparing serde-based parsing vs event-based parsing of SheetData.
//!
//! Run with: cargo bench -p ooxml-sml

use criterion::{BenchmarkId, Criterion, Throughput, black_box, criterion_group, criterion_main};
use ooxml_sml::types::{Cell, Row, SheetData};
use quick_xml::Reader;
use quick_xml::de::from_str;
use quick_xml::events::Event;
use std::io::Cursor;

/// Generate sample sheet data XML with the given number of rows and cells per row.
fn generate_sheet_data_xml(rows: usize, cells_per_row: usize) -> String {
    let mut xml = String::from("<sheetData>\n");
    for r in 1..=rows {
        xml.push_str(&format!(
            "  <row r=\"{}\" spans=\"1:{}\">\n",
            r, cells_per_row
        ));
        for c in 1..=cells_per_row {
            let col = col_letter(c);
            // Mix of string refs, numbers, and formulas
            match c % 3 {
                0 => {
                    xml.push_str(&format!(
                        "    <c r=\"{}{}\" t=\"s\"><v>{}</v></c>\n",
                        col,
                        r,
                        (r * cells_per_row + c) % 1000
                    ));
                }
                1 => {
                    xml.push_str(&format!(
                        "    <c r=\"{}{}\"><v>{:.2}</v></c>\n",
                        col,
                        r,
                        (r * c) as f64 / 100.0
                    ));
                }
                _ => {
                    xml.push_str(&format!(
                        "    <c r=\"{}{}\"><f>SUM(A1:A{})</f><v>{}</v></c>\n",
                        col,
                        r,
                        r,
                        r * c
                    ));
                }
            }
        }
        xml.push_str("  </row>\n");
    }
    xml.push_str("</sheetData>");
    xml
}

/// Convert column number to Excel column letter (1=A, 2=B, ..., 27=AA, etc.)
fn col_letter(col: usize) -> String {
    let mut result = String::new();
    let mut n = col;
    while n > 0 {
        n -= 1;
        result.insert(0, (b'A' + (n % 26) as u8) as char);
        n /= 26;
    }
    result
}

/// Parse sheet data using serde (generated types).
fn parse_serde(xml: &str) -> SheetData {
    from_str(xml).expect("serde parse failed")
}

/// Parse sheet data using event-based parsing (simplified version).
/// This mimics the approach used in the hand-written workbook parser.
fn parse_events(xml: &[u8]) -> (usize, usize) {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();

    let mut row_count = 0;
    let mut cell_count = 0;
    let mut in_row = false;
    let mut in_cell = false;
    let mut in_value = false;
    let mut _current_value = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let tag = e.name();
                match tag.as_ref() {
                    b"row" => {
                        in_row = true;
                        row_count += 1;
                        // Parse row attributes
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            let _key = attr.key;
                            let _val = String::from_utf8_lossy(&attr.value);
                        }
                    }
                    b"c" if in_row => {
                        in_cell = true;
                        cell_count += 1;
                        // Parse cell attributes
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            let _key = attr.key;
                            let _val = String::from_utf8_lossy(&attr.value);
                        }
                    }
                    b"v" if in_cell => {
                        in_value = true;
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                let tag = e.name();
                match tag.as_ref() {
                    b"row" => in_row = false,
                    b"c" => {
                        in_cell = false;
                        _current_value.clear();
                    }
                    b"v" => in_value = false,
                    _ => {}
                }
            }
            Ok(Event::Text(e)) if in_value => {
                _current_value = e.decode().unwrap_or_default().into_owned();
            }
            Ok(Event::Eof) => break,
            Err(e) => panic!("Error parsing: {:?}", e),
            _ => {}
        }
        buf.clear();
    }

    (row_count, cell_count)
}

fn bench_parse_sheet_data(c: &mut Criterion) {
    let mut group = c.benchmark_group("sheet_data_parsing");

    // Test with different sizes
    for (rows, cols) in [(10, 5), (100, 10), (1000, 10), (1000, 50)] {
        let xml = generate_sheet_data_xml(rows, cols);
        let xml_bytes = xml.as_bytes();
        let total_cells = rows * cols;

        group.throughput(Throughput::Elements(total_cells as u64));

        group.bench_with_input(
            BenchmarkId::new("serde", format!("{}x{}", rows, cols)),
            &xml,
            |b, xml| {
                b.iter(|| {
                    let data = parse_serde(black_box(xml));
                    black_box(data)
                })
            },
        );

        group.bench_with_input(
            BenchmarkId::new("events", format!("{}x{}", rows, cols)),
            xml_bytes,
            |b, xml| {
                b.iter(|| {
                    let counts = parse_events(black_box(xml));
                    black_box(counts)
                })
            },
        );
    }

    group.finish();
}

fn bench_parse_row(c: &mut Criterion) {
    let mut group = c.benchmark_group("row_parsing");

    let xml_small = r#"<row r="1" spans="1:5"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c><c r="D1"><v>4</v></c><c r="E1"><v>5</v></c></row>"#;
    let xml_large = generate_row_xml(100);

    group.bench_function("serde_small", |b| {
        b.iter(|| {
            let row: Row = from_str(black_box(xml_small)).unwrap();
            black_box(row)
        })
    });

    group.bench_function("serde_large", |b| {
        b.iter(|| {
            let row: Row = from_str(black_box(&xml_large)).unwrap();
            black_box(row)
        })
    });

    group.finish();
}

fn generate_row_xml(cells: usize) -> String {
    let mut xml = format!("<row r=\"1\" spans=\"1:{}\">\n", cells);
    for c in 1..=cells {
        let col = col_letter(c);
        xml.push_str(&format!("  <c r=\"{}1\"><v>{}</v></c>\n", col, c));
    }
    xml.push_str("</row>");
    xml
}

fn bench_parse_cell(c: &mut Criterion) {
    let mut group = c.benchmark_group("cell_parsing");

    let simple = r#"<c r="A1"><v>42</v></c>"#;
    let with_type = r#"<c r="A1" t="s" s="1"><v>0</v></c>"#;
    let with_formula = r#"<c r="A1"><f>SUM(B1:B100)</f><v>5050</v></c>"#;

    group.bench_function("simple", |b| {
        b.iter(|| {
            let cell: Cell = from_str(black_box(simple)).unwrap();
            black_box(cell)
        })
    });

    group.bench_function("with_type", |b| {
        b.iter(|| {
            let cell: Cell = from_str(black_box(with_type)).unwrap();
            black_box(cell)
        })
    });

    group.bench_function("with_formula", |b| {
        b.iter(|| {
            let cell: Cell = from_str(black_box(with_formula)).unwrap();
            black_box(cell)
        })
    });

    group.finish();
}

criterion_group!(
    benches,
    bench_parse_sheet_data,
    bench_parse_row,
    bench_parse_cell,
);
criterion_main!(benches);
