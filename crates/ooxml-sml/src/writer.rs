//! Excel workbook writing support.
//!
//! This module provides `WorkbookBuilder` for creating new Excel files.
//!
//! # Example
//!
//! ```no_run
//! use ooxml_sml::WorkbookBuilder;
//!
//! let mut wb = WorkbookBuilder::new();
//! let sheet = wb.add_sheet("Sheet1");
//! sheet.set_cell("A1", "Hello");
//! sheet.set_cell("B1", 42.0);
//! sheet.set_cell("A2", true);
//! wb.save("output.xlsx")?;
//! # Ok::<(), ooxml_sml::Error>(())
//! ```

use crate::error::Result;
use ooxml_opc::PackageWriter;
use std::collections::HashMap;
use std::fs::File;
use std::io::{BufWriter, Seek, Write};
use std::path::Path;

// Content types
const CT_WORKBOOK: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const CT_WORKSHEET: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const CT_SHARED_STRINGS: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
const CT_RELATIONSHIPS: &str = "application/vnd.openxmlformats-package.relationships+xml";
const CT_XML: &str = "application/xml";

// Relationship types
const REL_OFFICE_DOCUMENT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_WORKSHEET: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const REL_SHARED_STRINGS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

// Namespaces
const NS_SPREADSHEET: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_RELATIONSHIPS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// A value that can be written to a cell.
#[derive(Debug, Clone)]
pub enum WriteCellValue {
    /// String value.
    String(String),
    /// Numeric value.
    Number(f64),
    /// Boolean value.
    Boolean(bool),
    /// Formula (the formula text, not the result).
    Formula(String),
    /// Empty cell.
    Empty,
}

impl From<&str> for WriteCellValue {
    fn from(s: &str) -> Self {
        WriteCellValue::String(s.to_string())
    }
}

impl From<String> for WriteCellValue {
    fn from(s: String) -> Self {
        WriteCellValue::String(s)
    }
}

impl From<f64> for WriteCellValue {
    fn from(n: f64) -> Self {
        WriteCellValue::Number(n)
    }
}

impl From<i32> for WriteCellValue {
    fn from(n: i32) -> Self {
        WriteCellValue::Number(n as f64)
    }
}

impl From<i64> for WriteCellValue {
    fn from(n: i64) -> Self {
        WriteCellValue::Number(n as f64)
    }
}

impl From<bool> for WriteCellValue {
    fn from(b: bool) -> Self {
        WriteCellValue::Boolean(b)
    }
}

/// A cell being built in a sheet.
#[derive(Debug, Clone)]
struct BuilderCell {
    value: WriteCellValue,
}

/// Column width definition for writing.
#[derive(Debug, Clone)]
struct ColumnWidth {
    min: u32,
    max: u32,
    width: f64,
}

/// Row height definition for writing.
#[derive(Debug, Clone)]
struct RowHeight {
    row: u32,
    height: f64,
}

/// A sheet being built.
#[derive(Debug)]
pub struct SheetBuilder {
    name: String,
    cells: HashMap<(u32, u32), BuilderCell>,
    merged_cells: Vec<String>,
    column_widths: Vec<ColumnWidth>,
    row_heights: Vec<RowHeight>,
}

impl SheetBuilder {
    fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            cells: HashMap::new(),
            merged_cells: Vec::new(),
            column_widths: Vec::new(),
            row_heights: Vec::new(),
        }
    }

    /// Set a cell value by reference (e.g., "A1", "B2").
    pub fn set_cell(&mut self, reference: &str, value: impl Into<WriteCellValue>) {
        if let Some((row, col)) = parse_cell_reference(reference) {
            self.cells.insert(
                (row, col),
                BuilderCell {
                    value: value.into(),
                },
            );
        }
    }

    /// Set a cell value by row and column (1-based).
    pub fn set_cell_at(&mut self, row: u32, col: u32, value: impl Into<WriteCellValue>) {
        self.cells.insert(
            (row, col),
            BuilderCell {
                value: value.into(),
            },
        );
    }

    /// Set a formula in a cell.
    pub fn set_formula(&mut self, reference: &str, formula: impl Into<String>) {
        if let Some((row, col)) = parse_cell_reference(reference) {
            self.cells.insert(
                (row, col),
                BuilderCell {
                    value: WriteCellValue::Formula(formula.into()),
                },
            );
        }
    }

    /// Merge cells in a range (e.g., "A1:B2").
    ///
    /// Note: The value of the merged cell should be set in the top-left cell.
    pub fn merge_cells(&mut self, range: &str) {
        self.merged_cells.push(range.to_string());
    }

    /// Set the width of a column (in character units, Excel default is ~8.43).
    ///
    /// Column is specified by letter (e.g., "A", "B", "AA").
    pub fn set_column_width(&mut self, col: &str, width: f64) {
        if let Some(col_num) = column_letter_to_number(col) {
            self.column_widths.push(ColumnWidth {
                min: col_num,
                max: col_num,
                width,
            });
        }
    }

    /// Set the width of a range of columns.
    ///
    /// Columns are specified by letter (e.g., "A:C" for columns A through C).
    pub fn set_column_width_range(&mut self, start_col: &str, end_col: &str, width: f64) {
        if let (Some(min), Some(max)) = (
            column_letter_to_number(start_col),
            column_letter_to_number(end_col),
        ) {
            self.column_widths.push(ColumnWidth { min, max, width });
        }
    }

    /// Set the height of a row (in points, Excel default is ~15).
    pub fn set_row_height(&mut self, row: u32, height: f64) {
        self.row_heights.push(RowHeight { row, height });
    }

    /// Get the sheet name.
    pub fn name(&self) -> &str {
        &self.name
    }
}

/// Builder for creating Excel workbooks.
#[derive(Debug)]
pub struct WorkbookBuilder {
    sheets: Vec<SheetBuilder>,
    shared_strings: Vec<String>,
    string_index: HashMap<String, usize>,
}

impl Default for WorkbookBuilder {
    fn default() -> Self {
        Self::new()
    }
}

impl WorkbookBuilder {
    /// Create a new workbook builder.
    pub fn new() -> Self {
        Self {
            sheets: Vec::new(),
            shared_strings: Vec::new(),
            string_index: HashMap::new(),
        }
    }

    /// Add a new sheet to the workbook.
    pub fn add_sheet(&mut self, name: impl Into<String>) -> &mut SheetBuilder {
        self.sheets.push(SheetBuilder::new(name));
        self.sheets.last_mut().unwrap()
    }

    /// Get a mutable reference to a sheet by index.
    pub fn sheet_mut(&mut self, index: usize) -> Option<&mut SheetBuilder> {
        self.sheets.get_mut(index)
    }

    /// Get the number of sheets.
    pub fn sheet_count(&self) -> usize {
        self.sheets.len()
    }

    /// Save the workbook to a file.
    pub fn save<P: AsRef<Path>>(self, path: P) -> Result<()> {
        let file = File::create(path)?;
        let writer = BufWriter::new(file);
        self.write(writer)
    }

    /// Write the workbook to a writer.
    pub fn write<W: Write + Seek>(mut self, writer: W) -> Result<()> {
        // Collect all strings first to build shared string table
        self.collect_shared_strings();

        let mut pkg = PackageWriter::new(writer);

        // Add default content types
        pkg.add_default_content_type("rels", CT_RELATIONSHIPS);
        pkg.add_default_content_type("xml", CT_XML);

        // Build root relationships
        let root_rels = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="{}" Target="xl/workbook.xml"/>
</Relationships>"#,
            REL_OFFICE_DOCUMENT
        );

        // Build workbook relationships
        let mut wb_rels = String::new();
        wb_rels.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        wb_rels.push('\n');
        wb_rels.push_str(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );
        wb_rels.push('\n');

        for (i, _sheet) in self.sheets.iter().enumerate() {
            let rel_id = i + 1;
            wb_rels.push_str(&format!(
                r#"  <Relationship Id="rId{}" Type="{}" Target="worksheets/sheet{}.xml"/>"#,
                rel_id, REL_WORKSHEET, rel_id
            ));
            wb_rels.push('\n');
        }

        // Add shared strings relationship if we have strings
        if !self.shared_strings.is_empty() {
            let ss_rel_id = self.sheets.len() + 1;
            wb_rels.push_str(&format!(
                r#"  <Relationship Id="rId{}" Type="{}" Target="sharedStrings.xml"/>"#,
                ss_rel_id, REL_SHARED_STRINGS
            ));
            wb_rels.push('\n');
        }

        wb_rels.push_str("</Relationships>");

        // Build workbook.xml
        let mut workbook_xml = String::new();
        workbook_xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        workbook_xml.push('\n');
        workbook_xml.push_str(&format!(
            r#"<workbook xmlns="{}" xmlns:r="{}">"#,
            NS_SPREADSHEET, NS_RELATIONSHIPS
        ));
        workbook_xml.push('\n');
        workbook_xml.push_str("  <sheets>\n");

        for (i, sheet) in self.sheets.iter().enumerate() {
            let sheet_id = i + 1;
            let rel_id = i + 1;
            workbook_xml.push_str(&format!(
                r#"    <sheet name="{}" sheetId="{}" r:id="rId{}"/>"#,
                escape_xml(&sheet.name),
                sheet_id,
                rel_id
            ));
            workbook_xml.push('\n');
        }

        workbook_xml.push_str("  </sheets>\n");
        workbook_xml.push_str("</workbook>");

        // Write parts to package
        pkg.add_part("_rels/.rels", CT_RELATIONSHIPS, root_rels.as_bytes())?;
        pkg.add_part(
            "xl/_rels/workbook.xml.rels",
            CT_RELATIONSHIPS,
            wb_rels.as_bytes(),
        )?;
        pkg.add_part("xl/workbook.xml", CT_WORKBOOK, workbook_xml.as_bytes())?;

        // Write each sheet
        for (i, sheet) in self.sheets.iter().enumerate() {
            let sheet_num = i + 1;
            let sheet_xml = self.serialize_sheet(sheet);
            let part_name = format!("xl/worksheets/sheet{}.xml", sheet_num);
            pkg.add_part(&part_name, CT_WORKSHEET, sheet_xml.as_bytes())?;
        }

        // Write shared strings if any
        if !self.shared_strings.is_empty() {
            let ss_xml = self.serialize_shared_strings();
            pkg.add_part("xl/sharedStrings.xml", CT_SHARED_STRINGS, ss_xml.as_bytes())?;
        }

        pkg.finish()?;
        Ok(())
    }

    /// Collect all strings from cells into shared string table.
    fn collect_shared_strings(&mut self) {
        for sheet in &self.sheets {
            for cell in sheet.cells.values() {
                if let WriteCellValue::String(s) = &cell.value
                    && !self.string_index.contains_key(s)
                {
                    let idx = self.shared_strings.len();
                    self.shared_strings.push(s.clone());
                    self.string_index.insert(s.clone(), idx);
                }
            }
        }
    }

    /// Serialize a sheet to XML.
    fn serialize_sheet(&self, sheet: &SheetBuilder) -> String {
        let mut xml = String::new();
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push('\n');
        xml.push_str(&format!(r#"<worksheet xmlns="{}">"#, NS_SPREADSHEET));
        xml.push('\n');

        // Write column widths if any
        if !sheet.column_widths.is_empty() {
            xml.push_str("  <cols>\n");
            for col in &sheet.column_widths {
                xml.push_str(&format!(
                    r#"    <col min="{}" max="{}" width="{}" customWidth="1"/>"#,
                    col.min, col.max, col.width
                ));
                xml.push('\n');
            }
            xml.push_str("  </cols>\n");
        }

        xml.push_str("  <sheetData>\n");

        // Group cells by row
        let mut rows: HashMap<u32, Vec<(u32, &BuilderCell)>> = HashMap::new();
        for ((row, col), cell) in &sheet.cells {
            rows.entry(*row).or_default().push((*col, cell));
        }

        // Build row height lookup
        let row_height_map: HashMap<u32, f64> = sheet
            .row_heights
            .iter()
            .map(|rh| (rh.row, rh.height))
            .collect();

        // Sort rows and serialize
        let mut row_nums: Vec<_> = rows.keys().copied().collect();
        row_nums.sort();

        for row_num in row_nums {
            let cells = rows.get(&row_num).unwrap();
            let mut sorted_cells: Vec<_> = cells.clone();
            sorted_cells.sort_by_key(|(col, _)| *col);

            // Include row height if set
            if let Some(height) = row_height_map.get(&row_num) {
                xml.push_str(&format!(
                    r#"    <row r="{}" ht="{}" customHeight="1">"#,
                    row_num, height
                ));
            } else {
                xml.push_str(&format!(r#"    <row r="{}">"#, row_num));
            }
            xml.push('\n');

            for (col, cell) in sorted_cells {
                let ref_str = column_to_letter(col) + &row_num.to_string();
                xml.push_str(&self.serialize_cell(&ref_str, cell));
            }

            xml.push_str("    </row>\n");
        }

        xml.push_str("  </sheetData>\n");

        // Write merged cells if any
        if !sheet.merged_cells.is_empty() {
            xml.push_str(&format!(
                "  <mergeCells count=\"{}\">\n",
                sheet.merged_cells.len()
            ));
            for range in &sheet.merged_cells {
                xml.push_str(&format!(r#"    <mergeCell ref="{}"/>"#, escape_xml(range)));
                xml.push('\n');
            }
            xml.push_str("  </mergeCells>\n");
        }

        xml.push_str("</worksheet>");
        xml
    }

    /// Serialize a cell to XML.
    fn serialize_cell(&self, reference: &str, cell: &BuilderCell) -> String {
        match &cell.value {
            WriteCellValue::String(s) => {
                let idx = self.string_index.get(s).unwrap_or(&0);
                format!(r#"      <c r="{}" t="s"><v>{}</v></c>"#, reference, idx) + "\n"
            }
            WriteCellValue::Number(n) => {
                format!(r#"      <c r="{}"><v>{}</v></c>"#, reference, n) + "\n"
            }
            WriteCellValue::Boolean(b) => {
                let val = if *b { "1" } else { "0" };
                format!(r#"      <c r="{}" t="b"><v>{}</v></c>"#, reference, val) + "\n"
            }
            WriteCellValue::Formula(f) => {
                format!(r#"      <c r="{}"><f>{}</f></c>"#, reference, escape_xml(f)) + "\n"
            }
            WriteCellValue::Empty => String::new(),
        }
    }

    /// Serialize shared strings table to XML.
    fn serialize_shared_strings(&self) -> String {
        let mut xml = String::new();
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push('\n');
        xml.push_str(&format!(
            r#"<sst xmlns="{}" count="{}" uniqueCount="{}">"#,
            NS_SPREADSHEET,
            self.shared_strings.len(),
            self.shared_strings.len()
        ));
        xml.push('\n');

        for s in &self.shared_strings {
            xml.push_str(&format!("  <si><t>{}</t></si>\n", escape_xml(s)));
        }

        xml.push_str("</sst>");
        xml
    }
}

/// Parse a cell reference like "A1" into (row, col).
fn parse_cell_reference(reference: &str) -> Option<(u32, u32)> {
    let mut col_part = String::new();
    let mut row_part = String::new();

    for c in reference.chars() {
        if c.is_ascii_alphabetic() {
            col_part.push(c.to_ascii_uppercase());
        } else if c.is_ascii_digit() {
            row_part.push(c);
        }
    }

    let col = column_letter_to_number(&col_part)?;
    let row: u32 = row_part.parse().ok()?;

    Some((row, col))
}

/// Convert column letters to number (A=1, B=2, ..., Z=26, AA=27).
fn column_letter_to_number(letters: &str) -> Option<u32> {
    if letters.is_empty() {
        return None;
    }

    let mut result: u32 = 0;
    for c in letters.chars() {
        if !c.is_ascii_alphabetic() {
            return None;
        }
        result = result * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
    }
    Some(result)
}

/// Convert column number to letters (1=A, 2=B, ..., 26=Z, 27=AA).
fn column_to_letter(mut col: u32) -> String {
    let mut result = String::new();
    while col > 0 {
        col -= 1;
        result.insert(0, (b'A' + (col % 26) as u8) as char);
        col /= 26;
    }
    result
}

/// Escape XML special characters.
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_column_letter_to_number() {
        assert_eq!(column_letter_to_number("A"), Some(1));
        assert_eq!(column_letter_to_number("B"), Some(2));
        assert_eq!(column_letter_to_number("Z"), Some(26));
        assert_eq!(column_letter_to_number("AA"), Some(27));
        assert_eq!(column_letter_to_number("AB"), Some(28));
        assert_eq!(column_letter_to_number("AZ"), Some(52));
    }

    #[test]
    fn test_column_to_letter() {
        assert_eq!(column_to_letter(1), "A");
        assert_eq!(column_to_letter(2), "B");
        assert_eq!(column_to_letter(26), "Z");
        assert_eq!(column_to_letter(27), "AA");
        assert_eq!(column_to_letter(28), "AB");
        assert_eq!(column_to_letter(52), "AZ");
    }

    #[test]
    fn test_parse_cell_reference() {
        assert_eq!(parse_cell_reference("A1"), Some((1, 1)));
        assert_eq!(parse_cell_reference("B2"), Some((2, 2)));
        assert_eq!(parse_cell_reference("Z10"), Some((10, 26)));
        assert_eq!(parse_cell_reference("AA1"), Some((1, 27)));
    }

    #[test]
    fn test_workbook_builder() {
        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Test");
        sheet.set_cell("A1", "Hello");
        sheet.set_cell("B1", 42.0);
        sheet.set_cell("C1", true);
        sheet.set_formula("D1", "SUM(A1:C1)");

        assert_eq!(wb.sheet_count(), 1);
    }

    #[test]
    fn test_roundtrip_simple() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Sheet1");
        sheet.set_cell("A1", "Test Value");
        sheet.set_cell("B1", 123.45);

        // Write to memory
        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        // Read back
        buffer.set_position(0);
        let mut workbook = crate::Workbook::from_reader(buffer).unwrap();
        let read_sheet = workbook.sheet(0).unwrap();

        assert_eq!(read_sheet.name(), "Sheet1");

        let cell_a1 = read_sheet.cell("A1").unwrap();
        assert_eq!(cell_a1.value_as_string(), "Test Value");

        let cell_b1 = read_sheet.cell("B1").unwrap();
        assert_eq!(cell_b1.value_as_number(), Some(123.45));
    }

    #[test]
    fn test_roundtrip_merged_cells() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Sheet1");
        sheet.set_cell("A1", "Merged Header");
        sheet.merge_cells("A1:C1");
        sheet.set_cell("A2", "Data 1");
        sheet.set_cell("B2", "Data 2");
        sheet.merge_cells("A3:B4");
        sheet.set_cell("A3", "Block");

        // Write to memory
        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        // Read back
        buffer.set_position(0);
        let mut workbook = crate::Workbook::from_reader(buffer).unwrap();
        let read_sheet = workbook.sheet(0).unwrap();

        // Check merged cells were preserved
        let merged = read_sheet.merged_cells();
        assert_eq!(merged.len(), 2);
        assert_eq!(merged[0].reference(), "A1:C1");
        assert_eq!(merged[1].reference(), "A3:B4");

        // Check helper methods
        assert!(merged[0].contains(1, 1)); // A1
        assert!(merged[0].contains(1, 2)); // B1
        assert!(merged[0].contains(1, 3)); // C1
        assert!(!merged[0].contains(1, 4)); // D1 - outside

        // Check start/end cells
        assert_eq!(merged[0].start_cell(), "A1");
        assert_eq!(merged[0].end_cell(), "C1");
        assert_eq!(merged[1].start_cell(), "A3");
        assert_eq!(merged[1].end_cell(), "B4");

        // Check merged_cell_at helper
        assert!(read_sheet.merged_cell_at("A1").is_some());
        assert!(read_sheet.merged_cell_at("B1").is_some());
        assert!(read_sheet.merged_cell_at("D1").is_none());

        // Cell values should still be accessible
        let cell_a1 = read_sheet.cell("A1").unwrap();
        assert_eq!(cell_a1.value_as_string(), "Merged Header");
    }

    #[test]
    fn test_roundtrip_dimensions() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Sheet1");
        sheet.set_cell("A1", "Header 1");
        sheet.set_cell("B1", "Header 2");
        sheet.set_cell("C1", "Header 3");
        sheet.set_cell("A2", "Data");

        // Set column widths
        sheet.set_column_width("A", 20.0);
        sheet.set_column_width_range("B", "C", 15.5);

        // Set row heights
        sheet.set_row_height(1, 25.0);
        sheet.set_row_height(2, 18.0);

        // Write to memory
        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        // Read back
        buffer.set_position(0);
        let mut workbook = crate::Workbook::from_reader(buffer).unwrap();
        let read_sheet = workbook.sheet(0).unwrap();

        // Check column widths were preserved
        let columns = read_sheet.columns();
        assert_eq!(columns.len(), 2);

        // Column A (col 1)
        assert_eq!(columns[0].min, 1);
        assert_eq!(columns[0].max, 1);
        assert_eq!(columns[0].width, Some(20.0));

        // Columns B-C (cols 2-3)
        assert_eq!(columns[1].min, 2);
        assert_eq!(columns[1].max, 3);
        assert_eq!(columns[1].width, Some(15.5));

        // Test helper methods
        assert_eq!(read_sheet.column_width(1), Some(20.0));
        assert_eq!(read_sheet.column_width(2), Some(15.5));
        assert_eq!(read_sheet.column_width(3), Some(15.5));
        assert_eq!(read_sheet.column_width(4), None);

        // Check row heights were preserved
        let row1 = read_sheet.row(1).unwrap();
        assert_eq!(row1.height(), Some(25.0));

        let row2 = read_sheet.row(2).unwrap();
        assert_eq!(row2.height(), Some(18.0));
    }
}
