//! SpreadsheetML (XLSX) support for the ooxml library.
//!
//! This crate provides reading and writing of Excel spreadsheets (.xlsx files).
//!
//! # Reading Workbooks
//!
//! ```no_run
//! use ooxml_sml::{CellResolveExt, RowExt, Workbook};
//!
//! let mut workbook = Workbook::open("spreadsheet.xlsx")?;
//! for sheet in workbook.resolved_sheets()? {
//!     println!("Sheet: {}", sheet.name());
//!     for row in sheet.rows() {
//!         for cell in row.cells_iter() {
//!             print!("{}\t", cell.value_as_string(sheet.context()));
//!         }
//!         println!();
//!     }
//! }
//! # Ok::<(), ooxml_sml::Error>(())
//! ```
//!
//! # Writing Workbooks
//!
//! ```no_run
//! use ooxml_sml::WorkbookBuilder;
//!
//! let mut wb = WorkbookBuilder::new();
//! let sheet = wb.add_sheet("Sheet1");
//! sheet.set_cell("A1", "Hello");
//! sheet.set_cell("B1", 42.0);
//! sheet.set_cell("C1", true);
//! wb.save("output.xlsx")?;
//! # Ok::<(), ooxml_sml::Error>(())
//! ```

pub mod error;
pub mod workbook;
pub mod writer;

/// Generated types from the ECMA-376 SpreadsheetML schema.
///
/// These types map 1:1 to XML elements and attributes defined in ECMA-376 Part 1 §18.
/// They are produced by `ooxml-codegen` from the RELAX NG schemas and committed to avoid
/// requiring the schema files at build time. Use the extension traits in [`ext`] for
/// ergonomic access rather than working with these types directly.
///
/// Re-exported as [`types`].
pub mod generated;
/// Type aliases for the generated ECMA-376 types. See [`generated`] for details.
pub use generated as types;

/// Generated [`FromXml`](ooxml_xml::FromXml) parsers for all generated types.
///
/// Re-exported as [`parsers`].
pub mod generated_parsers;
/// Parsers for the generated ECMA-376 types. See [`generated_parsers`] for details.
pub use generated_parsers as parsers;

/// Generated [`ToXml`](ooxml_xml::ToXml) serializers for all generated types.
///
/// Re-exported as [`serializers`].
pub mod generated_serializers;
/// Serializers for the generated ECMA-376 types. See [`generated_serializers`] for details.
pub use generated_serializers as serializers;

// Lazy/streaming API for memory-efficient parsing
pub mod lazy;
pub use lazy::{LazyCell, LazyRow, LazyWorksheet};

// Extension traits for generated types (see ADR-003)
pub mod ext;
#[cfg(feature = "sml-pivot")]
pub use ext::PivotTableExt;
pub use ext::{
    CellExt, CellResolveExt, CellValue, Chart, ChartType, Comment, ResolveContext, ResolvedSheet,
    RowExt, SheetDataExt, WorksheetExt, parse_worksheet,
};
#[cfg(feature = "sml-styling")]
pub use ext::{ConditionalFormattingExt, ConditionalRuleExt, WorksheetConditionalFormattingExt};

pub use error::{Error, Result};
// Writer-required types from workbook module
pub use workbook::{
    ConditionalRuleType, DataValidationErrorStyle, DataValidationOperator, DataValidationType,
    DefinedNameExt, DefinedNameScope, StylesheetExt, Workbook, builtin_format_code,
    excel_date_to_ymd, excel_datetime_to_ymdhms, format_excel_date, format_excel_datetime,
};
pub use writer::{
    BorderLineStyle, BorderSideStyle, BorderStyle, CellStyle, CommentBuilder, ConditionalFormat,
    ConditionalFormatRule, DataValidationBuilder, DefinedNameBuilder, FillPattern, FillStyle,
    FontStyle, HorizontalAlignment, IgnoredErrorType, PageOrientation, PageSetupOptions,
    SheetBuilder, UnderlineStyle, VerticalAlignment, WorkbookBuilder, WriteCellValue,
};
