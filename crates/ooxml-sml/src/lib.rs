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

// Generated types from ECMA-376 schema.
// Access via `ooxml_sml::types::*` for generated structs/enums.
// This file is pre-generated and committed to avoid requiring spec downloads.
// To regenerate: OOXML_REGENERATE=1 cargo build -p ooxml-sml (with specs in /spec/)
pub mod generated;
pub use generated as types;

pub mod generated_parsers;
pub use generated_parsers as parsers;

// Extension traits for generated types (see ADR-003)
pub mod ext;
pub use ext::{
    CellExt, CellResolveExt, CellValue, Chart, ChartType, Comment, ResolveContext, ResolvedSheet,
    RowExt, SheetDataExt, WorksheetExt, parse_worksheet,
};

pub use error::{Error, Result};
// Writer-required types from workbook module
pub use workbook::{
    ConditionalRuleType, DataValidationErrorStyle, DataValidationOperator, DataValidationType,
    DefinedName, DefinedNameScope, Stylesheet, Workbook, builtin_format_code, excel_date_to_ymd,
    excel_datetime_to_ymdhms, format_excel_date, format_excel_datetime,
};
pub use writer::{
    BorderLineStyle, BorderSideStyle, BorderStyle, CellStyle, CommentBuilder, ConditionalFormat,
    ConditionalFormatRule, DataValidationBuilder, DefinedNameBuilder, FillPattern, FillStyle,
    FontStyle, HorizontalAlignment, SheetBuilder, UnderlineStyle, VerticalAlignment,
    WorkbookBuilder, WriteCellValue,
};
