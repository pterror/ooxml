//! SpreadsheetML (XLSX) support for the ooxml library.
//!
//! This crate provides reading and writing of Excel spreadsheets (.xlsx files).
//!
//! # Reading Workbooks
//!
//! ```no_run
//! use ooxml_sml::Workbook;
//!
//! let mut workbook = Workbook::open("spreadsheet.xlsx")?;
//! for sheet in workbook.sheets()? {
//!     println!("Sheet: {}", sheet.name());
//!     for row in sheet.rows() {
//!         for cell in row.cells() {
//!             print!("{}\t", cell.value_as_string());
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

pub use error::{Error, Result};
pub use workbook::{
    Border, BorderSide, Cell, CellFormat, CellValue, ColumnInfo, Comment, Fill, Font, MergedCell,
    NumberFormat, Row, Sheet, Stylesheet, Workbook,
};
pub use writer::{SheetBuilder, WorkbookBuilder, WriteCellValue};
