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

pub mod error;
pub mod workbook;

pub use error::{Error, Result};
pub use workbook::{Cell, CellValue, Row, Sheet, Workbook};
