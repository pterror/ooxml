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
    AutoFilter, Border, BorderSide, Cell, CellFormat, CellValue, Chart, ChartSeries, ChartType,
    ColorFilter, ColorScale, ColorScaleValue, ColumnInfo, Comment, ConditionalFormatting,
    ConditionalRule, ConditionalRuleType, CustomFilter, DataBar, DataValidation,
    DataValidationErrorStyle, DataValidationOperator, DataValidationType, DefinedName,
    DefinedNameScope, Fill, FilterColumn, FilterOperator, Font, FreezePane, IconSet, IconSetValue,
    MergedCell, NumberFormat, PanePosition, Row, Sheet, Stylesheet, Top10Filter, Workbook,
    builtin_format_code, excel_date_to_ymd, excel_datetime_to_ymdhms, format_excel_date,
    format_excel_datetime,
};
pub use writer::{
    BorderLineStyle, BorderSideStyle, BorderStyle, CellStyle, FillPattern, FillStyle, FontStyle,
    HorizontalAlignment, SheetBuilder, UnderlineStyle, VerticalAlignment, WorkbookBuilder,
    WriteCellValue,
};
