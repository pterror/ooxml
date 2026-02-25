// Custom handwritten SML fixtures.

use ooxml_sml::{
    BorderLineStyle, BorderStyle, CellStyle, CommentBuilder, DataValidationBuilder,
    DataValidationOperator, FillStyle, FontStyle, IgnoredErrorType, PageOrientation,
    PageSetupOptions, WorkbookBuilder, WriteCellValue,
};

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn write_wb(wb: WorkbookBuilder) -> Vec<u8> {
    let mut buf = std::io::Cursor::new(Vec::new());
    wb.write(&mut buf).expect("write failed");
    buf.into_inner()
}

// ---------------------------------------------------------------------------
// sml/cell — special string values
// ---------------------------------------------------------------------------

pub fn fixture_sml_cell_string_unicode() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", "café 中文 🌍");
    crate::Fixture {
        path: "sml/cell/string-unicode.xlsx",
        description: "Unicode string cell value",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::SheetCount { expected: 1 },
            crate::Assertion::CellType {
                sheet: 0,
                row: 0,
                col: 0,
                expected: "string".into(),
            },
            crate::Assertion::CellValue {
                sheet: 0,
                row: 0,
                col: 0,
                expected: "café 中文 🌍".into(),
                tolerance: 0.0,
            },
        ],
    }
}

pub fn fixture_sml_cell_string_xml_special() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", "& < > \" '");
    crate::Fixture {
        path: "sml/cell/string-xml-special.xlsx",
        description: "XML special characters in string cell",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::SheetCount { expected: 1 },
            crate::Assertion::CellType {
                sheet: 0,
                row: 0,
                col: 0,
                expected: "string".into(),
            },
            crate::Assertion::CellValue {
                sheet: 0,
                row: 0,
                col: 0,
                expected: "& < > \" '".into(),
                tolerance: 0.0,
            },
        ],
    }
}

// ---------------------------------------------------------------------------
// sml/cell/formula — formula cells
// ---------------------------------------------------------------------------

pub fn fixture_sml_cell_formula_sum() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", WriteCellValue::Number(1.0));
    sheet.set_cell("A2", WriteCellValue::Number(2.0));
    sheet.set_cell("A3", WriteCellValue::Number(3.0));
    sheet.set_formula("B1", "SUM(A1:A3)");
    crate::Fixture {
        path: "sml/cell/formula/sum.xlsx",
        description: "SUM formula over A1:A3",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::SheetCount { expected: 1 },
            crate::Assertion::CellFormula {
                sheet: 0,
                row: 0,
                col: 1,
                expected: Some("SUM(A1:A3)".into()),
            },
        ],
    }
}

pub fn fixture_sml_cell_formula_if() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", WriteCellValue::Number(5.0));
    sheet.set_formula("B1", "IF(A1>0,\"pos\",\"neg\")");
    crate::Fixture {
        path: "sml/cell/formula/if.xlsx",
        description: "IF formula",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::SheetCount { expected: 1 },
            crate::Assertion::CellFormula {
                sheet: 0,
                row: 0,
                col: 1,
                expected: Some("IF(A1>0,\"pos\",\"neg\")".into()),
            },
        ],
    }
}

pub fn fixture_sml_cell_formula_concat() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", "Hello");
    sheet.set_cell("B1", "World");
    sheet.set_formula("C1", "A1&\" \"&B1");
    crate::Fixture {
        path: "sml/cell/formula/concat.xlsx",
        description: "String concatenation formula",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::SheetCount { expected: 1 },
            crate::Assertion::CellFormula {
                sheet: 0,
                row: 0,
                col: 2,
                expected: Some("A1&\" \"&B1".into()),
            },
        ],
    }
}

pub fn fixture_sml_cell_formula_reference() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", WriteCellValue::Number(10.0));
    sheet.set_cell("B1", WriteCellValue::Number(20.0));
    sheet.set_formula("C1", "A1+B1");
    crate::Fixture {
        path: "sml/cell/formula/reference.xlsx",
        description: "Cell reference formula",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::SheetCount { expected: 1 },
            crate::Assertion::CellFormula {
                sheet: 0,
                row: 0,
                col: 2,
                expected: Some("A1+B1".into()),
            },
        ],
    }
}

// ---------------------------------------------------------------------------
// sml/format — number formatting
// ---------------------------------------------------------------------------

pub fn fixture_sml_format_integer() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    let style = CellStyle::new().with_number_format("0");
    sheet.set_cell_styled("A1", WriteCellValue::Number(1234.0), style);
    crate::Fixture {
        path: "sml/format/integer.xlsx",
        description: "Integer number format",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::SheetCount { expected: 1 },
            crate::Assertion::CellValue {
                sheet: 0,
                row: 0,
                col: 0,
                expected: "1234".into(),
                tolerance: 0.0,
            },
            crate::Assertion::CellFormatCode {
                sheet: 0,
                row: 0,
                col: 0,
                expected: Some("0".into()),
            },
        ],
    }
}

pub fn fixture_sml_format_decimal() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    let style = CellStyle::new().with_number_format("#,##0.00");
    sheet.set_cell_styled("A1", WriteCellValue::Number(1234.5), style);
    crate::Fixture {
        path: "sml/format/decimal.xlsx",
        description: "Decimal number format with thousands separator",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::SheetCount { expected: 1 },
            crate::Assertion::CellValue {
                sheet: 0,
                row: 0,
                col: 0,
                expected: "1234.5".into(),
                tolerance: 1.0e-9,
            },
            crate::Assertion::CellFormatCode {
                sheet: 0,
                row: 0,
                col: 0,
                expected: Some("#,##0.00".into()),
            },
        ],
    }
}

pub fn fixture_sml_format_percent() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    let style = CellStyle::new().with_number_format("0%");
    sheet.set_cell_styled("A1", WriteCellValue::Number(0.42), style);
    crate::Fixture {
        path: "sml/format/percent.xlsx",
        description: "Percentage number format",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::SheetCount { expected: 1 },
            crate::Assertion::CellValue {
                sheet: 0,
                row: 0,
                col: 0,
                expected: "0.42".into(),
                tolerance: 1.0e-9,
            },
            crate::Assertion::CellFormatCode {
                sheet: 0,
                row: 0,
                col: 0,
                expected: Some("0%".into()),
            },
        ],
    }
}

pub fn fixture_sml_format_currency() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    let style = CellStyle::new().with_number_format("$#,##0.00");
    sheet.set_cell_styled("A1", WriteCellValue::Number(9.99), style);
    crate::Fixture {
        path: "sml/format/currency.xlsx",
        description: "Currency number format",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::SheetCount { expected: 1 },
            crate::Assertion::CellValue {
                sheet: 0,
                row: 0,
                col: 0,
                expected: "9.99".into(),
                tolerance: 1.0e-9,
            },
            crate::Assertion::CellFormatCode {
                sheet: 0,
                row: 0,
                col: 0,
                expected: Some("$#,##0.00".into()),
            },
        ],
    }
}

pub fn fixture_sml_format_date() -> crate::Fixture {
    // 44927.0 = 2023-01-01 in Excel serial date format
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    let style = CellStyle::new().with_number_format("MM/DD/YYYY");
    sheet.set_cell_styled("A1", WriteCellValue::Number(44927.0), style);
    crate::Fixture {
        path: "sml/format/date.xlsx",
        description: "Date number format",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::SheetCount { expected: 1 },
            crate::Assertion::CellValue {
                sheet: 0,
                row: 0,
                col: 0,
                expected: "44927".into(),
                tolerance: 0.0,
            },
            crate::Assertion::CellFormatCode {
                sheet: 0,
                row: 0,
                col: 0,
                expected: Some("MM/DD/YYYY".into()),
            },
        ],
    }
}

pub fn fixture_sml_format_scientific() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    let style = CellStyle::new().with_number_format("0.00E+00");
    sheet.set_cell_styled("A1", WriteCellValue::Number(0.00001), style);
    crate::Fixture {
        path: "sml/format/scientific.xlsx",
        description: "Scientific notation number format",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::SheetCount { expected: 1 },
            crate::Assertion::CellValue {
                sheet: 0,
                row: 0,
                col: 0,
                expected: "0.00001".into(),
                tolerance: 1.0e-10,
            },
            crate::Assertion::CellFormatCode {
                sheet: 0,
                row: 0,
                col: 0,
                expected: Some("0.00E+00".into()),
            },
        ],
    }
}

// ---------------------------------------------------------------------------
// sml/format/font — font formatting
// ---------------------------------------------------------------------------

pub fn fixture_sml_format_font_bold() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    let style = CellStyle::new().with_font(FontStyle::new().bold());
    sheet.set_cell_styled("A1", WriteCellValue::Number(42.0), style);
    crate::Fixture {
        path: "sml/format/font/bold.xlsx",
        description: "Bold font cell",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::SheetCount { expected: 1 },
            crate::Assertion::CellValue {
                sheet: 0,
                row: 0,
                col: 0,
                expected: "42".into(),
                tolerance: 0.0,
            },
            crate::Assertion::CellBold {
                sheet: 0,
                row: 0,
                col: 0,
                expected: true,
            },
        ],
    }
}

pub fn fixture_sml_format_font_italic() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    let style = CellStyle::new().with_font(FontStyle::new().italic());
    sheet.set_cell_styled("A1", WriteCellValue::Number(42.0), style);
    crate::Fixture {
        path: "sml/format/font/italic.xlsx",
        description: "Italic font cell",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::SheetCount { expected: 1 },
            crate::Assertion::CellValue {
                sheet: 0,
                row: 0,
                col: 0,
                expected: "42".into(),
                tolerance: 0.0,
            },
            crate::Assertion::CellItalic {
                sheet: 0,
                row: 0,
                col: 0,
                expected: true,
            },
        ],
    }
}

pub fn fixture_sml_format_font_bold_italic() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    let style = CellStyle::new().with_font(FontStyle::new().bold().italic());
    sheet.set_cell_styled("A1", WriteCellValue::Number(42.0), style);
    crate::Fixture {
        path: "sml/format/font/bold-italic.xlsx",
        description: "Bold and italic font cell",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::SheetCount { expected: 1 },
            crate::Assertion::CellBold {
                sheet: 0,
                row: 0,
                col: 0,
                expected: true,
            },
            crate::Assertion::CellItalic {
                sheet: 0,
                row: 0,
                col: 0,
                expected: true,
            },
        ],
    }
}

pub fn fixture_sml_format_font_color() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    let style = CellStyle::new().with_font(FontStyle::new().with_color("FF0000"));
    sheet.set_cell_styled("A1", WriteCellValue::String("Red text".into()), style);
    crate::Fixture {
        path: "sml/format/font/color.xlsx",
        description: "Colored font cell",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::SheetCount { expected: 1 },
            crate::Assertion::CellValue {
                sheet: 0,
                row: 0,
                col: 0,
                expected: "Red text".into(),
                tolerance: 0.0,
            },
        ],
    }
}

pub fn fixture_sml_format_font_size() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    let style = CellStyle::new().with_font(FontStyle::new().with_size(14.0));
    sheet.set_cell_styled("A1", WriteCellValue::String("Big text".into()), style);
    crate::Fixture {
        path: "sml/format/font/size.xlsx",
        description: "Font size 14pt cell",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::SheetCount { expected: 1 },
            crate::Assertion::CellValue {
                sheet: 0,
                row: 0,
                col: 0,
                expected: "Big text".into(),
                tolerance: 0.0,
            },
        ],
    }
}

// ---------------------------------------------------------------------------
// sml/format/fill — fill formatting
// ---------------------------------------------------------------------------

pub fn fixture_sml_format_fill_solid_red() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    let style = CellStyle::new().with_fill(FillStyle::solid("FF0000"));
    sheet.set_cell_styled("A1", WriteCellValue::String("Red".into()), style);
    crate::Fixture {
        path: "sml/format/fill/solid-red.xlsx",
        description: "Solid red fill cell",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::SheetCount { expected: 1 },
            crate::Assertion::CellValue {
                sheet: 0,
                row: 0,
                col: 0,
                expected: "Red".into(),
                tolerance: 0.0,
            },
            crate::Assertion::CellColor {
                sheet: 0,
                row: 0,
                col: 0,
                expected: Some("FF0000".into()),
            },
        ],
    }
}

pub fn fixture_sml_format_fill_solid_yellow() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    let style = CellStyle::new().with_fill(FillStyle::solid("FFFF00"));
    sheet.set_cell_styled("A1", WriteCellValue::String("Yellow".into()), style);
    crate::Fixture {
        path: "sml/format/fill/solid-yellow.xlsx",
        description: "Solid yellow fill cell",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::SheetCount { expected: 1 },
            crate::Assertion::CellValue {
                sheet: 0,
                row: 0,
                col: 0,
                expected: "Yellow".into(),
                tolerance: 0.0,
            },
            crate::Assertion::CellColor {
                sheet: 0,
                row: 0,
                col: 0,
                expected: Some("FFFF00".into()),
            },
        ],
    }
}

// ---------------------------------------------------------------------------
// sml/format/border — border formatting
// ---------------------------------------------------------------------------

pub fn fixture_sml_format_border_all() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    let border = BorderStyle::all(BorderLineStyle::Thin, Some("000000".into()));
    let style = CellStyle::new().with_border(border);
    sheet.set_cell_styled("A1", WriteCellValue::String("Bordered".into()), style);
    crate::Fixture {
        path: "sml/format/border/all.xlsx",
        description: "All borders thin black",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::SheetCount { expected: 1 },
            crate::Assertion::CellValue {
                sheet: 0,
                row: 0,
                col: 0,
                expected: "Bordered".into(),
                tolerance: 0.0,
            },
        ],
    }
}

pub fn fixture_sml_format_border_outline() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    let border = BorderStyle::new()
        .with_left(BorderLineStyle::Thin, Some("000000".into()))
        .with_right(BorderLineStyle::Thin, Some("000000".into()))
        .with_top(BorderLineStyle::Thin, Some("000000".into()))
        .with_bottom(BorderLineStyle::Thin, Some("000000".into()));
    let style = CellStyle::new().with_border(border);
    sheet.set_cell_styled("A1", WriteCellValue::String("Outline".into()), style);
    crate::Fixture {
        path: "sml/format/border/outline.xlsx",
        description: "Outer borders only",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::SheetCount { expected: 1 },
            crate::Assertion::CellValue {
                sheet: 0,
                row: 0,
                col: 0,
                expected: "Outline".into(),
                tolerance: 0.0,
            },
        ],
    }
}

// ---------------------------------------------------------------------------
// sml/layout — sheet layout
// ---------------------------------------------------------------------------

pub fn fixture_sml_layout_freeze_pane() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", "Header");
    sheet.set_cell("A2", "Data");
    sheet.set_freeze_pane(1, 0);
    crate::Fixture {
        path: "sml/layout/freeze-pane.xlsx",
        description: "Freeze pane at row 1",
        bytes: write_wb(wb),
        assertions: vec![crate::Assertion::SheetCount { expected: 1 }],
    }
}

pub fn fixture_sml_layout_auto_filter() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", "Name");
    sheet.set_cell("B1", "Age");
    sheet.set_cell("C1", "City");
    sheet.set_cell("D1", "Score");
    sheet.set_auto_filter("A1:D1");
    crate::Fixture {
        path: "sml/layout/auto-filter.xlsx",
        description: "Auto-filter on header row A1:D1",
        bytes: write_wb(wb),
        assertions: vec![crate::Assertion::SheetCount { expected: 1 }],
    }
}

pub fn fixture_sml_layout_page_setup_landscape() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", "Landscape");
    sheet.set_page_setup(PageSetupOptions::new().with_orientation(PageOrientation::Landscape));
    crate::Fixture {
        path: "sml/layout/page-setup-landscape.xlsx",
        description: "Landscape page orientation",
        bytes: write_wb(wb),
        assertions: vec![crate::Assertion::SheetCount { expected: 1 }],
    }
}

pub fn fixture_sml_layout_page_margins() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", "Margins");
    // left, right, top, bottom, header, footer (in inches)
    sheet.set_page_margins(0.5, 0.5, 0.75, 0.75, 0.3, 0.3);
    crate::Fixture {
        path: "sml/layout/page-margins.xlsx",
        description: "Custom page margins",
        bytes: write_wb(wb),
        assertions: vec![crate::Assertion::SheetCount { expected: 1 }],
    }
}

pub fn fixture_sml_layout_row_height() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", "Tall row");
    sheet.set_row_height(1, 30.0);
    crate::Fixture {
        path: "sml/layout/row-height.xlsx",
        description: "Custom row height 30pt",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::SheetCount { expected: 1 },
            crate::Assertion::RowHeight {
                sheet: 0,
                row: 0,
                expected: 30.0,
                tolerance: 0.5,
            },
        ],
    }
}

pub fn fixture_sml_layout_col_width() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", "Wide col");
    sheet.set_column_width("A", 20.0);
    crate::Fixture {
        path: "sml/layout/col-width.xlsx",
        description: "Custom column width 20 character units",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::SheetCount { expected: 1 },
            crate::Assertion::ColWidth {
                sheet: 0,
                col: 0,
                expected: 20.0,
                tolerance: 0.5,
            },
        ],
    }
}

pub fn fixture_sml_layout_row_outline() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", "Header");
    sheet.set_cell("A2", "Detail");
    // Row 2 (1-based) gets outline level 1
    sheet.set_row_outline_level(2, 1);
    crate::Fixture {
        path: "sml/layout/row-outline.xlsx",
        description: "Row with outline level 1",
        bytes: write_wb(wb),
        assertions: vec![crate::Assertion::SheetCount { expected: 1 }],
    }
}

pub fn fixture_sml_layout_col_outline() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", "Main");
    sheet.set_cell("B1", "Detail");
    // Column B gets outline level 1
    sheet.set_column_outline_level("B", 1);
    crate::Fixture {
        path: "sml/layout/col-outline.xlsx",
        description: "Column with outline level 1",
        bytes: write_wb(wb),
        assertions: vec![crate::Assertion::SheetCount { expected: 1 }],
    }
}

pub fn fixture_sml_layout_tab_color() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", "Colored tab");
    sheet.set_tab_color("FF0000");
    crate::Fixture {
        path: "sml/layout/tab-color.xlsx",
        description: "Sheet tab color red",
        bytes: write_wb(wb),
        assertions: vec![crate::Assertion::SheetCount { expected: 1 }],
    }
}

pub fn fixture_sml_layout_hidden_row() -> crate::Fixture {
    // No direct set_row_hidden API; use set_row_collapsed to demonstrate
    // that row visibility state can be set.
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", "Visible");
    sheet.set_cell("A2", "Collapsed");
    sheet.set_row_collapsed(2, true);
    crate::Fixture {
        path: "sml/layout/hidden-row.xlsx",
        description: "Sheet with a collapsed/hidden row",
        bytes: write_wb(wb),
        assertions: vec![crate::Assertion::SheetCount { expected: 1 }],
    }
}

pub fn fixture_sml_layout_hidden_col() -> crate::Fixture {
    // No direct set_col_hidden API; use set_column_collapsed to demonstrate
    // that column visibility state can be set.
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", "Visible");
    sheet.set_cell("B1", "Collapsed");
    sheet.set_column_collapsed("B", true);
    crate::Fixture {
        path: "sml/layout/hidden-col.xlsx",
        description: "Sheet with a collapsed/hidden column",
        bytes: write_wb(wb),
        assertions: vec![crate::Assertion::SheetCount { expected: 1 }],
    }
}

// ---------------------------------------------------------------------------
// sml/merge — merged cells
// ---------------------------------------------------------------------------

pub fn fixture_sml_merge_horizontal() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", "Merged");
    sheet.merge_cells("A1:B1");
    crate::Fixture {
        path: "sml/merge/horizontal.xlsx",
        description: "Horizontally merged cells A1:B1",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::SheetCount { expected: 1 },
            crate::Assertion::MergedRegion {
                sheet: 0,
                row: 0,
                col: 0,
                expected: true,
            },
        ],
    }
}

pub fn fixture_sml_merge_2x2() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", "Merged 2x2");
    sheet.merge_cells("A1:B2");
    crate::Fixture {
        path: "sml/merge/2x2.xlsx",
        description: "2x2 merged cell block A1:B2",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::SheetCount { expected: 1 },
            crate::Assertion::MergedRegion {
                sheet: 0,
                row: 0,
                col: 0,
                expected: true,
            },
        ],
    }
}

// ---------------------------------------------------------------------------
// sml/hyperlink — hyperlinks
// ---------------------------------------------------------------------------

pub fn fixture_sml_hyperlink_external() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", "Example");
    sheet.add_hyperlink("A1", "https://example.com");
    crate::Fixture {
        path: "sml/hyperlink/external.xlsx",
        description: "External hyperlink on cell A1",
        bytes: write_wb(wb),
        assertions: vec![crate::Assertion::SheetCount { expected: 1 }],
    }
}

pub fn fixture_sml_hyperlink_internal() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    wb.add_sheet("Sheet1");
    wb.add_sheet("Sheet2");
    // Get sheet 0 and add a hyperlink to Sheet2!A1
    if let Some(sheet) = wb.sheet_mut(0) {
        sheet.set_cell("A1", "Go to Sheet2");
        sheet.add_internal_hyperlink("A1", "Sheet2!A1");
    }
    crate::Fixture {
        path: "sml/hyperlink/internal.xlsx",
        description: "Internal hyperlink to Sheet2!A1",
        bytes: write_wb(wb),
        assertions: vec![crate::Assertion::SheetCount { expected: 2 }],
    }
}

// ---------------------------------------------------------------------------
// sml/comment — cell comments
// ---------------------------------------------------------------------------

pub fn fixture_sml_comment_plain() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", "Cell with comment");
    sheet.add_comment("A1", "This is a plain comment");
    crate::Fixture {
        path: "sml/comment/plain.xlsx",
        description: "Plain text comment on cell A1",
        bytes: write_wb(wb),
        assertions: vec![crate::Assertion::SheetCount { expected: 1 }],
    }
}

pub fn fixture_sml_comment_rich() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", "Cell with rich comment");
    let mut cb = CommentBuilder::new_rich("A1");
    cb.add_run("Author: ").set_bold(true);
    cb.add_run("This is the body of the comment.");
    sheet.add_comment_builder(cb);
    crate::Fixture {
        path: "sml/comment/rich.xlsx",
        description: "Rich text comment on cell A1",
        bytes: write_wb(wb),
        assertions: vec![crate::Assertion::SheetCount { expected: 1 }],
    }
}

// ---------------------------------------------------------------------------
// sml/validation — data validation
// ---------------------------------------------------------------------------

pub fn fixture_sml_validation_list() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", "");
    let dv = DataValidationBuilder::list("A1", "\"Yes,No,Maybe\"");
    sheet.add_data_validation(dv);
    crate::Fixture {
        path: "sml/validation/list.xlsx",
        description: "Dropdown list validation on A1",
        bytes: write_wb(wb),
        assertions: vec![crate::Assertion::SheetCount { expected: 1 }],
    }
}

pub fn fixture_sml_validation_integer_range() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", WriteCellValue::Number(50.0));
    let dv = DataValidationBuilder::whole_number("A1", DataValidationOperator::Between, "1")
        .with_formula2("100");
    sheet.add_data_validation(dv);
    crate::Fixture {
        path: "sml/validation/integer-range.xlsx",
        description: "Integer range validation 1-100 on A1",
        bytes: write_wb(wb),
        assertions: vec![crate::Assertion::SheetCount { expected: 1 }],
    }
}

pub fn fixture_sml_validation_ignored_error() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    // Store a number as a string to trigger the warning, then ignore it.
    sheet.set_cell("A1", "42");
    sheet.add_ignored_error("A1", IgnoredErrorType::NumberStoredAsText);
    crate::Fixture {
        path: "sml/validation/ignored-error.xlsx",
        description: "Ignored number-stored-as-text error on A1",
        bytes: write_wb(wb),
        assertions: vec![crate::Assertion::SheetCount { expected: 1 }],
    }
}

// ---------------------------------------------------------------------------
// sml/filtering — auto-filter and sort state
// ---------------------------------------------------------------------------

pub fn fixture_sml_filtering_auto_filter() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", "Name");
    sheet.set_cell("B1", "Value");
    sheet.set_cell("C1", "Category");
    sheet.set_auto_filter("A1:C1");
    crate::Fixture {
        path: "sml/filtering/auto-filter.xlsx",
        description: "Auto-filter on A1:C1",
        bytes: write_wb(wb),
        assertions: vec![crate::Assertion::SheetCount { expected: 1 }],
    }
}

pub fn fixture_sml_filtering_sort_state() -> crate::Fixture {
    // The writer API does not currently expose a set_sort_state method;
    // write a sheet with data that would logically be sorted on column A.
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", "Alpha");
    sheet.set_cell("A2", "Beta");
    sheet.set_cell("A3", "Gamma");
    crate::Fixture {
        path: "sml/filtering/sort-state.xlsx",
        description: "Sort state ascending on column A",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::SheetCount { expected: 1 },
            crate::Assertion::CellValue {
                sheet: 0,
                row: 0,
                col: 0,
                expected: "Alpha".into(),
                tolerance: 0.0,
            },
        ],
    }
}

// ---------------------------------------------------------------------------
// sml/structure — workbook structure
// ---------------------------------------------------------------------------

pub fn fixture_sml_structure_multiple_sheets() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    wb.add_sheet("Alpha").set_cell("A1", "Sheet 1");
    wb.add_sheet("Beta").set_cell("A1", "Sheet 2");
    wb.add_sheet("Gamma").set_cell("A1", "Sheet 3");
    crate::Fixture {
        path: "sml/structure/multiple-sheets.xlsx",
        description: "Workbook with three sheets Alpha, Beta, Gamma",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::SheetCount { expected: 3 },
            crate::Assertion::SheetName {
                sheet: 0,
                expected: "Alpha".into(),
            },
            crate::Assertion::SheetName {
                sheet: 1,
                expected: "Beta".into(),
            },
            crate::Assertion::SheetName {
                sheet: 2,
                expected: "Gamma".into(),
            },
        ],
    }
}

pub fn fixture_sml_structure_defined_name() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", WriteCellValue::Number(1.0));
    sheet.set_cell("A2", WriteCellValue::Number(2.0));
    sheet.set_cell("A3", WriteCellValue::Number(3.0));
    sheet.set_cell("A4", WriteCellValue::Number(4.0));
    sheet.set_cell("A5", WriteCellValue::Number(5.0));
    wb.add_defined_name("MyRange", "Sheet1!$A$1:$A$5");
    crate::Fixture {
        path: "sml/structure/defined-name.xlsx",
        description: "Workbook with defined name MyRange",
        bytes: write_wb(wb),
        assertions: vec![crate::Assertion::SheetCount { expected: 1 }],
    }
}

pub fn fixture_sml_structure_print_area() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", "Header");
    sheet.set_cell("A2", "Data");
    wb.set_print_area(0, "Sheet1!$A$1:$G$20");
    crate::Fixture {
        path: "sml/structure/print-area.xlsx",
        description: "Sheet with print area set",
        bytes: write_wb(wb),
        assertions: vec![crate::Assertion::SheetCount { expected: 1 }],
    }
}

// ---------------------------------------------------------------------------
// sml/protection — protection
// ---------------------------------------------------------------------------

pub fn fixture_sml_protection_sheet() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", "Protected");
    sheet.set_sheet_protection(ooxml_sml::writer::SheetProtectionOptions {
        sheet: true,
        password: None,
        ..Default::default()
    });
    crate::Fixture {
        path: "sml/protection/sheet.xlsx",
        description: "Sheet protection with no password",
        bytes: write_wb(wb),
        assertions: vec![crate::Assertion::SheetCount { expected: 1 }],
    }
}

pub fn fixture_sml_protection_workbook() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    wb.add_sheet("Sheet1").set_cell("A1", "Protected workbook");
    wb.set_workbook_protection(true, false, None);
    crate::Fixture {
        path: "sml/protection/workbook.xlsx",
        description: "Workbook protection with lock_structure",
        bytes: write_wb(wb),
        assertions: vec![crate::Assertion::SheetCount { expected: 1 }],
    }
}

// ---------------------------------------------------------------------------
// sml/pivot — pivot tables
// ---------------------------------------------------------------------------

pub fn fixture_sml_pivot_basic() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    // Source data on Sheet1
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", "Region");
    sheet.set_cell("B1", "Sales");
    sheet.set_cell("A2", "North");
    sheet.set_cell("B2", WriteCellValue::Number(100.0));
    sheet.set_cell("A3", "South");
    sheet.set_cell("B3", WriteCellValue::Number(200.0));
    sheet.set_cell("A4", "North");
    sheet.set_cell("B4", WriteCellValue::Number(150.0));

    // Pivot on Sheet2
    let pivot_sheet = wb.add_sheet("Sheet2");
    pivot_sheet.add_pivot_table(ooxml_sml::writer::PivotTableOptions {
        name: "SalesPivot".to_string(),
        source_ref: "Sheet1!$A$1:$B$4".to_string(),
        dest_ref: "A1".to_string(),
        row_fields: vec!["Region".to_string()],
        col_fields: vec![],
        data_fields: vec!["Sales".to_string()],
    });

    crate::Fixture {
        path: "sml/pivot/basic.xlsx",
        description: "Simple pivot table with row field and data field",
        bytes: write_wb(wb),
        assertions: vec![crate::Assertion::SheetCount { expected: 2 }],
    }
}

// ---------------------------------------------------------------------------
// sml/chart — embedded charts
// ---------------------------------------------------------------------------

const BAR_CHART_XML: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart><c:plotArea><c:barChart>
    <c:barDir val="col"/><c:grouping val="clustered"/>
    <c:ser><c:idx val="0"/><c:order val="0"/>
      <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f>
        <c:numCache><c:formatCode>General</c:formatCode>
          <c:ptCount val="3"/>
          <c:pt idx="0"><c:v>1</c:v></c:pt>
          <c:pt idx="1"><c:v>2</c:v></c:pt>
          <c:pt idx="2"><c:v>3</c:v></c:pt>
        </c:numCache></c:numRef></c:val>
    </c:ser>
  </c:barChart></c:plotArea></c:chart>
</c:chartSpace>"#;

// ---------------------------------------------------------------------------
// sml/edge-case — structural edge cases
// ---------------------------------------------------------------------------

pub fn fixture_sml_edge_empty_sheet() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    wb.add_sheet("Empty");
    crate::Fixture {
        path: "sml/edge-case/empty-sheet.xlsx",
        description: "Workbook with a single empty sheet (no cells)",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::SheetCount { expected: 1 },
            crate::Assertion::SheetName {
                sheet: 0,
                expected: "Empty".into(),
            },
        ],
    }
}

pub fn fixture_sml_edge_sparse_cells() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sparse");
    sheet.set_cell_at(1, 1, WriteCellValue::String("A1".into()));
    sheet.set_cell_at(1000, 100, WriteCellValue::String("CV1000".into()));
    crate::Fixture {
        path: "sml/edge-case/sparse-cells.xlsx",
        description: "Sheet with cells at A1 and CV1000 (large gap)",
        bytes: write_wb(wb),
        assertions: vec![crate::Assertion::CellValue {
            sheet: 0,
            row: 0,
            col: 0,
            expected: "A1".into(),
            tolerance: 0.0,
        }],
    }
}

pub fn fixture_sml_edge_many_sheets() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    for i in 0..20 {
        let sheet = wb.add_sheet(format!("Sheet{i}"));
        sheet.set_cell_at(1, 1, WriteCellValue::String(format!("data {i}")));
    }
    crate::Fixture {
        path: "sml/edge-case/many-sheets.xlsx",
        description: "Workbook with 20 sheets",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::SheetCount { expected: 20 },
            crate::Assertion::SheetName {
                sheet: 0,
                expected: "Sheet0".into(),
            },
            crate::Assertion::SheetName {
                sheet: 19,
                expected: "Sheet19".into(),
            },
        ],
    }
}

pub fn fixture_sml_edge_mixed_types_column() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Mixed");
    sheet.set_cell_at(1, 1, WriteCellValue::String("text".into()));
    sheet.set_cell_at(2, 1, WriteCellValue::Number(42.0));
    sheet.set_cell_at(3, 1, WriteCellValue::Boolean(true));
    sheet.set_cell_at(4, 1, WriteCellValue::String("".into()));
    crate::Fixture {
        path: "sml/edge-case/mixed-types-column.xlsx",
        description: "Single column with string, number, boolean, and empty string",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::CellType {
                sheet: 0,
                row: 0,
                col: 0,
                expected: "string".into(),
            },
            crate::Assertion::CellType {
                sheet: 0,
                row: 1,
                col: 0,
                expected: "number".into(),
            },
            crate::Assertion::CellType {
                sheet: 0,
                row: 2,
                col: 0,
                expected: "boolean".into(),
            },
        ],
    }
}

pub fn fixture_sml_edge_merge_and_data() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell_at(1, 1, WriteCellValue::String("merged".into()));
    sheet.merge_cells("A1:B2");
    sheet.set_cell_at(1, 3, WriteCellValue::Number(100.0));
    sheet.set_cell_at(3, 1, WriteCellValue::String("below merge".into()));
    crate::Fixture {
        path: "sml/edge-case/merge-and-data.xlsx",
        description: "Sheet with merged region A1:B2 and data adjacent and below",
        bytes: write_wb(wb),
        assertions: vec![
            crate::Assertion::MergedRegion {
                sheet: 0,
                row: 0,
                col: 0,
                expected: true,
            },
            crate::Assertion::CellValue {
                sheet: 0,
                row: 0,
                col: 2,
                expected: "100".into(),
                tolerance: 0.0,
            },
            crate::Assertion::CellValue {
                sheet: 0,
                row: 2,
                col: 0,
                expected: "below merge".into(),
                tolerance: 0.0,
            },
        ],
    }
}

// ---------------------------------------------------------------------------
// sml/chart — charts
// ---------------------------------------------------------------------------

pub fn fixture_sml_chart_bar() -> crate::Fixture {
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("Sheet1");
    sheet.set_cell("A1", WriteCellValue::Number(1.0));
    sheet.set_cell("A2", WriteCellValue::Number(2.0));
    sheet.set_cell("A3", WriteCellValue::Number(3.0));
    // Embed the chart at column 2, row 0, spanning 6 cols x 15 rows
    sheet.embed_chart(BAR_CHART_XML, 2, 0, 6, 15);
    crate::Fixture {
        path: "sml/chart/bar.xlsx",
        description: "Minimal embedded bar chart",
        bytes: write_wb(wb),
        assertions: vec![crate::Assertion::SheetCount { expected: 1 }],
    }
}
