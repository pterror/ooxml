//! CC0 test fixture generation for the ooxml library.
//!
//! This crate reads `spec/ooxml-fixture-spec.yaml`, generates fixture functions
//! via codegen, and writes `.docx`/`.xlsx`/`.pptx` files and companion JSON
//! manifests to `fixtures/`.

pub mod manifest;

pub mod custom;

pub mod generated;

/// A single test fixture: an OOXML file and the assertions that verify it.
pub struct Fixture {
    /// Relative path within `fixtures/` (e.g. `"wml/text/bold.docx"`).
    pub path: &'static str,
    /// Human-readable description of what this fixture tests.
    pub description: &'static str,
    /// Raw bytes of the generated OOXML file.
    pub bytes: Vec<u8>,
    /// Assertions that a reader must satisfy when parsing `bytes`.
    pub assertions: Vec<Assertion>,
}

/// An assertion that can be verified by parsing the fixture file.
#[allow(dead_code)]
pub enum Assertion {
    // -------------------------------------------------------------------------
    // WML — document body
    // -------------------------------------------------------------------------
    /// Total number of paragraphs in the document body.
    ParagraphCount { expected: u32 },
    /// Plain text of a paragraph (all runs concatenated).
    ParagraphText { para: usize, expected: String },
    /// Style name applied to a paragraph.
    ParagraphStyle { para: usize, expected: String },
    /// Paragraph alignment: `"left"`, `"center"`, `"right"`, or `"justify"`.
    ParagraphAlign { para: usize, expected: String },
    /// Numbering list level of a paragraph (`None` means not in a list).
    ParagraphListLevel { para: usize, expected: Option<u32> },

    // -------------------------------------------------------------------------
    // WML — runs
    // -------------------------------------------------------------------------
    /// Plain text of a run.
    RunText {
        para: usize,
        run: usize,
        expected: String,
    },
    /// Whether a run is bold.
    RunBold {
        para: usize,
        run: usize,
        expected: bool,
    },
    /// Whether a run is italic.
    RunItalic {
        para: usize,
        run: usize,
        expected: bool,
    },
    /// Whether a run has any underline applied.
    RunUnderline {
        para: usize,
        run: usize,
        expected: bool,
    },
    /// Whether a run has strikethrough applied.
    RunStrikethrough {
        para: usize,
        run: usize,
        expected: bool,
    },
    /// Run foreground color as `"RRGGBB"`, or `None` for default color.
    RunColor {
        para: usize,
        run: usize,
        expected: Option<String>,
    },
    /// Run font size in points (half-points in the spec; converted here).
    RunFontSize {
        para: usize,
        run: usize,
        expected: f64,
        tolerance: f64,
    },
    /// Font family name, or `None` for the theme font.
    RunFontName {
        para: usize,
        run: usize,
        expected: Option<String>,
    },

    // -------------------------------------------------------------------------
    // WML — tables
    // -------------------------------------------------------------------------
    /// Number of rows in a table.
    TableRows { table: usize, expected: u32 },
    /// Number of cells in a table row.
    TableCols {
        table: usize,
        row: usize,
        expected: u32,
    },
    /// Plain text content of a table cell.
    TableCellText {
        table: usize,
        row: usize,
        col: usize,
        expected: String,
    },
    /// Column span of a table cell (1 = no merge).
    TableCellColspan {
        table: usize,
        row: usize,
        col: usize,
        expected: u32,
    },
    /// Row span of a table cell (1 = no merge).
    TableCellRowspan {
        table: usize,
        row: usize,
        col: usize,
        expected: u32,
    },

    // -------------------------------------------------------------------------
    // WML — miscellaneous
    // -------------------------------------------------------------------------
    /// URL of a hyperlink on a run, or `None` if the run is not a hyperlink.
    HyperlinkUrl {
        para: usize,
        run: usize,
        expected: Option<String>,
    },
    /// Total number of embedded images in the document.
    ImageCount { expected: u32 },
    /// Sorted list of bookmark names defined in the document.
    BookmarkNames { expected: Vec<String> },

    // -------------------------------------------------------------------------
    // SML — workbook
    // -------------------------------------------------------------------------
    /// Number of sheets in the workbook.
    SheetCount { expected: u32 },
    /// Name of a sheet (0-indexed).
    SheetName { sheet: usize, expected: String },

    // -------------------------------------------------------------------------
    // SML — cells
    // -------------------------------------------------------------------------
    /// Cell type: `"string"`, `"number"`, `"boolean"`, `"blank"`, or `"error"`.
    CellType {
        sheet: usize,
        row: usize,
        col: usize,
        expected: String,
    },
    /// Cell value as a string (numbers formatted without trailing zeros where
    /// possible); `tolerance` is used for numeric comparisons.
    CellValue {
        sheet: usize,
        row: usize,
        col: usize,
        expected: String,
        tolerance: f64,
    },
    /// Formula text of a cell, or `None` if the cell has no formula.
    CellFormula {
        sheet: usize,
        row: usize,
        col: usize,
        expected: Option<String>,
    },
    /// Number format code of a cell, or `None` for the general format.
    CellFormatCode {
        sheet: usize,
        row: usize,
        col: usize,
        expected: Option<String>,
    },
    /// Whether a cell's font is bold.
    CellBold {
        sheet: usize,
        row: usize,
        col: usize,
        expected: bool,
    },
    /// Whether a cell's font is italic.
    CellItalic {
        sheet: usize,
        row: usize,
        col: usize,
        expected: bool,
    },
    /// Cell background color as `"RRGGBB"`, or `None` for no fill.
    CellColor {
        sheet: usize,
        row: usize,
        col: usize,
        expected: Option<String>,
    },
    /// Whether the cell is part of a merged region.
    MergedRegion {
        sheet: usize,
        row: usize,
        col: usize,
        expected: bool,
    },
    /// Row height in points.
    RowHeight {
        sheet: usize,
        row: usize,
        expected: f64,
        tolerance: f64,
    },
    /// Column width in character units.
    ColWidth {
        sheet: usize,
        col: usize,
        expected: f64,
        tolerance: f64,
    },

    // -------------------------------------------------------------------------
    // PML — presentation
    // -------------------------------------------------------------------------
    /// Total number of slides.
    SlideCount { expected: u32 },
    /// Number of shapes on a slide.
    ShapeCount { slide: usize, expected: u32 },
    /// Concatenated text of all paragraphs/runs in a shape.
    ShapeText {
        slide: usize,
        shape: usize,
        expected: String,
    },
    /// Preset geometry type of a shape: `"rect"`, `"ellipse"`, etc.
    ShapeType {
        slide: usize,
        shape: usize,
        expected: String,
    },

    // -------------------------------------------------------------------------
    // PML — runs inside shapes
    // -------------------------------------------------------------------------
    /// Plain text of a run inside a shape paragraph.
    PmlRunText {
        slide: usize,
        shape: usize,
        para: usize,
        run: usize,
        expected: String,
    },
    /// Whether a PML run is bold.
    PmlRunBold {
        slide: usize,
        shape: usize,
        para: usize,
        run: usize,
        expected: bool,
    },
    /// Whether a PML run is italic.
    PmlRunItalic {
        slide: usize,
        shape: usize,
        para: usize,
        run: usize,
        expected: bool,
    },
    /// PML run color as `"RRGGBB"`, or `None` for theme color.
    PmlRunColor {
        slide: usize,
        shape: usize,
        para: usize,
        run: usize,
        expected: Option<String>,
    },
    /// PML run font size in points.
    PmlRunFontSize {
        slide: usize,
        shape: usize,
        para: usize,
        run: usize,
        expected: f64,
        tolerance: f64,
    },

    // -------------------------------------------------------------------------
    // PML — miscellaneous
    // -------------------------------------------------------------------------
    /// Whether a slide has speaker notes.
    SlideHasNotes { slide: usize, expected: bool },
    /// Plain text of the speaker notes for a slide.
    NotesText { slide: usize, expected: String },
    /// Number of embedded images on a slide.
    PmlImageCount { slide: usize, expected: u32 },
    /// Whether a slide has a transition defined.
    HasTransition { slide: usize, expected: bool },
}
