# TODO

Backlog for ooxml development. Items roughly ordered by priority within each section.

## Writing API Gaps

DocumentBuilder handles common cases but doesn't expose:

- [x] Headers/footers creation
- [x] Footnotes/endnotes creation
- [x] Comments creation
- [x] Anchored/floating images (use `run.add_anchored_image()` with `set_wrap_type()`)
- [x] Styles creation/writing - add_style(), set_default_paragraph/run(), serialize()

## Reading Enhancements

- [x] More w:rFonts attributes (hAnsi, eastAsia, cs) - now parses all four font attributes
- [x] Core properties (docProps/core.xml) - title, author, etc.
- [x] App properties (docProps/app.xml) - word count, pages, etc.

## Advanced Features

- [x] Revision tracking (w:ins, w:del for tracked changes) - Insertion/Deletion types with author/date, parsing and serialization
- [x] Math equations (integrate ooxml-omml crate) - parsing m:oMath, serialize_math_zone, MathZone re-export
- [x] Table of contents (read) - headings() method + SimpleField + outline_level
- [x] Bookmarks (add_bookmark_start/end methods, exported types)
- [x] Text boxes - text extraction via VmlPicture/Drawing/EmbeddedObject.text()
- [ ] SmartArt (limited)
- [x] Charts (limited) - SML reading via Chart, ChartType, ChartSeries types

## Developer Experience

- [x] Better error messages with context - ParseContext, Error::Parse with position, position_to_line_col()
- [x] More examples in docs - read_xlsx, cell_access for ooxml-sml; read_pptx, extract_text for ooxml-pml
- [x] Real-world usage examples in `examples/` directory (read_docx, create_docx, read_metadata)
- [ ] API documentation improvements
- [ ] Consider using Default + field assignment instead of with_ builder methods for simpler configuration structs

## Technical Debt

- [x] **Generate types for SML/PML/DML from schemas** - All crates now use codegen from ECMA-376 RELAX NG schemas (wml.rnc, sml.rnc, pml.rnc, dml-main.rnc). Generated types are committed to avoid spec dependency.

## Codegen Performance

- [ ] **Lazy/cursor API** - Alternate API that avoids struct materialization for high-performance streaming use cases. Iterator-based access to rows/cells without allocating intermediate structs.
- [x] **Feature-gated extra_attrs** - `extra-attrs` feature captures unhandled attributes in `extra_attrs: HashMap<String, String>` for roundtrip fidelity. Enabled by default.
- [x] **Feature-gated unknown children** - `extra-children` feature captures unknown child elements in `extra_children: Vec<ooxml_xml::RawXmlNode>`. Uses shared `ooxml-xml` crate.
- [x] **Per-field feature gating** - Uses `spec/ooxml-features.yaml` to gate non-core fields behind features (sml-styling, sml-formulas, etc.). 265 fields gated, 893 parser locations.
- [x] **Extension trait cfg attrs** - Feature-gated WorksheetExt and ResolvedSheet methods for `--no-default-features` support.

## Codegen Migration

Replace handwritten types with generated equivalents. **Requires extensive test fixtures first** to ensure no regressions.

### Prerequisites
- [ ] **Comprehensive roundtrip test suite** - Parse real documents, serialize, compare output. Cover all element types.
- [ ] **Fixture corpus** - Curated set of DOCX/XLSX/PPTX files exercising edge cases (from corpus analysis + synthetic).
- [ ] **Parity tests** - For each handwritten type, verify generated equivalent produces identical XML output.

### WML (Word)
- [ ] **Port feature flags to WML** - Add ooxml-features.yaml mappings for WML elements, feature-gate non-core fields.
- [ ] **Port extra-attrs/extra-children to WML** - Already has `unknown_children`, unify with codegen approach.
- [ ] **Replace handwritten WML types** - Swap document.rs/paragraph.rs/etc with generated types, update ext traits.
- [ ] **Delete handwritten WML code** - Remove old implementations once tests pass.

### PML (PowerPoint)
- [ ] **Port feature flags to PML** - Add ooxml-features.yaml mappings for PML elements.
- [ ] **Port extra-attrs/extra-children to PML** - Enable roundtrip fidelity features.
- [ ] **Replace handwritten PML types** - Swap with generated types, update ext traits.
- [ ] **Delete handwritten PML code** - Remove old implementations once tests pass.

### DML (Drawing)
- [ ] **Port feature flags to DML** - Add ooxml-features.yaml mappings for DML elements.
- [ ] **Port extra-attrs/extra-children to DML** - Enable roundtrip fidelity features.
- [ ] **Replace handwritten DML types** - Swap with generated types, update ext traits.
- [ ] **Delete handwritten DML code** - Remove old implementations once tests pass.

## Robustness

- [ ] Edge case handling from corpus analysis
- [ ] More comprehensive tests against real-world documents
- [ ] Fuzz testing for malformed input
- [ ] Synthetic fixtures from corpus insights - analyze failures, create minimal repro cases
- [ ] Auto-generated fixtures from ECMA-376 XSD schemas - ensure all element types are tested

## Other Formats

### SpreadsheetML (Excel) - ooxml-sml

- [x] Workbook structure - Workbook::open(), sheet_count(), sheet_names()
- [x] Worksheets - Sheet with rows(), cell(), dimensions()
- [x] Cells and values - Cell, CellValue (String, Number, Boolean, Error, Empty)
- [x] Formulas (as strings, not evaluated) - Cell::formula()
- [x] Basic formatting - Stylesheet, Font, Fill, Border, CellFormat types, Cell.style_index()
- [x] Shared strings - parsed and resolved automatically
- [x] Write support - WorkbookBuilder, SheetBuilder, set_cell(), set_formula(), save()
- [x] Merged cells - MergedCell, Sheet.merged_cells(), merge_cells() in writer
- [x] Column/row dimensions - ColumnInfo, Row.height(), set_column_width(), set_row_height() in writer
- [x] Cell comments/notes - Comment type, Sheet.comments(), Sheet.comment(ref)
- [x] Number format display - builtin_format_code(), date detection, Excel serial date conversion
- [x] Named ranges - DefinedName with workbook/sheet scope, defined_names(), defined_name()
- [x] Conditional formatting (read) - ConditionalFormatting, ConditionalRule, colorScale/dataBar/iconSet
- [x] Data validation (read) - DataValidation, list/whole/decimal/date/time/textLength types
- [x] Charts (read) - Chart, ChartType, ChartSeries types, Sheet.charts(), supports chartsheets

### PresentationML (PowerPoint) - ooxml-pml

- [x] Presentation structure - Presentation::open(), slide_count()
- [x] Slides - slide(), slides(), Slide with index(), shapes(), text()
- [x] Shapes and text - Shape with name(), paragraphs(), text(), has_text()
- [x] Slide notes - Slide::notes(), has_notes()
- [x] Images - Slide::pictures(), Picture with rel_id/name/description, get_image_data()
- [x] Basic transitions - Slide::transition(), Transition with type/speed/advance settings
- [x] Write support - PresentationBuilder, SlideBuilder with add_title(), add_text(), save()
- [x] Hyperlinks in shapes - Shape.hyperlinks(), Run.hyperlink_rel_id(), resolve_hyperlink()
- [x] Slide layouts/masters (read) - SlideMaster, SlideLayout, SlideLayoutType, layout_by_name()
- [x] Speaker notes in writer - SlideBuilder::set_notes(), notesSlide generation

### DrawingML - ooxml-dml

- [ ] Full DrawingML support (currently minimal)

## Infrastructure

- [ ] GitHub Pages documentation
- [ ] CI/CD for crates.io publishing
- [ ] Changelog generation
