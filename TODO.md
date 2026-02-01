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

- [x] **Lazy/cursor API** - `LazyWorksheet` provides streaming row iteration without materializing all rows. `Workbook::sheet_xml()` returns raw XML for lazy parsing.
- [x] **Feature-gated extra_attrs** - `extra-attrs` feature captures unhandled attributes in `extra_attrs: HashMap<String, String>` for roundtrip fidelity. Enabled by default.
- [x] **Feature-gated unknown children** - `extra-children` feature captures unknown child elements in `extra_children: Vec<ooxml_xml::RawXmlNode>`. Uses shared `ooxml-xml` crate.
- [x] **Per-field feature gating** - Uses `spec/ooxml-features.yaml` to gate non-core fields behind features (sml-styling, sml-formulas, etc.). 265 fields gated, 893 parser locations.
- [x] **Extension trait cfg attrs** - Feature-gated WorksheetExt and ResolvedSheet methods for `--no-default-features` support.

## WML Codegen Migration

Replace ~8,750 lines of handwritten WML parsing (document.rs + styles.rs) with codegen'd types and FromXml parsers.

### Phase 1: Fix codegen to compose element groups ✅
- [x] **Fix `collect_fields()` for EG_\* refs** - Element choice groups become `Vec<Box<EGType>>` fields; struct-like groups inline their fields.
- [x] **Flatten element group enums** - `collect_element_variants()` recursively follows EG_\* refs with cycle detection.
- [x] **Handle AG_\* attribute groups** - Inline attribute group fields into parent structs.
- [x] **Regenerate WML types** - Body, Paragraph, Run, RunProperties, SectionProperties, Table, CTRow, TableCell all now complete.

### Phase 2: Expand WML feature mappings
- [ ] **Expand ooxml-features.yaml WML section** - Currently ~20 lines covering 5 types. Use element constants in document.rs:3753-3915 as checklist. Map to feature tags (core, wml-styling, wml-tables, wml-layout, wml-hyperlinks, wml-drawings).
- [ ] **Regenerate and verify** - Check `#[cfg(feature = "...")]` annotations appear on the right fields.

### Phase 3: Generate FromXml parsers for WML
- [ ] **Add `generate_parsers()` to WML build.rs** - Copy SML's pattern (OOXML_GENERATE_PARSERS env var).
- [ ] **Generate `src/generated_parsers.rs`** - FromXml trait impls using parser_gen.rs infrastructure.
- [ ] **Add `pub mod generated_parsers` to lib.rs**.
- [ ] **Unit tests** - Parse XML snippets with `FromXml::from_xml()` for Document, Paragraph, Run, RunProperties, Table.

### Phase 4: Extension traits (ext.rs)
- [ ] **Pure traits** - `BodyExt` (iterate paragraphs/tables, extract text), `ParagraphExt` (iterate runs, get style/alignment), `RunExt` (get text, check bold/italic/underline, font size/color), `TableExt` (iterate rows, get properties), `SectionPropertiesExt` (page size, margins).
- [ ] **Resolve traits** - `WmlResolveContext` (holds Styles, Numbering), `ParagraphResolveExt` (resolve effective style via basedOn chain), `RunResolveExt` (resolve effective run properties). Port style inheritance logic from styles.rs (basedOn chain walking, depth-limit cycle detection, property merging).
- [ ] **Wrapper functions** - `parse_document(xml)`, `parse_styles(xml)` using generated FromXml.

### Phase 5: Parity tests
- [ ] **Parser parity tests** - Parse same XML with both handwritten and generated parsers, compare results via extension traits. Cover: simple docs, formatted text, tables (basic/merged/nested), section properties, hyperlinks, comments, footnotes.
- [ ] **Fixture parity** - All `.docx` files in `tests/fixtures/` parsed by both paths.

### Phase 6: Migrate writer + public API
- [ ] **Update writer.rs** - Change serialize functions to accept generated types.
- [ ] **Update lib.rs exports** - Re-export generated types + extension traits instead of handwritten types.
- [ ] **Type aliases** - Add backward-compat aliases where names differ (Cell → TableCell, Row → CTRow).

### Phase 7: Remove handwritten parsing
- [ ] **Gut document.rs** - Remove ~6,000+ lines of parsing code and type definitions; keep Document struct + OPC loading.
- [ ] **Delete styles.rs** - Logic moved to ext.rs resolve traits.
- [ ] **Expected result** - ~6,000+ lines removed, replaced by generated code + ~800-1000 lines of extension traits.

## Other Codegen Migrations

### SML (Spreadsheet)
- [ ] **Regenerate SML types with EG_\*/AG_\* inlining** - Codegen now supports it; requires regenerating generated_parsers.rs too.
- [ ] **Expand SML feature mappings** - Cover remaining ungated fields.

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
