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

## Critical: Generated Parsers Don't Handle Namespace Prefixes

- [x] **Fix codegen parsers to use `local_name()` instead of `name()`** - Generated `FromXml` parsers now use `local_name().as_ref()` for element and attribute matching, so namespace prefixes (`w:body`, `x:row`) are handled correctly.
- [ ] **Full namespace URI validation** - `local_name()` matching ignores namespace URIs entirely. Full validation would require switching from `quick_xml::Reader` to `NsReader` (yields `(ResolvedNamespace, Event)` tuples) and restructuring all generated parser event loops. Not needed in practice since each OOXML part has a single primary namespace and the parser is type-scoped, but worth tracking as a future correctness improvement.

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
- [ ] **Add EG_\* name mappings to ooxml-names.yaml** - Current names are mechanical (`EGBlockLevelElts`, `block_level_elts`). Add human-readable mappings (e.g. `BlockContent`, `ParagraphContent`, `rows`). Affects both generated.rs and generated_parsers.rs consistently.
- [ ] **Regenerate and verify** - Check `#[cfg(feature = "...")]` annotations appear on the right fields.

### Phase 3: Generate FromXml parsers for WML
- [x] **Add `generate_parsers()` to WML build.rs** - Copy SML's pattern (OOXML_GENERATE_PARSERS env var).
- [x] **Generate `src/generated_parsers.rs`** - FromXml trait impls using parser_gen.rs infrastructure. 24,884 lines of event-based parsers.
- [x] **Add `pub mod generated_parsers` to lib.rs**.
- [x] **Fix parser_gen.rs for EG_\*/AG_\* composition** - Handle element group content fields, attribute group inlining, recursive variant flattening, overlapping variant dedup, hexBinary types, CT_Empty strategy.
- [ ] **Unit tests** - Parse XML snippets with `FromXml::from_xml()` for Document, Paragraph, Run, RunProperties, Table.

### Phase 4: Extension traits (ext.rs) ✅
- [x] **Pure traits** - `DocumentExt`, `BodyExt`, `ParagraphExt`, `RunExt`, `RunPropertiesExt` (wml-styling gated), `HyperlinkExt`, `TableExt`, `RowExt`, `CellExt`, `SectionPropertiesExt` (wml-layout gated).
- [x] **Resolve traits** - `StyleContext` (holds styles + docDefaults), `RunResolveExt` (resolve bold/italic/font size/font/color via basedOn chain, depth-limit 20). `ResolvedDocument` wrapper.
- [x] **Wrapper functions** - `parse_document(xml)`, `parse_styles(xml)` using generated FromXml. Now works on real prefixed OOXML content.

### Phase 5: Parity tests ✅
- [x] **Parser parity tests** - Parse same XML with both handwritten and generated parsers, compare results via extension traits. Cover: simple docs, formatted text, tables (basic/merged/nested), section properties, hyperlinks, comments, footnotes.
- [x] **Fixture parity** - All `.docx` files in `tests/fixtures/` parsed by both paths.

### Phase 6: Replace handwritten document parser with generated ✅
- [x] **Upgrade stub types** - CTDrawing, CTRunTrackChange, EGPContentMath now capture raw XML (extra_attrs + extra_children) instead of skipping content.
- [x] **Add `read`/`write` feature gates** - `read` enables Document::from_reader() + generated parser + ext traits; `write` enables DocumentBuilder + handwritten types.
- [x] **Switch Document<R> to generated types** - Stores `types::Document` + `types::Styles` instead of handwritten Body/Styles. Users access content via ext traits.
- [x] **Add DrawingExt trait** - Walks raw XML in CTDrawing to extract image relationship IDs (inline + anchored).
- [x] **Gate modules behind features** - writer.rs behind `write`, styles.rs behind `write`, ext.rs behind `read`.
- [x] **Update integration tests** - Use generated types + ext traits for assertions.
- [ ] **Delete parse_document()** - ~2100 lines still present for header/footer parsing. Remove when headers/footers migrate to generated types.

### Phase 7: Migrate writer to generated types (HIGH PRIORITY)
- [ ] **Add ToXml codegen** - Generate `ToXml` trait impls alongside `FromXml`.
- [ ] **Update DocumentBuilder** - Accept generated types instead of handwritten types.
- [ ] **Delete handwritten types** - Remove Body, Paragraph, Run, etc. from document.rs.
- [ ] **Delete styles.rs** - Replaced by ext.rs style resolution.
- [ ] **Remove `write` feature gate** - Once writer uses generated types, merge read+write.

### Phase 8: Codegen stub type upgrades (HIGH PRIORITY)
- [ ] **Update codegen to generate extra_attrs/extra_children for stub types** - CTDrawing, CTRunTrackChange, EGPContentMath currently hand-edited. Codegen should produce the same pattern.
- [ ] **Delete parse_document()** - Once headers/footers use generated types, remove ~2100 lines of handwritten parsing code.

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
