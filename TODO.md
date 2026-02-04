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
- [ ] **Remove unnecessary Box in Vec<Box<T>>** - Codegen uses `Vec<Box<T>>` for all struct types even when not recursive. This adds unnecessary indirection. Consider generating `Vec<T>` for non-recursive types and reserving `Box<T>` for recursive types only. Clippy flags this with `clippy::vec_box`.
- [x] **Handle repeating choice patterns** - Fixed codegen to handle `OneOrMore(Group(Choice([...])))` patterns. Font/Fill types now have proper optional fields for each choice alternative.

## Codegen Performance

- [x] **Lazy/cursor API** - `LazyWorksheet` provides streaming row iteration without materializing all rows. `Workbook::sheet_xml()` returns raw XML for lazy parsing.
- [x] **Feature-gated extra_attrs** - `extra-attrs` feature captures unhandled attributes in `extra_attrs: HashMap<String, String>` for roundtrip fidelity. Enabled by default.
- [x] **Feature-gated unknown children** - `extra-children` feature captures unknown child elements in `extra_children: Vec<ooxml_xml::PositionedNode>`. Position-indexed for roundtrip ordering fidelity (ADR-004).
- [x] **Per-field feature gating** - Uses `spec/ooxml-features.yaml` to gate non-core fields behind features (sml-styling, sml-formulas, etc.). 265 fields gated, 893 parser locations.
- [x] **Extension trait cfg attrs** - Feature-gated WorksheetExt and ResolvedSheet methods for `--no-default-features` support.

## WML Codegen Migration

Replace ~8,750 lines of handwritten WML parsing (document.rs + styles.rs) with codegen'd types and FromXml parsers.

### Phase 1: Fix codegen to compose element groups ✅
- [x] **Fix `collect_fields()` for EG_\* refs** - Element choice groups become `Vec<Box<EGType>>` fields; struct-like groups inline their fields.
- [x] **Flatten element group enums** - `collect_element_variants()` recursively follows EG_\* refs with cycle detection.
- [x] **Handle AG_\* attribute groups** - Inline attribute group fields into parent structs.
- [x] **Regenerate WML types** - Body, Paragraph, Run, RunProperties, SectionProperties, Table, CTRow, TableCell all now complete.

### Phase 2: Expand WML feature mappings ✅
- [x] **Expand ooxml-features.yaml WML section** - Expanded from 5 types to 18 types with ~165 field→feature mappings across 10 feature flags (wml-styling, wml-tables, wml-layout, wml-hyperlinks, wml-drawings, wml-numbering, wml-comments, wml-fields, wml-track-changes, wml-settings).
- [x] **Add EG_\* name mappings to ooxml-names.yaml** - EG_* types renamed to human-readable names (BlockContent, ParagraphContent, RunContent, RowContent, CellContent, HeaderFooterRef, etc.). EG_* fields renamed (block_content, paragraph_content, run_content, rows, cells, header_footer_refs, range_markup). CT_HdrFtr→HeaderFooter, CT_FtnEdn→FootnoteEndnote, CT_FtnEdnRef→FootnoteEndnoteRef.
- [x] **Regenerate and verify** - 165 wml-* feature annotations, 1041 total. Slim build (`--no-default-features`) compiles.

### Phase 3: Generate FromXml parsers for WML
- [x] **Add `generate_parsers()` to WML build.rs** - Copy SML's pattern (OOXML_GENERATE_PARSERS env var).
- [x] **Generate `src/generated_parsers.rs`** - FromXml trait impls using parser_gen.rs infrastructure. 24,884 lines of event-based parsers.
- [x] **Add `pub mod generated_parsers` to lib.rs**.
- [x] **Fix parser_gen.rs for EG_\*/AG_\* composition** - Handle element group content fields, attribute group inlining, recursive variant flattening, overlapping variant dedup, hexBinary types, CT_Empty strategy.
- [x] **Unit tests** - Parse XML snippets with `FromXml::from_xml()` for Document, Paragraph, Run, RunProperties, Table, ParagraphProperties, SectionProperties, Style, Hyperlink, namespace-prefixed XML, and ADR-004 position tracking.

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
- [x] **Delete parse_document()** - Handwritten parser removed. Headers/footers/footnotes/endnotes/comments all use generated FromXml parsers via ext.rs helpers.

### Phase 7: Migrate writer to generated types ✅
- [x] **Add ToXml codegen** (Phase 7a) - Generate `ToXml` trait impls alongside `FromXml`.
- [x] **Update DocumentBuilder** (Phase 7c) - Uses generated types + ToXml serialization. convenience.rs provides builder methods.
- [x] **Delete handwritten types** (Phase 7c) - Removed Body, Paragraph, Run, Table, etc. from document.rs (~7000 lines deleted).
- [x] **Delete styles.rs** (Phase 7c) - Replaced by ext.rs style resolution.
- [x] **Remove `read`/`write` feature gates** (Phase 7c) - Merged into single unconditional API.

### Phase 8: Codegen attribute inheritance ✅
- [x] **Stub types have extra_attrs/extra_children** - CTDrawing, CTRunTrackChange, EGPContentMath now generated with roundtrip capture.
- [x] **Fix attribute inheritance from base types** - `collect_fields()` now inlines CT_* mixin refs (like AG_* groups). Comment has `author`/`date`/`id`, Bookmark has `id`, CTMarkupRange has `id`.
- [x] **CTPPrBase flattening** - ParagraphProperties now includes all CTPPrBase fields (alignment, numbering, spacing, indent). convenience.rs uses proper typed fields.
- [x] **Skip wildcard elements** - Wildcard element patterns (`element * { ... }`) are excluded from field collection to avoid generating invalid `_any` fields.
- [x] **Regenerate all crates** - All four crates (WML, SML, PML, DML) regenerated with the inheritance fix.

## Other Codegen Migrations

### SML (Spreadsheet)
- [x] **Regenerate SML types with EG_\*/AG_\* inlining** - Types, parsers, and serializers regenerated with latest codegen.
- [x] **Add SML serializer unit tests** - 25 roundtrip tests in test_generated_serializers.rs.
- [x] **Migrate SML writer to generated serializers** - Complete. All serializers now use generated ToXml: serialize_shared_strings, serialize_comments, serialize_sheet, serialize_workbook, serialize_styles.
- [ ] **Expand SML feature mappings** - Cover remaining ungated fields.

### PML (PowerPoint)
- [~] **Add PML parser/serializer generation** - Build infrastructure ready (build.rs), but blocked by codegen issues. See DML section.
- [ ] **Port feature flags to PML** - Add ooxml-features.yaml mappings for PML elements.
- [ ] **Port extra-attrs/extra-children to PML** - Enable roundtrip fidelity features.
- [ ] **Replace handwritten PML types** - Swap with generated types, update ext traits.
- [ ] **Delete handwritten PML code** - Remove old implementations once tests pass.

### DML (Drawing)
- [~] **Add DML parser/serializer generation** - Build infrastructure ready (build.rs), partially blocked by codegen issues:
  - [x] Optional field serialization for EG_* types (fixed in serializer_gen.rs)
  - [ ] CT wrapper type parsing (CTColor wraps EGColorChoice, parser needs to construct wrapper)
  - [ ] Type alias handling (EGOfficeArtExtensionList vs CTOfficeArtExtensionList)
  - [ ] Cross-crate type references (PML→DML types like CTColor, CTTextListStyle)
  - Requires fixes in parser_gen.rs

### Codegen: Namespace Prefix Convention ✅
- [x] **Configurable namespace prefixes in serializers** - Added `xml_serialize_prefix` to `CodegenConfig`:
  - `None` = unprefixed elements (SML/XLSX uses default namespace convention)
  - `Some("w")` = `w:` prefixed elements (WML/DOCX)
  - `Some("p")` = `p:` prefixed elements (PML/PPTX)
  - `Some("a")` = `a:` prefixed elements (DML when embedded)
  - Verified against real Office files from corpus: XLSX uses unprefixed, DOCX uses `w:`, PPTX uses `p:`
  - SML writer migration is now unblocked
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
