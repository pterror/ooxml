# ooxml - Rust OOXML Library

High-quality Rust library for reading/writing Office Open XML formats (DOCX, XLSX, PPTX).

Part of the [Rhizome](https://rhizome-lab.github.io/) ecosystem, providing OOXML support for [rescribe](https://github.com/pterror/rescribe) document conversion.

## Why?

The Rust ecosystem lacks a mature OOXML library. Python has `python-docx`, Java has Apache POI, .NET has Open XML SDK. Rust deserves the same.

## Design Principles

1. **Typed representations** - Structs for every element, not string soup
2. **Roundtrip fidelity** - Preserve unknown elements, don't lose data
3. **Lazy parsing** - Don't parse what you don't need
4. **Incremental adoption** - Start with common features, grow over time
5. **Spec-driven** - Follow ECMA-376 / ISO 29500, document deviations

## Architecture

```
ooxml/
├── crates/                    # Rust implementation
│   ├── ooxml/                 # Core: OPC packaging, relationships
│   ├── ooxml-wml/             # WordprocessingML (Word)
│   ├── ooxml-sml/             # SpreadsheetML (Excel) - future
│   ├── ooxml-pml/             # PresentationML (PowerPoint) - future
│   └── ooxml-codegen/         # Generate types from specs
│
├── fixtures/                  # Test files (CC0 licensed)
│   ├── wml/                   # Word documents
│   ├── sml/                   # Excel spreadsheets
│   └── pml/                   # PowerPoint presentations
│
└── specs/                     # Machine-readable spec (CC0 licensed)
    ├── elements/              # Element definitions
    ├── enums/                 # Enumeration values
    └── ...
```

See [ECOSYSTEM.md](ECOSYSTEM.md) for details on fixtures and specs.

## Scope

### v0.1 - Core + Word Basics

**ooxml-opc (core):**
- [x] OPC packaging (ZIP read/write)
- [x] Relationships (.rels files)
- [x] Content types ([Content_Types].xml)
- [ ] Core properties (docProps/core.xml)
- [ ] App properties (docProps/app.xml)

**ooxml-wml (WordprocessingML):**
- [x] Document structure (document.xml)
- [x] Paragraphs (`<w:p>`)
- [x] Runs (`<w:r>`) and text (`<w:t>`)
- [x] Basic formatting: bold, italic, underline, strikethrough
- [x] Font, size, color
- [x] Paragraph properties: alignment, spacing, indentation
- [x] Headings (via paragraph styles)
- [x] Lists (numbering definitions + abstract numbering)
- [x] Tables (basic: rows, cells)
- [x] Hyperlinks (internal and external)
- [x] Images (inline, embedded in word/media/)
- [x] Styles (styles.xml) - read and apply
- [ ] Styles (styles.xml) - write
- [x] Page breaks
- [x] Section breaks
- [x] Document settings (word/settings.xml)

### v0.2 - Extended Word (Reading)

- [x] Headers and footers
- [x] Footnotes and endnotes
- [ ] Table of contents (read)
- [ ] Bookmarks
- [x] Complex tables (merged cells, nested tables)
- [ ] Text boxes
- [x] Tabs and tab stops
- [x] Borders and shading
- [x] Table borders
- [x] Content controls (w:sdt)
- [x] Custom XML blocks (w:customXml)
- [x] VML pictures (w:pict)
- [x] Embedded objects (w:object)
- [x] Comments

### v0.2 - Extended Word (Writing)

- [ ] Headers and footers creation
- [ ] Footnotes and endnotes creation
- [ ] Comments creation
- [ ] Anchored/floating images

### v0.3 - Advanced Word

- [ ] Track changes (revisions: w:ins, w:del)
- [ ] Form fields
- [ ] Math (OMML integration)
- [ ] SmartArt (limited)
- [ ] Charts (limited)

### v0.4 - SpreadsheetML (Excel)

- [ ] ooxml-sml crate
- [ ] Workbook structure
- [ ] Worksheets
- [ ] Cells and values
- [ ] Formulas (as strings, not evaluated)
- [ ] Basic formatting
- [ ] Shared strings

### v0.5 - PresentationML (PowerPoint)

- [ ] ooxml-pml crate
- [ ] Presentation structure
- [ ] Slides
- [ ] Shapes and text
- [ ] Images
- [ ] Basic transitions

### Future

- [ ] ooxml-dml: Full DrawingML
- [ ] Advanced Excel (charts, pivot tables)
- [ ] Advanced PowerPoint (animations, SmartArt)

## Dependencies

```toml
[dependencies]
zip = "2"              # ZIP archive handling
quick-xml = "0.36"     # XML parsing/writing
thiserror = "2"        # Error handling

[dev-dependencies]
insta = "1"            # Snapshot testing
```

## API Design

### Reading

```rust
let doc = ooxml_wml::Document::open("input.docx")?;
for para in doc.body().paragraphs() {
    println!("Style: {:?}", para.style());
    for run in para.runs() {
        println!("  Text: {}", run.text());
        if run.is_bold() { println!("  (bold)"); }
    }
}
```

### Writing (DocumentBuilder)

```rust
let mut builder = ooxml_wml::DocumentBuilder::new();
builder.add_paragraph("Hello, World!");

// With formatting
let para = builder.body_mut().add_paragraph();
let mut run = para.add_run();
run.set_text("Bold text");
run.set_properties(RunProperties { bold: true, ..Default::default() });

// Lists
let num_id = builder.add_list(ListType::Bullet);
let para = builder.body_mut().add_paragraph();
para.set_numbering(num_id, 0);
para.add_run().set_text("List item");

// Images
let rel_id = builder.add_image(image_bytes, "image/png");
let mut drawing = Drawing::new();
drawing.add_image(&rel_id).set_width_inches(2.0);
para.add_run().drawings_mut().push(drawing);

builder.save("output.docx")?;
```

### Roundtrip (preserves unknown elements)

```rust
let mut doc = ooxml_wml::Document::open("input.docx")?;
// Modify...
doc.save("output.docx")?;
```

## Testing Strategy

1. **Unit tests** - Individual element parsing/serialization
2. **Roundtrip tests** - Open → save → compare (structural)
3. **Fixture tests** - Real DOCX files from various sources (see `fixtures/`)
4. **Snapshot tests** - Insta for XML output verification
5. **Fuzz tests** - Malformed input handling (future)

## References

- [ECMA-376](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/) - Office Open XML File Formats
- [ISO/IEC 29500](https://www.iso.org/standard/71691.html) - Same spec, ISO version
- [Open XML SDK docs](https://docs.microsoft.com/en-us/office/open-xml/open-xml-sdk) - Microsoft's reference
- [python-docx](https://python-docx.readthedocs.io/) - Good API inspiration
- [Apache POI](https://poi.apache.org/) - Java reference implementation

## Consumers

- **rescribe** - Document conversion library (primary motivation)
- Standalone use for DOCX manipulation
- Report generation
- Document automation

## License

- **Rust code** (`crates/`): MIT OR Apache-2.0
- **Ecosystem resources** (`fixtures/`, `specs/`): CC0 1.0 (public domain)
