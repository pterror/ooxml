# ooxml - Rust OOXML Library

High-quality Rust library for reading/writing Office Open XML formats (DOCX, XLSX, PPTX).

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
crates/
  ooxml/              # Core: OPC packaging, shared types
  ooxml-wml/          # WordprocessingML (Word documents)
  ooxml-sml/          # SpreadsheetML (Excel) - future
  ooxml-pml/          # PresentationML (PowerPoint) - future
  ooxml-dml/          # DrawingML (shared graphics) - future
```

## Scope

### v0.1 - Core + Word Basics

**ooxml (core):**
- [ ] OPC packaging (ZIP read/write)
- [ ] Relationships (.rels files)
- [ ] Content types ([Content_Types].xml)
- [ ] Core properties (docProps/core.xml)
- [ ] App properties (docProps/app.xml)

**ooxml-wml (WordprocessingML):**
- [ ] Document structure (document.xml)
- [ ] Paragraphs (`<w:p>`)
- [ ] Runs (`<w:r>`) and text (`<w:t>`)
- [ ] Basic formatting: bold, italic, underline, strikethrough
- [ ] Font, size, color
- [ ] Paragraph properties: alignment, spacing, indentation
- [ ] Headings (via paragraph styles)
- [ ] Lists (numbering definitions + abstract numbering)
- [ ] Tables (basic: rows, cells, simple borders)
- [ ] Hyperlinks
- [ ] Images (embedded in word/media/)
- [ ] Styles (styles.xml) - read and apply
- [ ] Page breaks, section breaks

### v0.2 - Extended Word

- [ ] Headers and footers
- [ ] Footnotes and endnotes
- [ ] Table of contents (read)
- [ ] Bookmarks
- [ ] Complex tables (merged cells, nested tables)
- [ ] Text boxes
- [ ] Tabs and tab stops
- [ ] Borders and shading

### v0.3 - Advanced Word

- [ ] Track changes (revisions)
- [ ] Comments
- [ ] Form fields and content controls
- [ ] Custom XML parts
- [ ] Math (OMML)
- [ ] SmartArt (limited)
- [ ] Charts (limited)

### Future

- [ ] ooxml-sml: Excel support
- [ ] ooxml-pml: PowerPoint support
- [ ] ooxml-dml: Full DrawingML

## Dependencies

```toml
[dependencies]
zip = "2"              # ZIP archive handling
quick-xml = "0.36"     # XML parsing/writing
thiserror = "2"        # Error handling

[dev-dependencies]
insta = "1"            # Snapshot testing
```

## API Design (Draft)

```rust
// Reading
let doc = ooxml_wml::Document::open("input.docx")?;
for para in doc.body().paragraphs() {
    println!("Style: {:?}", para.style());
    for run in para.runs() {
        println!("  Text: {}", run.text());
        println!("  Bold: {}", run.properties().bold());
    }
}

// Writing
let mut doc = ooxml_wml::Document::new();
let mut para = doc.body_mut().add_paragraph();
para.set_style("Heading1");
para.add_run().set_text("Hello, World!");
doc.save("output.docx")?;

// Roundtrip (preserves unknown elements)
let mut doc = ooxml_wml::Document::open("input.docx")?;
doc.body_mut().paragraphs_mut().next().unwrap()
    .add_run().set_text(" - modified");
doc.save("output.docx")?;
```

## Testing Strategy

1. **Unit tests** - Individual element parsing/serialization
2. **Roundtrip tests** - Open → save → compare (byte-level or structural)
3. **Fixture tests** - Real DOCX files from various sources
4. **Snapshot tests** - Insta for XML output verification
5. **Fuzz tests** - Malformed input handling

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
