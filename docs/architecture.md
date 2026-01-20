# OOXML Architecture & Design Philosophy

## What We Actually Do

At its core, this library does one thing: **parse XML into Rust structs and serialize them back**.

```
DOCX file (ZIP archive)
    └── word/document.xml (XML)
            ↓ parse
        Rust structs (Document, Paragraph, Run, etc.)
            ↓ serialize
    └── word/document.xml (XML)
DOCX file (ZIP archive)
```

That's it. No rendering, no layout, no fonts, no pagination. Just data structure manipulation.

## Why OOXML Isn't As Scary As It Seems

The ECMA-376 specification is ~6,000 pages. Intimidating! But here's the reality:

### What the spec covers that we DON'T need:

- **Rendering semantics** - How to draw text, calculate line breaks, pagination
- **Font metrics** - Glyph positioning, kerning, ligatures
- **Layout algorithms** - Column balancing, widow/orphan control
- **UI behaviors** - Selection, editing, cursor movement
- **Compatibility modes** - Emulating Word 97/2003/2007 quirks
- **Calculation engines** - Spreadsheet formulas, presentation animations
- **Legacy formats** - VML, OLE, binary blobs

### What we DO need (and it's manageable):

1. **ZIP/OPC packaging** - Standard ZIP with `[Content_Types].xml` and `.rels` files
2. **XML structure** - Elements like `<w:p>`, `<w:r>`, `<w:t>` with predictable nesting
3. **Relationships** - How parts reference each other (images, hyperlinks, styles)
4. **Style inheritance** - How paragraph/character styles cascade
5. **Numbering** - How lists work (abstract definitions + concrete instances)

### The 80/20 reality

From our corpus analysis:
- **95%** of documents need only: paragraphs, runs, basic formatting, tables
- **80%** add: lists, hyperlinks, simple styles
- **20%** use: images, headers/footers
- **<5%** use: footnotes, comments, fields, content controls

Most documents are simple. The spec is large because it must handle everything Microsoft Word can produce, but real-world documents use a tiny fraction.

## What Other Ecosystems Do

| Library | Language | Approach | Rendering? |
|---------|----------|----------|------------|
| **python-docx** | Python | Read/write Word docs | No |
| **Apache POI** | Java | Full Office suite support | No |
| **docx4j** | Java | Comprehensive Word support | No* |
| **Open XML SDK** | .NET | Microsoft's official library | No |
| **docx** | JavaScript | Read/write Word docs | No |
| **pandoc** | Haskell | Document conversion | No |

*docx4j has optional PDF export via XSL-FO, but that's conversion, not true rendering.

**Key insight**: None of them render documents. Rendering is a completely separate problem requiring:
- Font rasterization (FreeType, HarfBuzz)
- Text shaping for complex scripts
- Layout engines (like WebKit or Pango)
- Graphics backends (Cairo, Skia)

That's why Word, LibreOffice, and Google Docs are massive applications - the rendering is the hard part.

## Why Rust Didn't Have This Before

1. **Ecosystem maturity** - Rust's XML libraries (quick-xml) are newer
2. **Target audience** - Rust attracts systems programmers, not enterprise doc processors
3. **Enterprise adoption** - Companies processing Office docs use Python/Java/.NET
4. **Tedious work** - Someone has to read 6,000 pages of spec and implement it
5. **No pressing need** - Existing libraries in other languages work fine

The Rust ecosystem has excellent libraries for many domains (web, crypto, systems) but document processing wasn't a priority until now.

## Our Design Decisions

### 1. Typed over stringly

Every XML element gets a Rust struct. No `HashMap<String, String>` for attributes.

```rust
// Yes:
pub struct RunProperties {
    pub bold: bool,
    pub italic: bool,
    pub size: Option<u32>,
}

// No:
pub type RunProperties = HashMap<String, String>;
```

### 2. Lazy when possible

We don't parse the entire document upfront. Styles are parsed on demand. Images are loaded when requested.

### 3. Preserve unknown data (goal)

Currently a gap: we should store unparsed XML to enable roundtrip fidelity. Documents should survive read→write without losing features we don't understand.

### 4. Spec-driven naming

Our struct and field names match OOXML terminology from ECMA-376. `Run` not `TextSpan`. `Paragraph` not `Block`.

## The Real Complexity

The hard parts of OOXML aren't the XML parsing:

1. **Producer quirks** - Word, LibreOffice, Google Docs all produce slightly different XML
2. **Version differences** - OOXML has evolved (ECMA-376 1st through 5th editions)
3. **Style resolution** - Inheritance chains can be deep and circular references exist
4. **Numbering** - The list system is genuinely complex (abstract nums, overrides, levels)
5. **Relationships** - Everything references everything else via relationship IDs
6. **DrawingML** - Images use a completely separate markup language (DrawingML/a:)

But these are solvable with careful engineering. No magic required.

## What This Enables

With a solid parsing/writing library, you can build:

- **Document converters** - DOCX → HTML, Markdown, PDF
- **Content extractors** - Pull text, images, metadata
- **Document generators** - Create DOCX from templates
- **Diff tools** - Compare document versions
- **Search indexers** - Extract and index content
- **Validation tools** - Check document structure
- **Transformation pipelines** - Batch modify documents

None of these need rendering. They just need structured access to document content.
