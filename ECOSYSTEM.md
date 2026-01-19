# OOXML Ecosystem Resources

## Context: Rhizome Ecosystem

This library is part of the [Rhizome](https://rhizome-lab.github.io/) ecosystem:

```
Rhizome (data manipulation ecosystem)
└── Cambium (pipeline orchestrator for data conversion)
    └── rescribe (document conversion, pandoc-style)
        └── ooxml (OOXML format support) ← you are here
```

The ooxml library provides low-level DOCX/XLSX/PPTX support. rescribe will use it to implement `rescribe-read-docx` and `rescribe-write-docx` for document conversion pipelines.

## The Problem

The ECMA-376 spec is 5000+ pages across multiple parts. Implementing OOXML support in any language requires significant effort, and most of that effort is duplicated across every implementation (python-docx, Apache POI, Open XML SDK, docx4j, etc.).

1. **The spec is massive** - ECMA-376 has 4 parts, thousands of pages
2. **XSD schemas are incomplete** - They define structure but not semantics
3. **Real-world deviates from spec** - Microsoft Office, LibreOffice, Google Docs all have quirks
4. **No shared test fixtures** - Every project creates their own
5. **Knowledge is scattered** - Gotchas live in blog posts, Stack Overflow, source code comments

## What We're Building

This repo contains both:
- **Rust implementation** (`crates/`) - MIT/Apache-2.0 licensed
- **Ecosystem resources** (`fixtures/`, `specs/`) - CC0 (public domain)

The Rust code dogfoods the ecosystem resources. As we implement features, we build fixtures and document the spec. Anyone can use the fixtures and specs, regardless of language.

### Repository Structure

```
ooxml/
├── crates/                    # Rust implementation (MIT/Apache-2.0)
│   ├── ooxml/                 # Core: OPC packaging, relationships
│   ├── ooxml-wml/             # WordprocessingML (Word)
│   ├── ooxml-sml/             # SpreadsheetML (Excel) - future
│   ├── ooxml-pml/             # PresentationML (PowerPoint) - future
│   └── ooxml-codegen/         # Generate types from specs
│
├── fixtures/                  # Test files (CC0)
│   ├── wml/                   # Word documents
│   ├── sml/                   # Excel spreadsheets
│   └── pml/                   # PowerPoint presentations
│
├── specs/                     # Machine-readable spec (CC0)
│   ├── elements/              # Element definitions
│   ├── enums/                 # Enumeration values
│   ├── units/                 # Unit conversions
│   ├── relationships/         # Relationship types
│   └── deviations/            # Real-world quirks
│
└── tests/                     # Language-agnostic test suite (CC0)
```

---

## 1. Fixture Strategy: Corpus Mining

**Key insight:** Hand-crafted "real-world" fixtures aren't real. A document edited by 5 people over 3 years, round-tripped through Word → Google Docs → LibreOffice, with tracked changes from a legal review - you can't fake that. The cruft *is* the feature.

Our approach: **mine real corpora, distill edge cases into minimal fixtures.**

### The Workflow

```
External Corpora (NapierOne, docx-corpus, OfficeDissector)
                        │
                        ▼
        Run parser against thousands of real files
                        │
                        ▼
    Detect: failures, warnings, unknown elements, weird patterns
                        │
                        ▼
    Minimize: extract smallest XML that reproduces the behavior
                        │
                        ▼
    Document: create fixture + manifest explaining what it tests
                        │
                        ▼
    fixtures/wml/edge-cases/discovered-xyz.docx
```

### External Corpora (don't copy, reference)

| Corpus | Size | License | What it is |
|--------|------|---------|------------|
| [docx-corpus](https://github.com/superdoc-dev/docx-corpus) | Large | MIT | Common Crawl scrape - real web documents |
| [NapierOne](https://registry.opendata.aws/napierone/) | 5K each DOCX/XLSX/PPTX | Free w/ attribution | Modern academic dataset |
| [OfficeDissector](https://github.com/grierforensics/officedissector) | 600MB | Check repo | Security research, includes "wild" docs |
| [OPF Format Corpus](https://github.com/openpreserve/format-corpus) | Varied | CC0 | Digital preservation |
| [Open-Xml-PowerTools](https://github.com/OfficeDev/Open-Xml-PowerTools) | 555 files | MIT | Microsoft's test files (archived) |
| [LibreOffice test-files](https://github.com/niccokunzmann/libreoffice_test-files) | Varied | MPL-ish | Includes ooxml-strict conformance |

**Note:** There is no official ECMA-376 conformance test suite. Everyone rolls their own.

### Corpus Analysis Tooling

Since we're implementing a parser, we can query real corpora structurally:

```rust
// Potential tooling in ooxml-codegen or separate crate
find_all_element_paths()           // What element nesting exists in the wild?
find_attribute_values("w:val")     // What values does w:val actually take?
find_documents_with("w:sdt")       // Who uses content controls?
count_by_producer()                // Word 2019 vs LibreOffice vs Google
extract_minimal_reproducer(doc)    // Shrink a failing doc to smallest case
```

This turns corpus analysis into **research**: "What does OOXML look like in practice?"

### In-Repo Fixtures (minimal, documented)

The fixtures we *do* commit are:
1. **Minimal** - Smallest file that tests one specific thing
2. **Documented** - Manifest explains why it exists, what edge case it covers
3. **Distilled** - Often derived from corpus mining, not hand-crafted

```
fixtures/
  wml/
    basic/                        # Fundamental features
      simple-paragraph.docx
      manifest.yaml               # Documents the fixture
    edge-cases/                   # Discovered via corpus mining
      rsid-on-bookmark.docx       # Found in 3% of NapierOne docs
      nested-sdt-in-hyperlink.docx
    producers/                    # Same content, different tools
      hello-word-2019.docx
      hello-libreoffice-7.docx
      hello-google-docs.docx
  sml/
    ...
  pml/
    ...
```

### Manifest Format

Each fixture has a manifest:

```yaml
# fixtures/wml/edge-cases/rsid-on-bookmark/manifest.yaml
fixture:
  file: rsid-on-bookmark.docx
  discovered_from: NapierOne corpus
  prevalence: "~3% of documents"

  description: |
    Bookmarks with rsid attributes on both bookmarkStart and bookmarkEnd.
    The spec says rsid is optional, but some tools require it for roundtrip.

  tests:
    - Parser doesn't crash on rsid attributes
    - Bookmark is correctly identified
    - rsid is preserved on roundtrip

  xml_snippet: |
    <w:bookmarkStart w:id="0" w:name="ref1" w:rsidR="00A4722C"/>
    ...
    <w:bookmarkEnd w:id="0" w:rsidR="00A4722C"/>

  spec_reference: ECMA-376 Part 1, Section 17.13.6

  related_deviation: specs/deviations/rsid-everywhere.yaml
```

---

## 2. Machine-Readable Spec

Not just schemas - the full semantic information from ECMA-376.

### Element Database

```yaml
# specs/elements/wml/paragraph.yaml
element:
  name: p
  namespace: http://schemas.openxmlformats.org/wordprocessingml/2006/main
  prefix: w
  spec_reference: ECMA-376 Part 1, Section 17.3.1.22

  description: |
    Paragraph - the fundamental block-level content container.
    Contains runs of text with formatting, and paragraph-level properties.

  parents:
    - body
    - tc        # table cell
    - txbxContent  # text box
    - footnote
    - endnote
    - comment
    - header
    - footer

  children:
    - element: pPr
      min: 0
      max: 1
      description: Paragraph properties (must be first child if present)
    - element: r
      min: 0
      max: unbounded
    - element: hyperlink
      min: 0
      max: unbounded
    - element: bookmarkStart
      min: 0
      max: unbounded
    - element: bookmarkEnd
      min: 0
      max: unbounded

  attributes: []  # <w:p> has no attributes in standard usage

  examples:
    - description: Simple paragraph with text
      xml: |
        <w:p>
          <w:r>
            <w:t>Hello, World!</w:t>
          </w:r>
        </w:p>
    - description: Paragraph with style
      xml: |
        <w:p>
          <w:pPr>
            <w:pStyle w:val="Heading1"/>
          </w:pPr>
          <w:r>
            <w:t>Chapter 1</w:t>
          </w:r>
        </w:p>

  notes:
    - Runs and hyperlinks can be interleaved in any order
    - Empty paragraphs are valid and common (used for spacing)
```

### Enumeration Database

```yaml
# specs/enums/wml/st-jc.yaml
enumeration:
  name: ST_Jc
  spec_reference: ECMA-376 Part 1, Section 17.18.44

  description: Horizontal alignment / justification

  used_by:
    - element: jc
      attribute: val
      context: Paragraph alignment (w:pPr/w:jc)
    - element: jc
      attribute: val
      context: Table alignment (w:tblPr/w:jc)

  values:
    - value: left
      description: Align left (default for LTR languages)
    - value: center
      description: Center alignment
    - value: right
      description: Align right (default for RTL languages)
    - value: both
      description: Justified (both edges aligned)
    - value: distribute
      description: Distributed alignment (CJK typography)
    - value: start
      description: Align to start edge (language-aware)
    - value: end
      description: Align to end edge (language-aware)

  default: left  # for LTR documents

  notes:
    - "start" and "end" are preferred over "left" and "right" for bidi support
```

### Units and Conversions

```yaml
# specs/units/emu.yaml
unit:
  name: EMU
  full_name: English Metric Unit
  spec_reference: ECMA-376 Part 1, Section 20.1.2.1

  description: |
    Base unit for measurements in DrawingML.
    Used for image dimensions, positions, etc.

  conversions:
    - to: inch
      multiply_by: 0.00001102362205  # 1/914400
    - to: cm
      multiply_by: 0.000028         # 1/360000
    - to: point
      multiply_by: 0.00079375       # 1/12700
    - to: pixel_96dpi
      multiply_by: 0.00010583333    # 96/914400

  constants:
    per_inch: 914400
    per_cm: 360000
    per_point: 12700

  used_in:
    - DrawingML (a:)
    - Inline images (wp:extent)
    - Anchored images (wp:positionH, wp:positionV)
```

```yaml
# specs/units/half-point.yaml
unit:
  name: half-point
  spec_reference: ECMA-376 Part 1, Section 17.18.89

  description: |
    Unit for font sizes in WordprocessingML.
    A half-point is 1/144 of an inch (1/2 of a typographic point).

  conversions:
    - to: point
      multiply_by: 0.5
    - to: inch
      multiply_by: 0.006944444  # 1/144

  examples:
    - xml_value: "24"
      meaning: "12pt font"
    - xml_value: "48"
      meaning: "24pt font"

  used_in:
    - w:sz (font size)
    - w:szCs (complex script font size)
```

### Relationship Types

```yaml
# specs/relationships/wml.yaml
relationships:
  - type: http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument
    short_name: officeDocument
    description: Main document part
    typical_target: word/document.xml
    from: package root (_rels/.rels)

  - type: http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles
    short_name: styles
    description: Style definitions
    typical_target: word/styles.xml
    from: document part (word/_rels/document.xml.rels)

  - type: http://schemas.openxmlformats.org/officeDocument/2006/relationships/image
    short_name: image
    description: Embedded image
    typical_target: word/media/image1.png
    from: document part
    target_mode: Internal

  - type: http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink
    short_name: hyperlink
    description: External hyperlink URL
    typical_target: https://example.com
    from: document part
    target_mode: External
```

### Deviation Documentation

What the spec says vs. what implementations actually do.

```yaml
# specs/deviations/microsoft-word.yaml
deviations:
  - id: empty-toggle-elements
    spec_says: |
      Toggle properties like <w:b> (bold) should use w:val attribute.
      <w:b w:val="true"/> or <w:b w:val="1"/> means bold.
      <w:b w:val="false"/> or <w:b w:val="0"/> means not bold.

    word_does: |
      Omits w:val entirely when true: <w:b/>
      This is technically valid (default is true when omitted) but
      many implementations don't handle it correctly.

    recommendation: |
      Treat <w:b/> (no val attribute) as equivalent to <w:b w:val="true"/>

    affected_elements:
      - w:b (bold)
      - w:i (italic)
      - w:u (underline)
      - w:strike (strikethrough)

    spec_reference: ECMA-376 Part 1, Section 17.7.3

  - id: rsid-attributes
    spec_says: |
      RSID (Revision Save ID) attributes track document editing sessions.
      They are optional.

    word_does: |
      Adds rsid attributes to almost every element:
      w:rsidR, w:rsidRPr, w:rsidRDefault, w:rsidP, etc.
      These bloat the XML significantly but are harmless.

    recommendation: |
      Ignore rsid* attributes when reading.
      Don't generate them when writing (Word will add them on edit).

    affected_elements: Almost all elements

  - id: proofErr-elements
    spec_says: Grammar/spelling error markers are optional.

    word_does: |
      Inserts <w:proofErr w:type="spellStart"/> and <w:proofErr w:type="spellEnd"/>
      around words it thinks are misspelled.

    recommendation: |
      Ignore these elements when reading.
      Don't generate them when writing.
```

---

## 3. Language-Agnostic Test Suite

JSON-based test cases any implementation can run.

```json
{
  "test_suite": "wml-paragraph-parsing",
  "version": "1.0",
  "tests": [
    {
      "id": "simple-paragraph",
      "description": "Parse a simple paragraph with one run",
      "input_file": "fixtures/wml/text/simple-paragraph.docx",
      "expected": {
        "body": {
          "paragraphs": [
            {
              "runs": [
                {
                  "text": "Hello, World!",
                  "properties": {
                    "bold": false,
                    "italic": false
                  }
                }
              ]
            }
          ]
        }
      }
    }
  ]
}
```

---

## 4. Codegen Tools

Generate language-specific types from the machine-readable spec.

**Inputs:**
- Element database (YAML)
- Enumeration database (YAML)

**Outputs:**
- Rust structs and enums
- TypeScript interfaces
- Python dataclasses
- JSON Schema
- XSD (for validation)

This lives in `crates/ooxml-codegen/`. See also [Liana](https://rhizome-lab.github.io/) for API bindings IR.

---

## Prior Art / Inspiration

**Test Suites:**
- [JSON Schema Test Suite](https://github.com/json-schema-org/JSON-Schema-Test-Suite) - Language-agnostic conformance tests
- [CSS Test Suites](https://test.csswg.org/) - W3C conformance testing
- [Unicode Test Data](https://unicode.org/Public/) - Character encoding edge cases

**OOXML Implementations:**
- [python-docx](https://github.com/python-openxml/python-docx) - Python, MIT, good API design
- [Apache POI](https://github.com/apache/poi) - Java, Apache 2.0, very comprehensive
- [docx4j](https://github.com/plutext/docx4j) - Java, Apache 2.0
- [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) - .NET, MIT, Microsoft's official

**Document Corpora:**
- [docx-corpus](https://github.com/superdoc-dev/docx-corpus) - Real documents from Common Crawl
- [NapierOne](https://registry.opendata.aws/napierone/) - 500K+ files, academic dataset
- [OfficeDissector](https://github.com/grierforensics/officedissector) - Security-focused corpus
- [Govdocs1](https://digitalcorpora.org/corpora/file-corpora/files/) - 1M government documents (older)

**The Gap:** There is no official ECMA-376/ISO-29500 conformance test suite. OASIS has XML conformance tests, but nothing OOXML-specific. This is a significant gap we aim to help fill.

---

## Contributing

We welcome contributions of:
- **Corpus findings** - "I ran against NapierOne and found X pattern in Y% of files"
- **Minimal reproducers** - Distilled fixtures that test specific edge cases
- **Spec transcriptions** - YAML element/enum definitions from ECMA-376
- **Deviation documentation** - Real-world quirks (Word does X, LibreOffice does Y)
- **Producer comparisons** - Same document saved by different tools
- **Codegen templates** - Generate types for other languages

---

## License

- **Rust code** (`crates/`): MIT OR Apache-2.0
- **Ecosystem resources** (`fixtures/`, `specs/`, `tests/`): CC0 1.0 (public domain)

See LICENSE-MIT, LICENSE-APACHE, and fixtures/LICENSE for details.
