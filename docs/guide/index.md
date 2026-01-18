# Introduction

`ooxml` is a Rust library for reading and writing Office Open XML formats (DOCX, XLSX, PPTX).

## Features

- **Type-safe** - Structs for every XML element
- **Roundtrip fidelity** - Unknown elements preserved
- **Lazy parsing** - Only parse what you need
- **Spec-driven** - Built against ECMA-376

## Crates

| Crate | Description |
|-------|-------------|
| `ooxml` | Core: OPC packaging, relationships, content types |
| `ooxml-wml` | WordprocessingML (Word documents) |

## Quick Example

```rust
use ooxml_wml::Document;

// Read a document
let doc = Document::open("input.docx")?;
for para in doc.body().paragraphs() {
    println!("{}", para.text());
}

// Modify and save
let mut doc = Document::open("input.docx")?;
doc.body_mut().add_paragraph().add_run().set_text("Hello!");
doc.save("output.docx")?;
```
