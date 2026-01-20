# OOXML Implementation Status

This document tracks what's implemented, partially implemented, and missing in the ooxml-wml crate.

## Reading (Parsing)

### Document Structure

| Element | Status | Notes |
|---------|--------|-------|
| `w:document` | ✅ | Root element |
| `w:body` | ✅ | Document body |
| `w:sectPr` | ✅ | Section properties (page size, margins, orientation) |

### Block-Level Content

| Element | Status | Notes |
|---------|--------|-------|
| `w:p` | ✅ | Paragraphs |
| `w:tbl` | ✅ | Tables |
| `w:sdt` | ❌ | Structured document tags (content controls) |
| `w:customXml` | ❌ | Custom XML blocks |

### Paragraph Content

| Element | Status | Notes |
|---------|--------|-------|
| `w:r` | ✅ | Runs (text spans) |
| `w:hyperlink` | ✅ | Hyperlinks |
| `w:bookmarkStart` | ✅ | Bookmark anchors |
| `w:bookmarkEnd` | ✅ | Bookmark anchors |
| `w:commentRangeStart` | ✅ | Comment anchors |
| `w:commentRangeEnd` | ✅ | Comment anchors |
| `w:fldSimple` | ✅ | Simple fields |
| `w:fldChar` | ✅ | Complex fields (begin/separate/end markers) |

### Run Content

| Element | Status | Notes |
|---------|--------|-------|
| `w:t` | ✅ | Text |
| `w:tab` | ✅ | Tab characters (converted to `\t`) |
| `w:br` | ✅ | Line breaks (converted to `\n`) |
| `w:br w:type="page"` | ✅ | Page breaks |
| `w:sym` | ✅ | Symbol characters |
| `w:drawing` | ✅ | DrawingML container |
| `w:pict` | ❌ | VML pictures (legacy) |
| `w:object` | ❌ | Embedded objects |

### Paragraph Properties (`w:pPr`)

| Property | Status | Notes |
|----------|--------|-------|
| `w:pStyle` | ✅ | Paragraph style reference |
| `w:jc` | ✅ | Alignment (left, center, right, justify) |
| `w:ind` | ✅ | Indentation (left, right, firstLine, hanging) |
| `w:spacing` | ✅ | Spacing (before, after, line) |
| `w:numPr` | ✅ | List numbering |
| `w:pBdr` | ✅ | Paragraph borders (top, bottom, left, right, between, bar) |
| `w:shd` | ✅ | Shading/background |
| `w:tabs` | ✅ | Tab stop definitions |
| `w:outlineLvl` | ✅ | Outline level (0-9) |
| `w:keepNext` | ✅ | Keep with next paragraph |
| `w:keepLines` | ✅ | Keep lines together |
| `w:pageBreakBefore` | ✅ | Page break before |
| `w:widowControl` | ✅ | Widow/orphan control |

### Run Properties (`w:rPr`)

| Property | Status | Notes |
|----------|--------|-------|
| `w:rStyle` | ✅ | Character style reference |
| `w:b` | ✅ | Bold |
| `w:i` | ✅ | Italic |
| `w:u` | ✅ | Underline with styles (single, double, wavy, dotted, etc.) |
| `w:strike` | ✅ | Strikethrough |
| `w:dstrike` | ✅ | Double strikethrough |
| `w:color` | ✅ | Text color |
| `w:sz` | ✅ | Font size (in half-points) |
| `w:rFonts` | ✅ | Font (ascii attribute only) |
| `w:highlight` | ✅ | Highlight color (16 standard colors) |
| `w:vertAlign` | ✅ | Superscript/subscript |
| `w:caps` | ✅ | All capitals |
| `w:smallCaps` | ✅ | Small capitals |
| `w:vanish` | ✅ | Hidden text |
| `w:shd` | ✅ | Shading/background |

### Section Properties (`w:sectPr`)

| Property | Status | Notes |
|----------|--------|-------|
| `w:pgSz` | ✅ | Page size (width, height, orientation) |
| `w:pgMar` | ✅ | Page margins (top, bottom, left, right, header, footer, gutter) |
| `w:cols` | ✅ | Column definitions |
| `w:docGrid` | ✅ | Document grid settings |
| `w:type` | ✅ | Section type (continuous, nextPage, etc.) |

### Table Elements

| Element | Status | Notes |
|---------|--------|-------|
| `w:tbl` | ✅ | Table container |
| `w:tr` | ✅ | Table row |
| `w:tc` | ✅ | Table cell |
| `w:tblPr` | ✅ | Table properties (width, justification, indent, layout, borders, shading) |
| `w:tblGrid` | ✅ | Column definitions |
| `w:gridCol` | ✅ | Column width |
| `w:trPr` | ✅ | Row properties (height, header, cantSplit) |
| `w:tcPr` | ✅ | Cell properties (width, borders, shading, merge, alignment) |
| `w:tblBorders` | ✅ | Table borders |
| `w:tcBorders` | ✅ | Cell borders (top, bottom, left, right, insideH, insideV) |
| `w:gridSpan` | ✅ | Horizontal cell merge |
| `w:vMerge` | ✅ | Vertical cell merge (restart/continue) |
| `w:shd` | ✅ | Cell shading (fill, pattern) |
| `w:tcW` | ✅ | Cell width (dxa, pct, auto) |
| `w:vAlign` | ✅ | Cell vertical alignment |
| `w:trHeight` | ✅ | Row height (exact, atLeast, auto) |
| `w:tblHeader` | ✅ | Header row (repeats on each page) |
| `w:tblW` | ✅ | Table width |
| `w:tblInd` | ✅ | Table indent |
| `w:tblLayout` | ✅ | Table layout (fixed, autofit) |
| `w:jc` (in tblPr) | ✅ | Table justification (left, center, right) |

### Images (DrawingML)

| Element | Status | Notes |
|---------|--------|-------|
| `wp:inline` | ✅ | Inline images |
| `wp:anchor` | ✅ | Anchored/floating images with text wrapping |
| `a:blip` | ✅ | Image reference |
| `wp:extent` | ✅ | Image dimensions |
| `wp:docPr` | ✅ | Image description/alt text |

### Document Parts

| Part | Status | Notes |
|------|--------|-------|
| `word/document.xml` | ✅ | Main document |
| `word/styles.xml` | ✅ | Style definitions |
| `word/numbering.xml` | ✅ | List definitions |
| `word/_rels/document.xml.rels` | ✅ | Document relationships |
| `word/header*.xml` | 🔶 | Header references parsed, content parts not loaded |
| `word/footer*.xml` | 🔶 | Footer references parsed, content parts not loaded |
| `word/footnotes.xml` | 🔶 | Footnote references parsed, content parts not loaded |
| `word/endnotes.xml` | 🔶 | Endnote references parsed, content parts not loaded |
| `word/comments.xml` | ❌ | Comments |
| `word/settings.xml` | ❌ | Document settings |

## Writing (Serialization)

### DocumentBuilder API

| Feature | Status | Notes |
|---------|--------|-------|
| `add_paragraph()` | ✅ | Basic paragraphs |
| `add_heading()` | ✅ | Heading paragraphs |
| `add_formatted_text()` | ✅ | Bold, italic, underline, strike, color |
| `add_page_break()` | ✅ | Page breaks |
| `add_hyperlink()` | ✅ | Hyperlinks |
| `add_image()` | ✅ | Inline images |
| `add_table()` | ✅ | Basic tables |
| `add_list()` | ✅ | Bulleted and numbered lists |

### Run Properties Written

| Property | Status |
|----------|--------|
| Bold | ✅ |
| Italic | ✅ |
| Underline | ✅ |
| Underline styles | ✅ |
| Strikethrough | ✅ |
| Double strikethrough | ✅ |
| Font size | ✅ |
| Font family | ✅ |
| Text color | ✅ |
| Highlight | ✅ |
| Superscript/subscript | ✅ |
| All caps | ✅ |
| Small caps | ✅ |

### Section Properties Written

| Property | Status |
|----------|--------|
| Page size | ✅ |
| Page orientation | ✅ |
| Page margins | ✅ |

## Priority Gaps

Based on [corpus analysis](./corpus-analysis.md), these are the most impactful missing features:

### High Priority (affects 20%+ of documents)

1. ~~**Underline styles**~~ ✅ Now supports all 17 underline styles
2. ~~**Highlight colors**~~ ✅ Now supports all 16 standard highlight colors
3. ~~**Table cell properties**~~ ✅ Now supports borders, shading, width, and merging
4. ~~**Table properties**~~ ✅ Now supports width, justification, indent, layout, borders, shading
5. ~~**Row properties**~~ ✅ Now supports height, header rows, and cantSplit
6. ~~**Table grid**~~ ✅ Now supports column width definitions

### Medium Priority (affects 5-20% of documents)

7. **Headers/Footers** - 🔶 References parsed, content parts pending
8. ~~**Anchored images**~~ ✅ Now supports floating images with text wrapping
9. ~~**Tab stops**~~ ✅ Now supports custom tab stop definitions

### Lower Priority (affects <5% of documents)

10. **Footnotes/Endnotes** - 🔶 References parsed, content parts pending
11. **Comments** - Rarely present in final documents
12. ~~**Superscript/Subscript**~~ ✅ Now implemented
13. **Content controls** - Enterprise/form documents

## Roundtrip Preservation

Unknown elements and attributes **are preserved** during roundtrip via position-tracking:

- `PositionedNode` - stores unknown XML elements with their original position
- `PositionedAttr` - stores unknown attributes with their original position
- Elements are serialized back in their original order relative to known elements

This enables lossless roundtripping of documents containing elements we don't explicitly parse.
