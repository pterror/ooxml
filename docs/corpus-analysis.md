# OOXML Corpus Analysis Report

This document summarizes findings from analyzing the NapierOne DOCX corpus and provides recommendations for development priorities.

## Corpus Overview

**Source**: [NapierOne Dataset](https://www.napierone.com/) - A diverse collection of real-world Office documents
**Analyzed**: 5,000 DOCX files
**Success Rate**: 99.98% (4,999 parsed successfully)
**Analysis Date**: 2026-01-20

## Feature Prevalence

Understanding what features appear most frequently in real documents helps prioritize development efforts.

### Document Structure

| Feature | Documents | Percentage | Priority |
|---------|-----------|------------|----------|
| Paragraphs | 5,000 | 100% | **Core** |
| Bold text | 4,784 | 95.7% | **Core** |
| Alignment | 4,040 | 80.8% | **Core** |
| Lists/Numbering | 3,599 | 72.0% | **High** |
| Text color | 3,353 | 67.1% | **High** |
| Tables | 3,252 | 65.1% | **High** |
| Hyperlinks | 2,898 | 58.0% | **High** |
| Italic text | 2,378 | 47.6% | **Core** |
| Images | 315 | 6.3% | **Medium** |

### Volume Statistics

- **Total paragraphs**: 1,395,629 (avg 279/doc)
- **Total tables**: 23,588 (avg 4.7/doc with tables)
- **Total images**: 1,239 (avg 3.9/doc with images)
- **Total hyperlinks**: 33,059 (avg 11.4/doc with links)

### Extremes Observed

| Metric | Maximum | Implication |
|--------|---------|-------------|
| Paragraphs per doc | 18,244 | Need streaming/lazy parsing |
| Tables per doc | 177 | Table-heavy documents exist |
| Images per doc | 294 | Media-heavy documents exist |
| Hyperlinks per doc | 1,211 | Link-heavy documents exist |
| Max table nesting | 1 | Nested tables are rare |

## Font Usage

Top fonts encountered (helps with font fallback strategies):

| Font | Documents | Percentage |
|------|-----------|------------|
| Arial | 2,753 | 55.1% |
| Calibri | 580 | 11.6% |
| MS Gothic | 347 | 6.9% |
| Times New Roman | 284 | 5.7% |
| Helvetica | 140 | 2.8% |
| Verdana | 128 | 2.6% |
| Tahoma | 98 | 2.0% |
| Segoe UI Symbol | 94 | 1.9% |
| Foundry Form Sans | 89 | 1.8% |
| Georgia | 61 | 1.2% |

**Recommendation**: Default to Arial/Calibri for best compatibility. Flag documents using uncommon fonts.

## Style Usage

Most referenced paragraph/character styles:

| Style | Documents | Type |
|-------|-----------|------|
| Hyperlink | 2,695 | Character |
| ListParagraph | 2,608 | Paragraph |
| Heading1 | 1,072 | Paragraph |
| Heading2 | 794 | Paragraph |
| Default | 597 | Paragraph |
| BodyText | 584 | Paragraph |
| NormalWeb | 532 | Paragraph |
| Heading3 | 527 | Paragraph |
| NoSpacing | 453 | Paragraph |
| Header | 403 | Paragraph |

**Recommendation**: Ensure robust style inheritance resolution. ListParagraph and Heading styles are critical.

## Edge Cases

Documents exhibiting unusual patterns that warrant special handling:

### By Frequency

| Edge Case | Documents | Severity | Notes |
|-----------|-----------|----------|-------|
| External hyperlinks | 2,898 | Info | Links to external URLs |
| Page breaks | 1,666 | Info | Explicit page break elements |
| Embedded images | 315 | Info | Images in document |
| Very large (>500 paragraphs) | 240 | Noteworthy | May need lazy loading |
| Many tables (>20) | 33 | Unusual | Complex table layouts |
| Very large fonts (>72pt) | 10 | Unusual | Decorative text |
| Many styles (>50) | 2 | Rare | Complex formatting |
| Many fonts (>15) | 1 | Rare | Mixed typography |

### By Severity

- **Rare** (9 docs): Extreme outliers requiring special handling
- **Unusual** (105 docs): Edge cases worth testing
- **Noteworthy** (172 docs): Common enough to warrant test fixtures
- **Info** (4,879 docs): Normal variation

## Validation Findings

### Warnings (non-blocking issues)

| Warning | Count | Implication |
|---------|-------|-------------|
| Unresolved style references | 16,120 | Documents reference undefined styles |
| Uncommon fonts | 756 | Documents use non-standard fonts |

**Key Insight**: Many documents reference styles that don't exist in their styles.xml. Our parser must handle missing style references gracefully.

### Errors

No blocking errors found - all 4,999 successfully parsed documents are structurally valid.

## WML Element Coverage

Coverage against ECMA-376 WordprocessingML elements:

- **Coverage**: 99.1% (105/106 tracked elements)
- **Missing**: `w:customXml` (content controls with custom XML)

### Most Common Elements (in 100% of documents)

Core document structure:
- `w:document`, `w:body`, `w:sectPr`
- `w:p`, `w:r`, `w:t`
- `w:pPr`, `w:rPr`
- `w:styles`, `w:style`

### Elements by Frequency Category

**Universal (100%)**:
`w:document`, `w:body`, `w:p`, `w:r`, `w:t`, `w:pPr`, `w:rPr`, `w:styles`, `w:style`, `w:tblPr`, `w:sectPr`

**Very Common (>80%)**:
`w:jc`, `w:spacing`, `w:sz`, `w:b`, `w:tbl`, `w:tr`, `w:tc`

**Common (50-80%)**:
`w:hyperlink`, `w:numPr`, `w:color`, `w:i`, `w:ind`

**Occasional (10-50%)**:
`w:drawing`, `w:u`, `w:highlight`, `w:tabs`

**Rare (<10%)**:
`w:footnotes`, `w:endnotes`, `w:comments`, `w:strike`

## Development Priorities

Based on corpus analysis, recommended implementation order:

### Tier 1: Core (Required for 95%+ of documents)

1. **Document structure**: `w:document`, `w:body`, `w:sectPr` ✅
2. **Paragraphs**: `w:p`, `w:pPr` ✅
3. **Runs**: `w:r`, `w:rPr`, `w:t` ✅
4. **Basic formatting**: `w:b`, `w:i`, `w:u`, `w:sz`, `w:color` ✅
5. **Styles**: `w:styles`, `w:style`, inheritance ✅
6. **Alignment**: `w:jc` ✅

### Tier 2: High Priority (Required for 60-95% of documents)

7. **Tables**: `w:tbl`, `w:tr`, `w:tc`, `w:tblPr` ✅
8. **Lists**: `w:numPr`, `w:numbering.xml` ✅
9. **Hyperlinks**: `w:hyperlink` ✅
10. **Spacing**: `w:spacing`, `w:ind` ⚠️ Partial
11. **Fonts**: `w:rFonts` ⚠️ Partial

### Tier 3: Medium Priority (Required for 10-60% of documents)

12. **Images**: `w:drawing`, `wp:inline`, `wp:anchor` ✅
13. **Page breaks**: `w:br` with `w:type="page"` ✅
14. **Tabs**: `w:tab`, `w:tabs` ❌
15. **Underline styles**: `w:u` with style attribute ⚠️ Partial
16. **Highlighting**: `w:highlight` ❌

### Tier 4: Lower Priority (Required for <10% of documents)

17. **Headers/Footers**: `w:hdr`, `w:ftr`, `w:headerReference` ❌
18. **Footnotes/Endnotes**: `w:footnote`, `w:endnote` ❌
19. **Comments**: `w:comment`, `w:commentRangeStart` ❌
20. **Bookmarks**: `w:bookmarkStart`, `w:bookmarkEnd` ❌
21. **Fields**: `w:fldSimple`, `w:fldChar` ❌
22. **Strikethrough**: `w:strike`, `w:dstrike` ❌

### Legend

- ✅ Implemented
- ⚠️ Partial implementation
- ❌ Not yet implemented

## Test Fixture Recommendations

Based on edge cases found, extract these documents as test fixtures:

### Must-Have Fixtures

```bash
# Large documents (stress testing)
ooxml-corpus extract /mnt/ssd/ooxml-corpora/napierone/DOCX/ ./fixtures/wml/large \
  --feature paragraphs-gt-1000 --max 3

# Table-heavy documents
ooxml-corpus extract /mnt/ssd/ooxml-corpora/napierone/DOCX/ ./fixtures/wml/tables \
  --feature tables --max 5

# Complex styling
ooxml-corpus extract /mnt/ssd/ooxml-corpora/napierone/DOCX/ ./fixtures/wml/styles \
  --interesting --max 5
```

### Recommended Test Categories

1. **Minimal documents**: <10 paragraphs, simple formatting
2. **Table variations**: Simple, merged cells, nested (if found)
3. **List variations**: Bulleted, numbered, multi-level
4. **Image variations**: Inline, anchored, multiple
5. **Hyperlink variations**: Internal, external, styled
6. **Style-heavy**: Many unique styles, deep inheritance
7. **Large documents**: >1000 paragraphs (lazy loading test)

## Failure Analysis

Only 1 document failed to parse:

| File | Error | Cause |
|------|-------|-------|
| 3278-docx.docx | Missing required part | References `/word/document2.xml` which doesn't exist |

**Recommendation**: This appears to be a corrupted or non-standard document. No action needed - our error handling correctly identifies the issue.

## Appendix: Running Your Own Analysis

```bash
# Full analysis with persistence
ooxml-corpus /path/to/docx/files \
  --features --edge-cases --validate \
  --db corpus.db --corpus my-corpus

# View statistics
ooxml-corpus stats corpus.db my-corpus

# Coverage report
ooxml-corpus coverage /path/to/docx/files

# Extract interesting fixtures
ooxml-corpus extract /path/to/docx/files ./fixtures \
  --interesting --max 10

# JSON output for further processing
ooxml-corpus /path/to/docx/files --features --json > analysis.json
```

## References

- [ECMA-376 Standard](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/)
- [NapierOne Dataset](https://www.napierone.com/)
- [Open XML SDK Documentation](https://docs.microsoft.com/en-us/office/open-xml/open-xml-sdk)
