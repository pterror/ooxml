# SML Attribute Usage Thresholds

Based on corpus analysis of 500 XLSX files (4.8M rows, 41M cells, 1.25M formulas).

## Proposed Tiers

| Tier | Threshold | Behavior |
|------|-----------|----------|
| **CORE** | >50% | Always parse |
| **COMMON** | 10-50% | Default on, can disable |
| **RARE** | <10% | Default off, opt-in |

## Row Attributes (`<row>`)

| Attribute | Usage | Tier | Notes |
|-----------|-------|------|-------|
| `r` | 100% | CORE | Row number (required) |
| `x14ac:dyDescent` | 93.7% | CORE | Extension attr, very common |
| `spans` | 77.9% | CORE | Cell span hints |
| `customFormat` | 43.1% | COMMON | Custom formatting flag |
| `s` | 43.1% | COMMON | Style index |
| `hidden` | 21.9% | COMMON | Hidden row |
| `ht` | 10.5% | COMMON | Row height |
| `customHeight` | 8.0% | RARE | Custom height flag |
| `outlineLevel` | 0.24% | RARE | Outline grouping |
| `thickBot` | 0.17% | RARE | Thick bottom border |
| `thickTop` | 0.04% | RARE | Thick top border |
| `collapsed` | 0% | RARE | Outline collapsed (not seen) |
| `ph` | 0% | RARE | Phonetic (not seen) |

## Cell Attributes (`<c>`)

| Attribute | Usage | Tier | Notes |
|-----------|-------|------|-------|
| `r` | 100% | CORE | Cell reference (required) |
| `s` | 89.2% | CORE | Style index |
| `t` | 31.4% | COMMON | Cell type |
| `cm` | 0% | RARE | Cell metadata (not seen) |
| `vm` | 0% | RARE | Value metadata (not seen) |
| `ph` | 0% | RARE | Phonetic (not seen) |

## Formula Attributes (`<f>`)

| Attribute | Usage | Tier | Notes |
|-----------|-------|------|-------|
| `t` | 49.8% | COMMON | Formula type (shared/array/etc) |
| `si` | 49.5% | COMMON | Shared formula index |
| `ref` | 1.6% | RARE | Shared formula range |
| `ca` | 0.09% | RARE | Calculate always |
| `xml:space` | 0% | RARE | Whitespace preservation |
| Others | 0% | RARE | dt2D, dtr, del1, del2, r1, r2, bx |

## Worksheet Children

| Element | Usage | Tier |
|---------|-------|------|
| `sheetData` | 100% | CORE |
| `pageMargins` | 100% | CORE |
| `dimension` | 100% | CORE |
| `sheetFormatPr` | 100% | CORE |
| `sheetViews` | 100% | CORE |
| `cols` | 92.2% | CORE |
| `conditionalFormatting` | 89.1% | CORE |
| `pageSetup` | 65.1% | CORE |
| `mergeCells` | 56.1% | CORE |
| `sheetPr` | 50.0% | COMMON |
| `hyperlinks` | 28.2% | COMMON |
| `headerFooter` | 25.3% | COMMON |
| `drawing` | 22.7% | COMMON |
| `sheetProtection` | 19.8% | COMMON |
| `phoneticPr` | 14.9% | COMMON |
| `dataValidations` | 10.3% | COMMON |
| `autoFilter` | 8.0% | RARE |
| `sortState` | 7.0% | RARE |
| `legacyDrawing` | 6.7% | RARE |
| `printOptions` | 6.7% | RARE |
| `extLst` | 5.5% | RARE |
| `rowBreaks` | 5.0% | RARE |
| Others | <5% | RARE |

## Implementation Strategy

### Option A: Compile-time features (Cargo features)
```toml
[features]
default = ["sml-common-attrs"]
sml-core-attrs = []  # Always included
sml-common-attrs = ["sml-core-attrs"]
sml-all-attrs = ["sml-common-attrs"]
```

### Option B: Runtime skip flags
Generate parsers that check a global/thread-local config:
```rust
if config.parse_rare_attrs {
    // parse thickTop, thickBot, etc.
}
```

### Option C: Lazy parsing
Parse core attrs eagerly, defer rare attrs to on-demand parsing.

## Recommendation

**Option A (Cargo features)** for attributes since:
- Zero runtime cost when disabled
- Most users won't need rare attributes
- Can still access via `sml-all-attrs` feature

Generate different parser code based on features:
- `sml-core-attrs`: Only parse r, s, spans for rows; r, s, t for cells
- `sml-common-attrs`: Add hidden, ht, customFormat, etc.
- `sml-all-attrs`: Parse everything per spec
