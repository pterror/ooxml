# ADR-003: Generated Types as Primary Data Model

## Status
Accepted

## Context

We have two options for representing OOXML structures:
1. Hand-written types with resolved values (current workbook.rs ~3800 lines)
2. Generated types from ECMA-376 schemas with extension traits

The generated approach provides:
- Spec compliance (types derived from official schemas)
- Reduced maintenance (codegen handles schema changes)
- Better performance (1.5-2x faster parsing than serde, per ADR-002)

## Decision

Use generated types as the primary data model with a layered architecture:

```
┌─────────────────────────────────────────────────────────────┐
│ Layer 3: High-level API (Workbook, Sheet)                   │
│   - Owns ResolveContext                                     │
│   - Provides ergonomic access patterns                      │
├─────────────────────────────────────────────────────────────┤
│ Layer 2: ResolveContext                                     │
│   - SharedStrings, Stylesheet, Themes                       │
│   - Passed explicitly to resolution methods                 │
├─────────────────────────────────────────────────────────────┤
│ Layer 1: Extension Traits (CellExt, RowExt, etc.)           │
│   - Add convenience methods to generated types              │
│   - Pure methods (no context) vs resolved methods (context) │
├─────────────────────────────────────────────────────────────┤
│ Layer 0: Generated Types (types::*)                         │
│   - Spec-driven structure from ECMA-376 schemas             │
│   - Raw XML values, not resolved                            │
│   - Parsed via generated FromXml trait                      │
└─────────────────────────────────────────────────────────────┘
```

### Layer 0: Generated Types

Generated directly from ECMA-376 RELAX NG schemas. Store raw XML values.

```rust
// From codegen - DO NOT MODIFY
pub struct Cell {
    pub reference: Option<CellRef>,
    pub cell_type: Option<CellType>,
    pub value: Option<XmlString>,  // Raw: "0" (shared string index)
    pub formula: Option<Box<CellFormula>>,
    pub style_index: Option<u32>,
    // ... other spec fields
}
```

### Layer 1: Extension Traits

Add behavior without modifying generated code. Two categories:

**Pure methods** (no context needed):
```rust
pub trait CellExt {
    /// Parse column number from reference (e.g., "B5" -> 2)
    fn column_number(&self) -> Option<u32>;

    /// Parse row number from reference (e.g., "B5" -> 5)
    fn row_number(&self) -> Option<u32>;

    /// Check if cell has a formula
    fn has_formula(&self) -> bool;
}
```

**Resolved methods** (context required):
```rust
pub trait CellResolveExt {
    /// Get resolved cell value
    fn resolved_value(&self, ctx: &ResolveContext) -> CellValue;

    /// Get value as display string
    fn value_as_string(&self, ctx: &ResolveContext) -> String;

    /// Get value as number (if applicable)
    fn value_as_number(&self, ctx: &ResolveContext) -> Option<f64>;
}
```

### Layer 2: Resolution Context

Holds shared state needed for value resolution:

```rust
pub struct ResolveContext {
    pub shared_strings: Vec<String>,
    pub stylesheet: Stylesheet,
    // Future: themes, external links, etc.
}

impl ResolveContext {
    /// Resolve a shared string index to the actual string
    pub fn shared_string(&self, index: usize) -> Option<&str>;

    /// Get number format for a style index
    pub fn number_format(&self, style_index: u32) -> Option<&str>;
}
```

### Layer 3: High-level API

Ergonomic wrappers that own context and provide convenient access:

```rust
pub struct Workbook<R: Read + Seek> {
    archive: ZipArchive<R>,
    context: ResolveContext,
    sheet_names: Vec<String>,
}

pub struct Sheet<'a> {
    name: String,
    data: types::SheetData,      // Generated type
    context: &'a ResolveContext, // Borrowed from Workbook
}

impl Sheet<'_> {
    /// Iterate rows with automatic value resolution
    pub fn rows(&self) -> impl Iterator<Item = &types::Row>;

    /// Get cell with convenient value access
    pub fn cell(&self, reference: &str) -> Option<CellView<'_>>;
}

/// Convenience wrapper for cell access with bound context
pub struct CellView<'a> {
    cell: &'a types::Cell,
    context: &'a ResolveContext,
}

impl CellView<'_> {
    pub fn value_as_string(&self) -> String {
        self.cell.value_as_string(self.context)
    }
}
```

## Memory Efficiency

### Shared Strings (Major Win)

Excel stores repeated strings once in a shared string table. Cells reference by index.

**Hand-written (eager resolution):**
```
SharedStrings: ["Hello", "World", ...]  // Loaded then discarded
Cell 1: value = "Hello"  // Copied
Cell 2: value = "Hello"  // Copied again
Cell 3: value = "Hello"  // Copied again
// 1000 cells with same string = 1000 copies
```

**Generated (lazy resolution):**
```
SharedStrings: ["Hello", "World", ...]  // Kept in ResolveContext
Cell 1: value = "0"  // Index only
Cell 2: value = "0"  // Index only
Cell 3: value = "0"  // Index only
// 1000 cells = 1 string + 1000 tiny indices
```

### Cell Structure Size

Generated cells have more optional fields but use `Option<Box<T>>` for large nested types:
- `Option<Box<T>>` = 8 bytes (null pointer optimization)
- Most fields are `None` in typical spreadsheets
- Raw value strings are tiny (indices like "0", "42")

## Performance

### Parsing (Already Benchmarked)

| Parser | Cell (simple) | Speedup vs serde |
|--------|---------------|------------------|
| serde | 648 ns | baseline |
| FromXml (generated) | 358 ns | 1.8x faster |

### Value Resolution

| Access Pattern | Cost |
|----------------|------|
| Single cell lookup | O(1) shared string lookup |
| Iterate all cells | O(n) with n lookups |
| Bulk export | Consider caching resolved values |

For typical use cases (read specific cells, display ranges), lazy resolution is efficient.
For bulk operations, the high-level API can provide batch resolution.

## Migration Path

1. **Phase 1**: Add extension traits module (`src/cell_ext.rs`, `src/row_ext.rs`)
2. **Phase 2**: Implement ResolveContext
3. **Phase 3**: Update Sheet to use `types::SheetData` internally
4. **Phase 4**: Update public API to expose generated types
5. **Phase 5**: Remove hand-written Row/Cell types
6. **Phase 6**: Deprecate and remove old parsing code

## Consequences

### Positive
- Single source of truth for types (generated from spec)
- Better memory efficiency for shared strings
- Faster parsing (1.5-2x vs serde)
- Explicit context passing (no hidden state)
- Easier to maintain (codegen handles schema updates)

### Negative
- API change for consumers (need to pass context for resolution)
- Slightly more verbose for simple cases
- Two-step access pattern (get cell, then resolve value)

### Mitigated by Layer 3
The high-level API (CellView, Sheet with bound context) provides ergonomic access
that hides the context passing for common use cases.

## References

- ADR-002: Event-based parsing (no serde)
- ECMA-376 Part 1: Shared String Table (§18.4)
- Benchmarks: `crates/ooxml-sml/benches/parse_benchmark.rs`
