# ADR 002: Event-Based XML Parsing for Generated Types

## Status

Accepted

## Context

We generate Rust types from ECMA-376 RELAX NG schemas. These types need to be parseable from XML. Two approaches were considered:

1. **Serde-based parsing** - Use `quick-xml`'s serde support with `#[derive(Deserialize)]`
2. **Event-based parsing** - Generate `quick-xml` event loop code alongside types

## Benchmarks

We benchmarked both approaches parsing `<sheetData>` elements with varying sizes:

| Dataset | Serde | Events | Serde Overhead |
|---------|-------|--------|----------------|
| 50 cells (10x5) | 45.5 µs | 15.9 µs | **2.9x slower** |
| 1,000 cells (100x10) | 827 µs | 290 µs | **2.9x slower** |
| 10,000 cells (1000x10) | 8.85 ms | 3.0 ms | **3.0x slower** |
| 50,000 cells (1000x50) | 45.1 ms | 14.0 ms | **3.2x slower** |

Throughput comparison:
- Serde: ~1.1 Melem/s (million elements per second)
- Events: ~3.4 Melem/s

The overhead is consistent across all sizes (~3x), indicating this is fundamental to serde's deserialization approach (reflection-like dispatch, temporary allocations, trait object overhead).

## Decision

We chose **option 2: event-based parsing**.

Rationale:

- 3x performance difference is significant for document conversion (large spreadsheets can have millions of cells)
- We control the codegen, so complexity is amortized across all generated types
- Serde remains available for convenience (debugging, tests, small documents)
- Event-based code is more explicit and debuggable when issues arise

## Implementation

The codegen will generate:

1. A `FromXml` trait for parsing from `quick_xml::Reader`
2. Implementations for each generated struct and enum
3. Attribute parsing via iteration over `BytesStart::attributes()`
4. Child element dispatching via tag name matching
5. A `skip_element()` helper for unknown elements (roundtrip preservation)

Types will have both:
- `#[derive(Deserialize)]` for serde-based parsing (convenience)
- `impl FromXml` for event-based parsing (performance)

## Consequences

**Positive:**
- 3x faster parsing for performance-critical paths
- Generated code is readable and debuggable
- Both parsing modes available depending on use case

**Negative:**
- Larger generated code (both serde attrs and FromXml impls)
- More complex codegen logic
- Two code paths to maintain (though both are generated)

**Trade-offs:**
- Could skip serde entirely, but convenience value is high for tests/debugging
- Could generate only for "hot" types, but consistency is valuable

## Benchmark Reproduction

```bash
cargo bench -p ooxml-sml --bench parse_benchmark
```

Results captured 2026-01-23 on Linux with Rust stable.

## References

- [quick-xml Performance](https://github.com/tafia/quick-xml) - Underlying XML library
- Benchmark code: `crates/ooxml-sml/benches/parse_benchmark.rs`
