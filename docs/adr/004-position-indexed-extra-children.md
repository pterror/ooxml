# ADR-004: Position-Indexed Extra Children for Roundtrip Fidelity

## Status
Accepted

## Context

OOXML documents contain elements we don't model ("unknown children"). For roundtrip fidelity — parse a document and write it back without data loss — these must be preserved. The question is how to preserve their *position* among siblings.

There are currently two approaches in the codebase:

**Handwritten parsers** (e.g., `DocumentSettings` in WML) use `Vec<PositionedNode>`, where each unknown child stores its original index among all siblings. This allows interleaving unknowns back into their original positions during serialization.

**Generated parsers** (codegen) use `Vec<RawXmlNode>` with no position tracking. Unknown children are appended after all known children during serialization. This is a fidelity regression: an element like `<a/><unknown/><b/>` roundtrips as `<a/><b/><unknown/>`.

This matters because the modify-and-save use case — open a document, change one field, save — should produce minimal structural diff. Users diffing before/after will see spurious reordering of elements they didn't touch.

### Alternatives considered

**Gap/slot model**: Define an ordered list of "known child slots" per type. Assign unknowns to the gap between adjacent slots. Insertion is ergonomic (just put things in the right slot), but deriving slot order from OOXML schemas with heavy `choice`/group reuse is complex, and the codegen cost is per-type rather than one-time.

**Simple index model** (chosen): Every child — known or unknown — gets a monotonically increasing position during parsing. Unknowns store this position. During serialization, interleave unknowns among known children by position. Schema-agnostic, one-time codegen cost, works uniformly across all four crates.

The gap/slot model and index model provide equivalent fidelity. The index model is simpler to implement in codegen and doesn't require per-type metadata.

### Future: position-tracking known children

A further improvement would be to also track the original positions of known children, allowing no-op writes to preserve the exact original child order (even when it deviates from schema order). This would eliminate diffs caused by schema-order canonicalization. That's a separate, larger change and is not part of this ADR.

## Decision

Switch `extra_children` from `Vec<RawXmlNode>` to `Vec<PositionedNode>` in all generated types, restoring parity with handwritten parsers.

### Parsing (parser_gen.rs)

Add a `child_idx: usize` counter that increments for every child element encountered (known or unknown). When capturing an unknown child, wrap it as `PositionedNode::new(child_idx, node)`.

```
// Pseudocode for generated parser
let mut child_idx = 0;
loop {
    match event {
        Start(e) if known => { parse_known_field(e); child_idx += 1; }
        Start(e) if unknown => {
            let node = RawXmlElement::from_reader(reader, &e);
            extra_children.push(PositionedNode::new(child_idx, RawXmlNode::Element(node)));
            child_idx += 1;
        }
        // Same for Empty events
    }
}
```

### Serialization (serializer_gen.rs)

Instead of appending `extra_children` at the end, interleave them among known children by position. Each known child field is emitted at a logical index (its position in the schema-order field list, accounting for Vec fields that expand to multiple children). Before emitting known child N, flush any `PositionedNode` entries whose position falls before N.

```
// Pseudocode for generated serializer
let mut extra_iter = self.extra_children.iter().peekable();
let mut emit_idx = 0;

fn flush_extras(extra_iter, emit_idx, writer) {
    while extra_iter.peek().map(|e| e.position <= emit_idx) {
        extra_iter.next().unwrap().node.write_to(writer);
    }
}

// For each known child field in schema order:
flush_extras(&mut extra_iter, emit_idx, writer);
write_known_child(field, writer);
emit_idx += count_of_emitted_children;

// Flush remaining extras at end
for extra in extra_iter { extra.node.write_to(writer); }
```

### Struct field (codegen.rs)

Change the generated field type:
```rust
// Before
pub extra_children: Vec<ooxml_xml::RawXmlNode>,

// After
pub extra_children: Vec<ooxml_xml::PositionedNode>,
```

## Consequences

### Positive
- Restores roundtrip ordering fidelity for all 400+ generated types across WML/SML/PML/DML
- One-time codegen change benefits every crate and future schemas
- Uses existing `PositionedNode` type (already defined in ooxml-xml)
- Schema-agnostic — works regardless of content model complexity
- Minimal API surface change (field type changes, but most code uses extension traits)

### Negative
- Slightly more memory per unknown child (one extra `usize`)
- Hand-written code that constructs `extra_children` directly needs updating (small — mostly test helpers)
- Interleaving logic in serializer is slightly more complex than "append at end"

### Not addressed
- Known children are still emitted in schema order, not original parse order. A document parsed with known children in non-schema order will be canonicalized. This is acceptable for now and can be addressed separately if needed.
- Insertion semantics for newly-added children in the modify-and-save path. New children currently get no position and appear in schema order. A `(position, bias)` scheme could make insertion more precise, but there's no demand for it yet.

## References

- `PositionedNode` definition: `crates/ooxml-xml/src/raw_xml.rs`
- Handwritten example: `DocumentSettings` parser in `crates/ooxml-wml/src/document.rs`
- ADR-003: Generated types as primary data model
- ECMA-376 Part 1, Section 8: Packaging conventions
