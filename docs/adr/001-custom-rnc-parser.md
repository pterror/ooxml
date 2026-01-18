# ADR 001: Custom RNC Parser for Code Generation

## Status

Accepted

## Context

We need to generate Rust types from the ECMA-376 RELAX NG Compact (.rnc) schema files. This enables type-safe, spec-driven code generation rather than hand-writing 400+ structs and enums.

Options considered:

1. **relaxng-rust crate** - Existing Rust RELAX NG parser
2. **Convert RNC to XSD** - Use external tools, then parse XSD
3. **Custom RNC parser** - Write a minimal parser for OOXML schemas

## Decision

We chose **option 3: custom RNC parser**.

Rationale:

- **relaxng-rust** was last updated February 2021, fails 125 of 384 tests, and is not actively maintained
- The OOXML .rnc files use a consistent subset of RNC syntax (no includes, no grammar refs, predictable patterns)
- A focused parser (~400 lines) is simpler than depending on a general-purpose but incomplete library
- We only need to extract structure for codegen, not validate documents
- Full control over error messages and AST shape

## Consequences

**Positive:**
- No external dependency for parsing
- Parser is tailored to OOXML patterns (handles `xsd:string`, `element default`, etc.)
- Easy to extend if we encounter new patterns

**Negative:**
- Maintenance burden if RNC syntax varies significantly in other ECMA-376 parts
- Won't handle arbitrary RNC files (by design)

**Risks:**
- SpreadsheetML or PresentationML schemas might use patterns we haven't seen yet (mitigated: we can extend the parser when needed)

## References

- [relaxng-rust](https://github.com/dholroyd/relaxng-rust) - Evaluated alternative
- [RELAX NG Compact Tutorial](https://relaxng.org/compact-tutorial-20030326.html) - Syntax reference
- ECMA-376 5th Edition spec files in `/spec/`
