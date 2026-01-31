# CLAUDE.md

Instructions for Claude Code working on the ooxml project.

## Project Overview

High-quality Rust library for reading/writing Office Open XML formats (DOCX, XLSX, PPTX).

See `SPEC.md` for the full specification and roadmap.

## Architecture

```
crates/
  ooxml/          # Core: OPC packaging, relationships, content types
  ooxml-wml/      # WordprocessingML (Word documents)
```

## Key Design Decisions

1. **Typed over stringly** - Every XML element should have a Rust struct
2. **Roundtrip preservation** - Unknown/unsupported elements must be preserved
3. **Lazy when possible** - Don't parse everything upfront
4. **Spec-driven** - ECMA-376 is the source of truth

## Development

```bash
cargo test          # Run tests
cargo clippy        # Lint
cargo doc --open    # View docs
```

## Current Priority

v0.1 goal: Basic Word document support
1. OPC packaging (ZIP read/write, relationships)
2. Document structure (body, paragraphs, runs)
3. Basic formatting (bold, italic, underline)
4. Styles (read style definitions, apply to content)
5. Tables (basic)
6. Images (embedded)

## Workflow

**Batch cargo commands** to minimize round-trips:
```bash
cargo clippy --all-targets --all-features -- -D warnings && cargo test
```
After editing multiple files, run the full check once — not after each edit. Formatting is handled automatically by the pre-commit hook (`cargo fmt`).

**When making the same change across multiple crates**, edit all files first, then build once.

**Minimize file churn.** When editing a file, read it once, plan all changes, and apply them in one pass. Avoid read-edit-build-fail-read-fix cycles by thinking through the complete change before starting.

**Use `normalize view` for structural exploration:**
```bash
~/git/rhizone/normalize/target/debug/normalize view <file>    # outline with line numbers
~/git/rhizone/normalize/target/debug/normalize view <dir>     # directory structure
```

## Commit Convention

Use conventional commits: `type(scope): message`

Types:
- `feat` - New feature
- `fix` - Bug fix
- `refactor` - Code change that neither fixes a bug nor adds a feature
- `docs` - Documentation only
- `chore` - Maintenance (deps, CI, etc.)
- `test` - Adding or updating tests

Scope is optional but recommended (e.g., `wml`, `corpus`, `packaging`).

## Code Conventions

**Error handling:**
- Use `thiserror` for error enums
- Each crate has its own `Error` and `Result` type
- Wrap external errors with `#[from]` where appropriate
- Use `Invalid(String)` for malformed input, `Unsupported(String)` for unimplemented features

**Naming:**
- Struct names match OOXML element names where practical (e.g., `Run` for `<w:r>`)
- Use full words, not abbreviations (except where OOXML itself abbreviates)
- Prefix internal modules with the crate's domain (`wml_`, `sml_`, etc.) only if disambiguation needed

**XML handling:**
- Use `quick-xml` for parsing and serialization
- Preserve unknown elements/attributes in a catch-all field for roundtrip fidelity
- Namespaces: define constants for common OOXML namespaces (see ECMA-376 Part 1, §8)

**Serialization patterns:**
- Enums that map to XML values should have a `to_xml_value()` method
- Avoid inline match expressions when converting enums to XML strings
- Keep serialization logic close to the type definition, not scattered in writer code

**Dependencies:**
- Workspace dependencies in root `Cargo.toml`
- Internal crates use `{ workspace = true }`

## Testing Strategy

- Unit tests for individual elements
- Roundtrip tests: open → save → compare
- Fixture tests with real .docx files in `tests/fixtures/`
- Use `insta` for snapshot testing XML output

## References

- [ECMA-376 Standard](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/) - Source of truth
- [Open XML SDK Docs](https://docs.microsoft.com/en-us/office/open-xml/open-xml-sdk)
- [python-docx source](https://github.com/python-openxml/python-docx) - Good API reference

## Consumer

This library will be used by `rescribe` (document conversion library) for DOCX support.
The rescribe team is waiting on this to implement `rescribe-read-docx` and `rescribe-write-docx`.

## Core Rules

- **Reference the spec** - When implementing OOXML elements, cite the relevant ECMA-376 section
- **Test as you go** - Every new struct/parser needs a unit test
- **Preserve unknown data** - Never silently drop XML elements or attributes we don't understand
- **Verify roundtrips** - Changes to serialization must pass roundtrip tests

## Negative Constraints

Do not:
- Parse entire documents eagerly - use lazy loading for large files
- Invent element names - use OOXML terminology from ECMA-376
- Panic on malformed input - return `Error::Invalid` instead
- Add format-specific code to `ooxml` core - that belongs in `ooxml-wml`, etc.
- Commit without running `cargo clippy` and `cargo test`
- Use path dependencies in Cargo.toml - causes clippy to stash changes across repos
- Use `--no-verify` - fix the issue or fix the hook
- Assume tools are missing - check if `nix develop` is available for the right environment
