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

**Always commit when work is complete.** After finishing a task and verifying it passes `cargo clippy` and `cargo test`, create a commit before moving on. Don't leave working changes uncommitted.

## Session Handoff

Use plan mode as a handoff mechanism when:
- A task is fully complete (committed, pushed, docs updated)
- The session has drifted from its original purpose
- Context has accumulated enough that a fresh start would help

Before entering plan mode:
- Update TODO.md with any remaining work
- Update memory files with anything worth preserving across sessions

Then enter plan mode and write a plan file that either:
- Proposes the next task if it's clear: "next up: X — see TODO.md"
- Flags that direction is needed: "task complete / session drifted — see TODO.md"

ExitPlanMode hands control back to the user to approve, redirect, or stop.

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

**Goal: every feature exercised by a fixture, every fixture validated by a roundtrip test.**

Layers (in order of scope):
1. **XML unit tests** — individual element parse/serialize, no full document. SML has these in `crates/ooxml-sml/tests/fixtures/xml/`. WML and PML need the same.
2. **Fixture roundtrip tests** — `ooxml-fixtures` generates 178 CC0 `.docx`/`.xlsx`/`.pptx` files with typed JSON manifests. Each reader crate should have a test that reads every fixture and asserts all manifest assertions pass.
3. **Structural edge cases** — nested tables, mixed content, unusual-but-valid constructions. These live in the fixture crate alongside the standard fixtures.
4. **Malformed/adversarial tests** — truncated ZIP, missing parts, broken XML, unknown namespaces. These live in each crate's `tests/` as `Result::is_err()` assertions. Do NOT put these in the fixture crate (they're not CC0-distributable examples).
5. **Corpus tests** — parse large real-world corpora and assert >95% success rate. Marked `#[ignore]` since the corpus is not vendored.

**NapierOne corpus:** Available locally at `corpora/napierone/` but cannot be vendored due to licensing. Corpus tests are `#[ignore]`'d and require the corpus to be present. SML and PML have corpus tests; WML needs one.

**Fixture coverage checklist** — these categories must each have at least one fixture:
- WML: text formatting, paragraphs, lists, tables, headers/footers, images, hyperlinks, comments, footnotes, endnotes, track changes, field codes (TOC), forms, math (all OMML operators), bookmarks, text boxes, nested tables
- SML: cell values/types, formulas, number formats, font/fill/border styling, layout (freeze/autofilter), merges, hyperlinks, data validation, conditional formatting, multiple sheets, comments, protection
- PML: shapes (all types), connectors, groups, text formatting, tables, images, hyperlinks, animations, transitions, notes, master/layout, charts

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

## Generated Files — IMPORTANT

Generated files (`generated.rs`, `generated_parsers.rs`, `generated_serializers.rs`) are **committed to the repo** so that users can build without the ECMA-376 RNC schemas in `/spec/`.

**When you change `ooxml-codegen`, you MUST regenerate ALL crates and commit the results.** The codegen is shared — a change to `codegen.rs`, `parser_gen.rs`, or `serializer_gen.rs` affects every crate's generated output. Stale generated files will silently diverge from the codegen and cause confusing bugs later.

```bash
# Regenerate everything (requires schemas in /spec/)
OOXML_REGENERATE=1 OOXML_GENERATE_PARSERS=1 OOXML_GENERATE_SERIALIZERS=1 \
  cargo build -p ooxml-wml -p ooxml-sml -p ooxml-pml -p ooxml-dml
```

After regenerating, commit all changed `generated*.rs` files in the same commit (or immediately after) the codegen change.

## Static Analysis for Config Files

The codegen includes static analysis to detect unmapped types and fields in `ooxml-names.yaml` and `ooxml-features.yaml`. Run during regeneration:

```bash
OOXML_ANALYZE=1 OOXML_REGENERATE=1 cargo build -p ooxml-wml -p ooxml-sml -p ooxml-pml -p ooxml-dml
```

This reports:
- **Unmapped types**: Types in schema without name mappings (will use default PascalCase naming)
- **Unmapped fields**: Fields in schema without feature mappings (will always be included)

**Goal**: Once all types/fields are mapped, enable `warn_unmapped: true` in `CodegenConfig` to fail builds on new unmapped items. This ensures config files stay in sync with schema changes.

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
