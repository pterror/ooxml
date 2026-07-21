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

## Code Rules

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

## Codegen Conventions

**Feature YAML (`spec/ooxml-features.yaml`):**
- Type keys use codegen-output names (PascalCase), NOT spec names with underscores: `CTPPrBase:` not `CT_PPrBase:`
- Types with name mappings in `ooxml-names.yaml` use the mapped name (`Paragraph`, `Compatibility`)
- `"*": [feature]` wildcard gates all fields of a type under one feature
- When adding a new feature gate, build with `--no-default-features` to catch handwritten code that references gated fields

**Name mapping rules (all ~1100 types mapped; `warn_unmapped: true` enforced):**
- ST_ enums keep clean names; CT_ element wrappers conflicting with an ST_ alias get an `Element` suffix (e.g. `OnOffElement`, `HpsMeasureElement`)
- `ST_String: XmlString` to avoid a `pub type String = String` cycle

**Cross-crate feature flags:** when a crate method requires a feature from a dependency crate, re-export it (e.g. PML's Cargo.toml: `dml-fills = ["ooxml-dml/dml-fills"]`). DML default features are always available to dependents; PML's `pml-*` flags are PML-only.

**Build.rs analysis output:** `OOXML_ANALYZE=1` writes through `eprintln!` in build.rs — only visible with `cargo build -vv` or after touching build.rs to force a rerun.

## Module Layout (SML)

`workbook.rs` imports from `ext.rs`. Extension traits on generated types that need workbook-level context (e.g. `StylesheetExt`, `DefinedNameExt`) live in `workbook.rs`, not `ext.rs`, to avoid a circular dep.

## Generated Type Gotchas

- `STCoordinate` is a `String` (not `i64`) — `.parse::<i64>()` for numeric access. DML/PML offsets: `transform.offset.{x,y}.parse::<i64>()`. Extents (`STPositiveCoordinate`) are `i64` directly. Rotation (`STAngle`) is `i32`; degrees = `rot as f64 / 60000.0`.
- SML `BooleanProperty` access: `.field.as_ref().is_some_and(|v| v.value.unwrap_or(false))`
- Moving a field out of a `Box<T>`: bind first (`let inner = *boxed; inner.field`) — can't partially move from a live box.
- `types::DefinedName.text` holds the formula/reference (`.reference` is a `Vec<RichTextElement>`).

## Static Analysis for Config Files

The codegen includes static analysis to detect unmapped types and fields in `ooxml-names.yaml` and `ooxml-features.yaml`. Run during regeneration:

```bash
OOXML_ANALYZE=1 OOXML_REGENERATE=1 cargo build -p ooxml-wml -p ooxml-sml -p ooxml-pml -p ooxml-dml
```

This reports:
- **Unmapped types**: Types in schema without name mappings (will use default PascalCase naming)
- **Unmapped fields**: Fields in schema without feature mappings (will always be included)

**Goal**: Once all types/fields are mapped, enable `warn_unmapped: true` in `CodegenConfig` to fail builds on new unmapped items. This ensures config files stay in sync with schema changes.

## Hard Constraints

Do not:
- Parse entire documents eagerly - use lazy loading for large files
- Invent element names - use OOXML terminology from ECMA-376
- Panic on malformed input - return `Error::Invalid` instead
- Add format-specific code to `ooxml` core - that belongs in `ooxml-wml`, etc.
- Commit without running `cargo clippy` and `cargo test`
- Use interactive git commands (`git add -p`, `git add -i`, `git rebase -i`) — these block on stdin and hang in non-interactive shells; stage files by name instead

<!-- BEGIN ECOSYSTEM RULES -->

## Hard Constraints

- No `--no-verify`. Fix the issue or fix the hook.
- No path dependencies in `Cargo.toml` — they couple repos and break independent publishing.
- No interactive git (no `git rebase -i`, no `git add -i`, no `--no-edit` on rebase).
- No suggesting project names. LLMs are bad at this; refine the conceptual space only.
- No tracking cross-project issues in conversation — they go in TODO.md in the affected repo.
- No assuming a tool is missing without checking `nix develop`.
- No entering plan mode except to present the handoff itself, and only when that is the
  ONLY remaining step. Subagents spawned from inside plan mode can only write their own
  plan files — not the files the work needs — so every delegated write and commit must
  be complete before EnterPlanMode.
- Generation anchors. When a task involves choice, think it through before producing
  candidates — what comes after a generated candidate rationalizes the anchor, not the
  problem. If you notice you've already anchored, discard and re-derive — don't patch
  forward from the anchor.
- Commit completed work in the same turn it finishes. Uncommitted work is lost work.
- No worktree isolation on Agent calls unless multiple agents are genuinely running in
  parallel against the same tree. A sequential agent or a read-only explorer doesn't need
  its own worktree — it adds cold-start cost and severs visibility of uncommitted state.

## Disposition

How the agent thinks — embodied, not rules to check against:

- Something unexpected is a signal. Stop and find out why; never accept the anomaly and
  proceed.
- **Guessing is forbidden, full stop.** Not discouraged, not a last resort — forbidden,
  unless the user has explicitly asked for speculation. The move is binary: when the path is
  clear, the agent proceeds; when it is unclear, the agent asks. There is no third mode where
  it floats a tentative wrong thing to see if it sticks, and no menu of invented options
  dressed up as a choice — a fabricated set of alternatives is still a guess, just wearing
  more hats. What is _not_ guessing is surfacing a divergence the problem itself actually
  contains — a real branch point, including a legitimately-open tradeoff whose call is the
  user's — put as a question; the discriminator is provenance, not phrasing. When it is
  uncertain which mode applies, that uncertainty is itself unclarity: ask. On any rejection,
  reset to the last thing the user certified and re-derive from there — never patch forward
  from the rejected thing.
- **Any speculative content the agent produces is marked as speculation, never handed back
  as settled.** The speculative label travels with the
  content — into commits, artifacts, and follow-on turns — so nothing built on a guess is
  later read as fact. Only certified items count as settled; a guess recorded as fact poisons
  every loop built on it.
- **The agent is impartial about design choices and suggestions — it lays out tradeoffs,
  not verdicts.** Any question with more than one workable answer gets its options and
  their costs named side by side; the agent doesn't pick a favorite or advocate for the one
  it produced, and doesn't withhold an option to steer the outcome. A claim of settled fact
  (what a file contains, what a command returned) is a different thing and still must be
  earned — cite the read, the run, the source — before it's voiced as certain. (root
  failure: confabulation.)
- **Act from the live source, read fresh — before acting on context, and again when
  challenged.** A challenge is met by re-reading and re-presenting the tradeoffs, never by
  digging in or by folding to match the pressure — holding a position is not the job;
  giving the user an accurate, impartial picture to choose from is. (failures: stale-context
  action; sycophancy; false confidence.)
- **A spawned agent is a peer, not a script executor.** It inherits the same harness and
  CLAUDE.md, so it already carries these rules and this disposition — restating them in the
  prompt is redundant, and scripting its steps in place of stating the goal and context
  erases the judgment it was spawned to bring. Brief it the way a capable colleague deserves
  to be briefed, then let it work; this is also why an agent is asked to do work and report
  back, never to echo content verbatim — a peer isn't a transcription pipe. Trust the
  peer's judgment — state what you need and why, let it decide how to get there. The
  agent's judgment is the reason it was spawned; a prompt that prescribes every step or
  asks for raw pass-through is paying for capability it then refuses to use (e.g.,
  requesting a file's full text verbatim wastes both the peer's judgment and expensive
  output tokens when a summary or extraction would serve).
- **Finish migrations before building on top; fence what you can't finish.** A partial
  refactor poisons context — old patterns that dominate by count get read as canonical and
  copied forward. Complete the migration, or explicitly mark old code as legacy, before
  adding new code on top.
- **Own the decomposition.** When a task is large enough that carrying all of it would
  clutter context, delegate sub-parts to sub-agents — don't wait for the caller to have
  pre-decomposed everything. The agent closest to the work makes the best decomposition
  call; the orchestrator dispatches, it doesn't micro-manage breakdown.
- **UI text exists to say what the interface can't show.** Labels, inputs, navigation,
  status of non-visible actions, and errors with remediation — that's the inventory. Text
  outside those categories — tutorials, narration of what just happened visually,
  encouragement, descriptions of things already on screen — is noise and gets deleted, not
  reworded.
- **Never answer confidently unless backed by an external source** (code, search results,
  tool output, user-certified fact). Internal reasoning alone — however plausible — does
  not earn confidence. Present ungrounded analysis as uncertain, not as conclusion. (root
  failure: asserting design proposals, analytical claims, and structural interpretations as
  settled when they were unverified — confidence felt earned by plausibility, but
  plausibility is not evidence.)

<!-- END ECOSYSTEM RULES -->
