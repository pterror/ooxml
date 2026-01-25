# Fuzz Testing for ooxml-sml

This directory contains fuzz targets for testing malformed XML input handling.

## Prerequisites

Fuzz testing requires **nightly Rust** for address sanitizer support.

### Using Nix (recommended)

The project provides a fuzz shell with nightly Rust via fenix:

```bash
# From project root
nix develop .#fuzz

# Then run fuzz targets
cargo fuzz build
cargo fuzz run fuzz_worksheet_xml
```

### Using rustup

```bash
rustup install nightly
rustup default nightly
# or use +nightly flag with each command
```

## Running Fuzz Tests

```bash
# Build all targets
cargo +nightly fuzz build

# Run worksheet XML parser fuzzing
cargo +nightly fuzz run fuzz_worksheet_xml

# Run serde deserialize fuzzing
cargo +nightly fuzz run fuzz_serde_deserialize

# Run for a limited time (e.g., 60 seconds)
cargo +nightly fuzz run fuzz_worksheet_xml -- -max_total_time=60
```

## Targets

- **fuzz_worksheet_xml**: Tests the `FromXml` event-based parser for `Worksheet`
- **fuzz_serde_deserialize**: Tests `quick-xml` serde deserialization for various types

## Expected Behavior

All fuzz targets should gracefully handle malformed input by returning errors,
never panicking or crashing. If a crash is found, it indicates a bug that needs fixing.

## Corpus

After running, discovered interesting inputs are saved in `corpus/<target>/`.
These can be used as regression tests.
