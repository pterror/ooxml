# TODO

Backlog for ooxml development. Items roughly ordered by priority within each section.

## Writing API Gaps

DocumentBuilder handles common cases but doesn't expose:

- [x] Headers/footers creation
- [x] Footnotes/endnotes creation
- [x] Comments creation
- [x] Anchored/floating images (use `run.add_anchored_image()` with `set_wrap_type()`)
- [ ] Styles creation/writing

## Reading Enhancements

- [x] More w:rFonts attributes (hAnsi, eastAsia, cs) - now parses all four font attributes
- [x] Core properties (docProps/core.xml) - title, author, etc.
- [x] App properties (docProps/app.xml) - word count, pages, etc.

## Advanced Features

- [ ] Revision tracking (w:ins, w:del for tracked changes)
- [ ] Math equations (integrate ooxml-omml crate)
- [ ] Table of contents (read)
- [ ] Bookmarks
- [ ] Text boxes
- [ ] SmartArt (limited)
- [ ] Charts (limited)

## Developer Experience

- [ ] Better error messages with context (line numbers, element paths)
- [ ] More examples in docs
- [x] Real-world usage examples in `examples/` directory (read_docx, create_docx, read_metadata)
- [ ] API documentation improvements

## Robustness

- [ ] Edge case handling from corpus analysis
- [ ] More comprehensive tests against real-world documents
- [ ] Fuzz testing for malformed input

## Other Formats

### SpreadsheetML (Excel) - ooxml-sml

- [ ] Workbook structure
- [ ] Worksheets
- [ ] Cells and values
- [ ] Formulas (as strings, not evaluated)
- [ ] Basic formatting
- [ ] Shared strings

### PresentationML (PowerPoint) - ooxml-pml

- [ ] Presentation structure
- [ ] Slides
- [ ] Shapes and text
- [ ] Images
- [ ] Basic transitions

### DrawingML - ooxml-dml

- [ ] Full DrawingML support (currently minimal)

## Infrastructure

- [ ] GitHub Pages documentation
- [ ] CI/CD for crates.io publishing
- [ ] Changelog generation
