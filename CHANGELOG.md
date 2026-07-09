# Changelog

All notable changes to `simple-super-doc` are documented here. The format is based
on [Keep a Changelog](https://keepachangelog.com/), and this project adheres to
[Semantic Versioning](https://semver.org/).

## [0.11.6] - 2026-07-09

Map inline run/paragraph elements that were silently dropped, found by analysis
and real docx (pandoc `unicode.docx`, `inline_formatting.docx`).

### Fixed
- **`w:sym`**: symbol characters (Insert Symbol, Wingdings/Symbol fonts) emit the
  codepoint carrying the symbol font (guarded against an out-of-range `w:char`).
- **`w:cr`**: a carriage return becomes a soft line break.
- **`w:noBreakHyphen` / `w:softHyphen`**: emit U+2011 / U+00AD instead of nothing.
- **`w:vanish`**: hidden text is no longer rendered in the final view.
- **Paragraph shading (`pPr > w:shd`)**: a paragraph-level fill now backs the whole
  block (run-level `w:shd` was already handled).
- **`w:caps` / `w:smallCaps`**: rendered via `text-transform` / `font-variant`.
- **`w:dstrike`**: double strikethrough.

## [0.11.5] - 2026-07-08

Header/footer and letterhead fidelity, found by visual-diffing a real Clicksign
contract (Google Docs export) against Word.

### Fixed
- **Footnotes on a split table**: a footnote referenced by a row that stayed on an
  earlier page was dropped entirely and its reserved height was ignored, so the
  table broke at the wrong point. The page-aware paginator now reserves and records
  each split piece's footnotes on the page that holds it.
- **Bullet glyph**: a bullet list honored only CSS `disc` (`•`), ignoring the
  numbering's `w:lvlText`. A hyphen bullet (`lvlText="-"`) now renders `-`; symbol
  font glyphs (Wingdings, U+2022) still fall back to `disc`.
- **Header/footer images**: an image referenced inside a header/footer resolves
  against that part's own relationships file (`word/_rels/headerN.xml.rels`), not
  the document's, so letterhead logos are no longer dropped when the rIds collide.
- **Runs with multiple `<w:t>`**: a run packing `text <w:tab/> text` into one `<w:r>`
  (Google Docs export) no longer loses its text. `parseRun` splits the segments so
  the tabs stay attached to the trailing run and its position is preserved.
- **Header/footer tab alignment**: a tabbed paragraph with no explicit tab stops now
  gets Word's implicit center + right stops, so `left <tab> right` footers split to
  the margins instead of drifting to center.
- **Anchored header/footer logos**: a `wp:anchor` logo is positioned from its
  `positionH`/`positionV` offset (letterhead), instead of flowing inline at the left.
  Body-level floating layout stays out of scope.

### Added
- Browser (Playwright) regression tests for the page-aware renderer — table
  splitting, per-page footnotes, header/footer positioning — run with
  `npm run test:browser` (headless Chromium via `playwright-core`, wired into CI).
- Public "Try it live" GitHub Pages demo link in the README.

## [0.11.4] - 2026-07-01

### Changed
- Release automation: `npm version` auto-pushes tags so CI publishes on `v*` tags.

## [0.11.3] - 2026-07-01

### Security
- Sanitize image sources (defense-in-depth): only `data:image/*`, `http(s)` and
  relative image URLs are allowed, matching the existing hyperlink-scheme guard.

## [0.11.2]

### Fixed
- Respect table width, alignment (`w:jc`) and row heights (`w:trHeight`).

## [0.11.1]

### Fixed
- Drop the phantom frame-only page rendered between a cover and the content.

## [0.11.0]

### Fixed
- Per-page headers/footers: resolve the `first` and `even` variants (via `w:titlePg`
  and `w:evenAndOddHeaders`), not just `type="default"`.

## [0.10.0]

### Fixed
- Right-to-left: `w:bidi` paragraphs and `w:rtl` runs render with `dir="rtl"`.

## [0.9.0]

### Fixed
- Legacy VML images (`w:pict`/`v:imagedata`) render.
- External/linked images (`r:link`, `TargetMode="External"`) render from their URL.

## [0.8.0]

### Fixed
- OMML equations (`m:oMath`) recovered as linear text (inline, display and in cells).

## [0.7.0]

### Fixed
- Tracked moves (`w:moveTo`/`w:moveFrom`) surfaced like insertions/deletions.
- Smart tags (`w:smartTag`) and custom XML (`w:customXml`) unwrapped transparently.

## [0.6.0]

### Fixed
- Content controls (`w:sdt`): their block/inline/nested content renders in place
  instead of being dropped.

## [0.5.1]

### Fixed
- `PAGE`/`NUMPAGES` fields packed into a single `<w:r>` (Google Docs export) resolve.

## [0.5.0]

### Added
- Feature-complete typed IR + HTML renderer: text/styles, lists, tables
  (`gridSpan`/`vMerge`, borders), images, headers/footers, page-aware background
  pagination.

### Fixed
- Text boxes (`w:txbxContent`) recovered into the flow instead of being dropped.
- Table cell borders resolved by cascade (table style → `tblBorders` → `tcBorders`).
- A full-page background (cover) page is kept even when its only flow text is empty.

[0.11.5]: https://github.com/Rafael-T-Santos/simple-super-doc/releases/tag/v0.11.5
[0.11.4]: https://github.com/Rafael-T-Santos/simple-super-doc/releases/tag/v0.11.4
