# Changelog

All notable changes to `simple-super-doc` are documented here. The format is based
on [Keep a Changelog](https://keepachangelog.com/), and this project adheres to
[Semantic Versioning](https://semver.org/).

## [0.15.0] - 2026-08-06

Single line spacing is now the font's own line height, which is what Word means
by it. This moves where pages break in almost any document.

### Fixed
- **Single spacing was rendering ~25% too tight.** OOXML writes it as
  `w:line="240" w:lineRule="auto"`, and 240 counts 240ths of a *line* — a line
  being the font's own box (ascent + descent + gap), not the font size. The
  parser divided 240 by 240 and emitted the CSS number `1`, pinning every line
  to exactly the font size. Against a LibreOffice render of the same seven-page
  contract this renderer produced six pages; per-page text agreement on the
  worst page was 0.195. It now maps to `line-height: normal` and both contracts
  land on seven pages, with the worst page at 0.929 and page counts matching the
  reference on all five documents tested.

  There was a tell: with no `w:spacing` at all the renderer fell back to 1.15,
  so a document that *declared* single spacing came out tighter than one that
  said nothing.

### Added
- `ComputedStyle.lineHeightSingle` marks single spacing explicitly. It has to be
  explicit rather than merely absent: styles merge with `Object.assign`, where a
  missing key cannot override an inherited one, so a paragraph asking for single
  over a document default of 1.15 lines silently kept 1.15.

Non-single multipliers (1.5 lines, double) are unchanged: they still scale the
font size rather than the natural line. By the spec they should scale the line,
but doing that measured worse — a document default of 1.15 lines became 1.38 and
pushed the same contract to nine pages — so it is left alone pending better
evidence.

## [0.14.0] - 2026-08-06

### Added
- **Embedded fonts are loaded and used.** A `.docx` can carry its own font files
  (`w:embedRegular`, `w:embedBold`, `w:embedItalic`, `w:embedBoldItalic` in
  `fontTable.xml`). Until now they were ignored, so a document whose typeface was
  not installed rendered in whatever the browser fell back to. They are now read,
  de-obfuscated when Word applied its `.odttf` scrambling, and registered with
  the browser during `parse()` — before the renderer measures anything, since a
  font arriving after layout would reflow text under page breaks already decided.
  `DocxDocument.fonts` exposes them as data URLs so a consumer rendering
  elsewhere can emit its own `@font-face` rules. A font part that is corrupt or
  in a format the browser cannot load is skipped rather than failing the parse.

### Correction to 0.13.3
The 0.13.3 notes claimed the page-count difference against a LibreOffice render
was caused by these missing fonts, citing "345 lines against the reference's
445". **That was wrong on both counts.** The line-count figure came from a broken
measurement: `Range.getClientRects()` returns one rect per *block* when the range
spans block-level children, so paragraphs inside table cells counted as a single
line. And loading the real fonts turns out to make text *narrower*, not wider —
the same string measures 574px in DM Sans against 593px in the fallback — so the
missing fonts could not have been making our lines hold more text.

Embedded font loading is a real fidelity fix and is worth having on its own, but
it does not change the page count on the documents this was tested against. Why
this renderer lays a long contract out in 6 pages where LibreOffice uses 7 is
still open.

## [0.13.3] - 2026-08-06

A paragraph or list item too tall for the space left now continues on the next
page instead of jumping to it whole.

### Added
- **Line-level page breaking.** Until now the paginator could cut a table between
  its rows and a list between its items, but any single block that did not fit
  moved whole. One 100-line list item in a real contract therefore left 60% of a
  page blank. Paragraphs and list items are now cut between their line boxes.
  The cut is made with `Range.extractContents`, so a break landing inside a
  `<span>`, a link or a tracked change keeps its formatting on both pieces; a
  word is never split; and Word's widow/orphan rule is honored, so a break that
  would strand a single line is skipped and the block moves whole after all. A
  continued list item shows no second marker and does not disturb the numbering.
  A table CELL is still not cut: a tall single row moves whole.

### Fixed
- **Pages no longer absorb margins that collapse out of a block.** Whether a
  block fit was decided from its border box, which excludes a last child's bottom
  margin. The margin still takes space on the page, so the page silently
  stretched. The fit test now measures where the next block would actually start.

**On page counts:** this changes where breaks fall, so a document may render in
fewer pages than before. Comparing against a LibreOffice render of the same
contract, our page count moved from 7 to 6 while the reference stayed at 7. The
cause is not yet established — see the correction in 0.14.0, which retracts the
explanation originally given here.

## [0.13.2] - 2026-08-06

### Fixed
- **A template page no longer grows past its own height.** The paginator for
  documents with a full-page background measured each block on its own and read
  the wrapper's `offsetHeight`, which drops the block's margins: they collapse
  straight through a bare wrapper. The per-page total came out short, so content
  ran past the page box — and since the box is sized with `min-height`, the sheet
  silently stretched instead of breaking. Blocks are now measured together and
  each one's footprint is the distance to the next, which is the spacing the
  layout engine actually produced. Same rule the plain paginator already used.
  Page counts and page contents are unchanged on the documents this was found on;
  only the overflow goes away.

## [0.13.1] - 2026-08-06

A long list no longer jumps whole to the next page, leaving most of a page blank.

### Fixed
- **Lists split across pages.** The paginator could only cut a table between its
  rows; every other block moved whole to the next page. A list taller than the
  space left therefore jumped in one piece, and on a real contract that left 60%
  of a page empty while the reference render filled it. Lists now continue on the
  next page like tables do, and an ordered list's continuation resumes its
  numbering instead of restarting at 1. Measured against a LibreOffice render of
  the same contract, per-page text agreement went from 0.53 to 0.89 on the worst
  page, and four pages that wasted 33-60% of their height now waste 8-12%.
- **Pages no longer stretch past their own height.** The paginator counted a
  split block's consumed space with `offsetHeight`, which omits margins, while
  measuring every other block from its laid-out position. The two disagreed by up
  to 48px on a page, quietly pushing content past the page box.

Known limit: a single list item taller than the space left still moves whole,
because splitting *inside* an item means line-level breaking, which the renderer
does not do for any block.

## [0.13.0] - 2026-08-06

Table cells now honor the two properties that decide where their text sits and
which way it runs.

### Added
- **Cell vertical alignment (`w:vAlign`).** A cell set to center or bottom in Word
  no longer renders its text pinned to the top of the row. `TableCell.verticalAlign`
  exposes it on the IR. OOXML's `both` (vertical justification, which CSS has no
  equivalent for) maps to center, matching how Word draws a single line of it. Top
  is Word's default and stays implicit.
- **Rotated cells (`w:textDirection`).** The narrow vertical label column Word
  produces with "Text Direction" now renders rotated instead of horizontal, in both
  directions: `tbRl` (reads top-to-bottom) and `btLr` (reads bottom-to-top).
  `TableCell.textDirection` exposes it on the IR. The rotation is applied so the
  table sizing algorithm measures the rotated text, which means the row grows to
  hold a long label instead of clipping it. The horizontal `lrTb` default and the
  vertical-CJK `*V` variants are left horizontal on purpose.

Known limit: only a cell's own `w:tcPr` is read. A table *style* that sets
`vAlign` is not carried through, since the style cascade for cells is border-only.

## [0.12.0] - 2026-08-05

Heading rows now repeat when a table breaks across pages, and three more places
where a run's own text was being thrown away are fixed.

### Added
- **Repeating table heading rows (`w:tblHeader`).** A row marked as a heading in
  Word now repeats at the top of every page the table continues onto, instead of
  appearing only on the first. `TableRow.isHeader` exposes the flag on the IR.

### Fixed
- **Text next to a text box is no longer lost.** A text box nests a whole
  paragraph inside the run that carries it, and the paragraph's raw XML was being
  cut short at the text box's own paragraph end. Anything after that point in the
  enclosing paragraph lost its document order, and text sharing a run with the
  text box was dropped entirely.
- **Text around a footnote reference survives.** `<w:t>before</w:t>
  <w:footnoteReference/><w:t>after</w:t>` in one run rendered as just the
  reference marker; both pieces of text were dropped. Several references in one
  run are each placed correctly and numbered in document order.
- **A footnote's own text is no longer lost.** When Word packs the note's
  auto-number marker and the note text into the same run, the marker stood for the
  whole run and the note came out empty.

### Changed
- The demo's version badge is taken from `package.json` when the site is built,
  so it tracks releases instead of drifting (it had read v0.11.0 since then).

## [0.11.9] - 2026-08-05

Two content and layout fixes found in a real contract: a table changed shape at a
page break, and text sharing a run with an image disappeared.

### Fixed
- **A table that breaks across pages keeps its column widths.** The continued
  piece on the next page was rebuilt without the document's column definitions,
  so its columns came out all the same width while the first page showed them
  correctly. Measured on a real document: a 124px/556px pair rendered as
  339px/339px after the break. Tables spanning three or more pages are covered
  too, since each continuation is split again from the previous one.
- **Text is no longer lost when an image sits in the same run.** Word and Google
  Docs both write `<w:t>…</w:t><w:drawing/><w:t>…</w:t>` inside a single run. The
  image was treated as the whole run, so the text on both sides of it vanished.
  The image now lands between the texts, in document order, and several images in
  one run are each placed correctly.

## [0.11.8] - 2026-08-05

Documentation only — no code change. The README describing the 0.11.7 ordering
fixes landed after the 0.11.7 tag, so the package published to npm still carried
the previous text. This release ships it.

### Changed
- README: the ordering section now mentions that order is recovered in headers,
  footers and notes, and *within* a run; the breaks entry covers soft `<w:br/>` /
  `<w:cr/>` and a `Ctrl+Enter` page break packed into a run alongside its text.

## [0.11.7] - 2026-08-05

Line breaks now land where Word puts them, and a document containing an empty run
no longer fails to open at all.

### Fixed
- **Text no longer disappears around a page break.** Pressing `Ctrl+Enter` in the
  middle of a paragraph makes Word put the page break *inside* the run holding the
  surrounding text. That whole run was treated as the break, so the text on both
  sides of it vanished from the output. The paragraph now splits at the break with
  the text before and after it intact.
- **A line break sits between the two pieces of text it separates.** A `<w:br/>` or
  `<w:cr/>` between two text segments of one run was rendered after both of them,
  so `A<br>B` came out as `AB<br>`. Several consecutive breaks also collapsed into
  one; N breaks now render as N.
- **Empty runs no longer crash the parse.** A document containing `<w:r></w:r>` or
  `<w:r/>` — which Word emits routinely — threw a `TypeError` and failed the entire
  document. Empty runs are now read as the empty runs they are.
- **Headers, footers and footnotes keep their document order.** These parts never
  received the raw XML the parser uses to recover ordering, so all of the above
  applied to them, *and* a hyperlink in the middle of a header or footer paragraph
  was moved to the end of the line. Order is now recovered there too.

### Changed
- A run's children (text, tabs, breaks, symbols, hyphens) are read in true document
  order rather than inferred from their counts. Tabs packed into one run alongside
  text — the Google Docs `text <w:tab/> text` shape — no longer depend on the older
  positional heuristic, which stays only as a fallback. A run holding text followed
  by a tab now reports the tab after that text instead of in front of it.

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
