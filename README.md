# simple-super-doc

Browser-only TypeScript library that parses a `.docx` file into a typed
intermediate representation (IR) and renders it to HTML. Built for document
editors and viewers that need a faithful, inspectable model of the document.

- **Typed IR** — `parse()` returns a `DocxDocument` of `Block`s (paragraphs,
  tables) with computed styles, lists, images, and page size. Inspect or
  transform it before rendering.
- **HTML rendering** — `render()` paints the IR into a container element.
- **No DOM at parse time** — parsing is pure; only `render()` touches the DOM.
- **Security** — text is written with `textContent`, never `innerHTML`.

## Install

```bash
npm install simple-super-doc
```

## Usage

```ts
import { parse, render } from 'simple-super-doc'

const buf = await file.arrayBuffer()       // a .docx File/Blob
const doc = await parse(buf)               // typed DocxDocument IR
render(doc, document.getElementById('view')!)

// Optional: show tracked changes (deletions struck through, insertions
// underlined) instead of the accepted/final view.
render(doc, view, { showRevisions: true })
```

`parse` throws `DocxParseError` (`code: 'INVALID_ZIP' | 'MISSING_ENTRY'`) for
non-docx or malformed input.

## What it covers

- **Text & styles** — bold/italic/underline/strikethrough, super/subscript,
  font size/family, color, highlight and shading, paragraph alignment
  (including OOXML justified `both`), spacing, indentation and borders, and the
  full style cascade (docDefaults → named style → direct formatting).
- **Structure** — numbered and bulleted lists (nested), tables with
  `gridSpan`/`vMerge`, column widths, cell margins and cell borders (resolved by
  cascade: table style → `tblBorders` → `tcBorders`, so borders that come only
  from a table style like `TableGrid` are drawn), and block/run ordering
  recovered from the raw XML (paragraphs, tables, mid-paragraph hyperlinks and
  tracked changes keep their real sequence — in the body **and inside cells**).
- **Images** — inline and anchored, as base64 data URLs.
- **Headers & footers** — default `headerReference`/`footerReference`, rendered
  in the page margins on every page; distinct per section, with a section that
  declares none inheriting the previous section's (OOXML semantics).
- **Page breaks** — `w:pageBreakBefore` and explicit `<w:br w:type="page"/>`.
- **Tab stops** — right/center/decimal stops with dot/hyphen/underscore leaders
  (table-of-contents rows render as `Title …… 12`).
- **Fields** — `PAGE` and `NUMPAGES` resolve live (robust to `\* MERGEFORMAT`
  switches); other fields (DATE, REF, PAGEREF, TOC, HYPERLINK, …) render their
  cached result. Both the complex `fldChar` form and the compact `w:fldSimple`
  form are supported.
- **Footnotes & endnotes** — footnotes at the bottom of the referencing page;
  endnotes on a final page.
- **Tracked changes** — deletions and insertions, shown or hidden via the
  `showRevisions` render option.
- **Multiple sections** — per-section page size and orientation (e.g. a
  landscape table page between portrait pages), per-section headers/footers,
  with continuous page numbering. Each section is routed to the plain or
  full-page-background path by its own content, so template documents may mix
  page sizes across sections.

## Page-aware rendering

Some templates paginate themselves with full-page background images (a cover
slide, framed body pages, a closing slide) but carry **no explicit page
breaks** — no `<w:br w:type="page"/>` and no `<w:lastRenderedPageBreak/>`.
Word/LibreOffice decide the pages purely in their layout engine.

When `render()` detects a full-page background (a `behindDoc=1` anchor), it
reconstructs the pages with a two-pass approach and a set of heuristics. These
live in [`src/renderer/layout.ts`](src/renderer/layout.ts) as pure, unit-tested
functions ([`test/layout-heuristics.test.ts`](test/layout-heuristics.test.ts)):

- **Two-pass measurement** — each block is measured in a hidden container
  (Pass 1), then content is distributed into page-sized boxes (Pass 2). This
  needs a real layout engine, so it runs in the browser, not jsdom.
- **Empty paragraphs are spacers** — they collapse to 0px in a browser but
  occupy a line in Word, so empty paragraphs get a min line height
  (`EMPTY_LINE_EM`). Their vertical rhythm positions things like a cover's
  customer-name field.
- **Section headings force a new page** — a heading is large text
  (`fontSize >= 24`) or a wide-short "text-as-image" banner; a small icon that
  precedes a text heading travels with it. Each framed section starts a page.
- **Watermarks overlay, not flow** — large floating images (decorative
  frames/watermarks) are absolutely positioned behind the text and contribute
  zero pagination height, instead of pushing content onto extra pages.
- **Background regions by image order** — floating backgrounds are positioned
  by absolute page coordinates, so the *order* of distinct background images
  drives which page each covers (cover → page 0, body frame → pages 1+), not
  their XML position.
- **Full-bleed slides** — a page whose only content is a near-full-width image
  is drawn edge-to-edge; blank pages are dropped.

These thresholds are tuned for proposal/letter-style templates. A document
without a full-page background skips all of this and renders as a normal flow.

## Limitations & non-goals

This is a faithful but *simple* HTML renderer, not a Word layout engine. The
following are intentionally out of scope:

- **Floating layout** — text wrapping around floating images, multi-column
  layouts, shapes, charts, and SmartArt are not laid out. (Inline images and
  tables are supported.) Text inside a text box (`w:txbxContent`, DrawingML or
  VML) is recovered into the flow so it is never lost, but the box is not
  floated or positioned.
- **2D math layout** — OMML equations (`m:oMath`) are recovered as their linear
  text (e.g. `A=πr2`), inline and in order, so the content is never lost, but
  fractions, superscripts and radicals are not laid out two-dimensionally.
- **Pixel-exact pagination** — without Word's line-breaking and layout engine,
  page breaks are reconstructed by two-pass DOM measurement and heuristics.
  Pagination is close but not guaranteed to match Word/LibreOffice line for
  line. For a byte-faithful page image, convert the `.docx` to PDF.
- **Comments** — review comments are parsed away (treated as noise); only
  tracked-change insertions/deletions are surfaced (via `showRevisions`).

## Development

```bash
npm run typecheck   # tsc --noEmit
npm test            # vitest
npm run build       # esm + d.ts into dist/
npm run build:demo  # IIFE bundle into demo/ for the drag-and-drop demo
```

## License

MIT
